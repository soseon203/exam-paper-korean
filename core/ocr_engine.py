"""Claude Vision API를 이용한 수학 시험지 OCR 엔진."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

from PIL import Image
import anthropic

logger = logging.getLogger(__name__)

from core.pdf_handler import image_to_base64
from utils.config import get_api_key, CLAUDE_MODEL, CLAUDE_MAX_TOKENS


@dataclass
class OCRQuality:
    """OCR 응답 검증 결과."""

    valid: bool = True
    warnings: list[str] = field(default_factory=list)
    question_count: int = 0
    equation_count: int = 0

# 한국어 수학 시험지 전용 OCR 프롬프트
EXAM_OCR_PROMPT = """당신은 한국 수학 시험지를 정밀하게 OCR하는 전문가입니다.
이미지에서 모든 텍스트와 수식을 정확하게 추출하세요.

## 출력 규칙

1. **JSON 형식으로만** 응답하세요. 다른 텍스트 없이 순수 JSON만 출력하세요.
2. 수식은 반드시 **LaTeX 형식**으로 변환하세요.
3. 인라인 수식은 `$...$` 표기 없이, type을 "equation"으로 지정하세요.
4. 독립행 수식(별도 줄에 표시된 수식)은 type을 "equation_block"으로 지정하세요.
5. 문제 번호, 배점, 선택지를 정확하게 구분하세요.
6. 선택지 번호는 ①②③④⑤ 를 1,2,3,4,5로 변환하세요.

## 출력 JSON 구조

```json
{
  "header": "페이지 상단 텍스트 (과목명, 학년 등)",
  "questions": [
    {
      "number": 1,
      "score": 3,
      "contents": [
        {"type": "text", "value": "다음 식의 값을 구하시오."},
        {"type": "equation_block", "value": "\\\\frac{1}{2} + \\\\frac{1}{3}"}
      ],
      "choices": [
        {
          "number": 1,
          "contents": [
            {"type": "text", "value": ""},
            {"type": "equation", "value": "\\\\frac{5}{6}"}
          ]
        }
      ],
      "sub_questions": []
    }
  ]
}
```

## 수식 범위 규칙 (매우 중요!)
- **equation으로 처리할 것**: 영문 변수(a, b, x, y), 숫자(1, 2, 3), 수학 기호(+, -, =, ×, ÷), 분수, 지수, 루트, 적분, 시그마 등 순수 수학 표현만
- **text로 처리할 것**: 한글 텍스트, 괄호와 그 안의 한글 (예: "(는 경우)", "(가)", "(나)"), 조사, 문장부호
- 하나의 문장 안에 수식과 텍스트가 섞여 있으면 **반드시 분리**하세요.
  - 올바른 예: {"type":"equation","value":"a = 0"}, {"type":"text","value":"이고 "}, {"type":"equation","value":"b"}, {"type":"text","value":"는 경우의 꼴로 나타낼 수 있다."}
  - 잘못된 예: {"type":"equation","value":"a = 0, a, b는 경우의 꼴로"}
- 모든 숫자는 수식으로 처리하세요. 예: "2" → {"type":"equation","value":"2"}
- 단, 문제 번호(1., 2.)만 text로 처리합니다. 배점 숫자도 수식입니다.
- 쉼표(,)로 구분된 여러 독립 수식은 개별 equation 블록 + 텍스트 쉼표로 분리하세요.
  - 올바른 예: {"type":"equation","value":"A=2^6"}, {"type":"text","value":", "}, {"type":"equation","value":"B=3^6"}
  - 잘못된 예: {"type":"equation","value":"A=2^6, B=3^6, C=5^4"}
- 순환소수 순환마디(점)는 LaTeX \\dot{}으로 표현하세요. 예: 0.\\dot{1}3\\dot{6}
- □(빈칸) 기호가 있으면 수식으로: {"type":"equation","value":"\\square"}

## 주의사항
- 한글 텍스트는 정확하게 보존하세요.
- 수식 기호를 놓치지 마세요: 분수, 지수, 루트, 적분, 시그마 등
- 배점이 표시되어 있으면 score에 숫자로 기록하세요.
- 이미지에 문제가 여러 개 있으면 모두 추출하세요.
- 선택지가 없는 주관식 문제는 choices를 빈 배열로 두세요.
- sub_questions는 (가), (나) 등의 소문항에 사용하세요.
"""


class OCREngine:
    """Claude Vision API 기반 OCR 엔진."""

    def __init__(self, api_key: str | None = None):
        self.api_key = api_key or get_api_key()
        self.client = anthropic.Anthropic(api_key=self.api_key)

    def recognize_page(self, image: Image.Image) -> dict:
        """한 페이지 이미지에서 텍스트+수식 추출.

        Args:
            image: 페이지 이미지 (PIL Image)

        Returns:
            구조화된 OCR 결과 dict
        """
        base64_image = image_to_base64(image, format="PNG")

        message = self.client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=CLAUDE_MAX_TOKENS,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": base64_image,
                            },
                        },
                        {
                            "type": "text",
                            "text": EXAM_OCR_PROMPT,
                        },
                    ],
                }
            ],
        )

        response_text = message.content[0].text
        return self._extract_json(response_text)

    def _extract_json(self, text: str) -> dict:
        """응답에서 JSON 추출 (LaTeX 수식이 포함된 경우도 처리)."""
        import json
        import re

        # JSON 블록 추출 시도
        text = text.strip()

        # ```json ... ``` 블록 처리
        if "```json" in text:
            start = text.index("```json") + 7
            end = text.index("```", start)
            text = text[start:end].strip()
        elif "```" in text:
            start = text.index("```") + 3
            end = text.index("```", start)
            text = text[start:end].strip()

        # { 로 시작하는 JSON 찾기
        if not text.startswith("{"):
            brace_start = text.find("{")
            if brace_start != -1:
                text = text[brace_start:]

        # trailing comma 제거
        text = re.sub(r",\s*([}\]])", r"\1", text)

        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # 복구 시도: value 내부의 이스케이프 안 된 역슬래시 처리
        # JSON 문자열 안의 \를 \\로 이스케이프 (이미 이스케이프된 것 제외)
        fixed = re.sub(
            r'(?<=: ")(.*?)(?=")',
            lambda m: m.group(0).replace("\\", "\\\\")
                .replace("\\\\n", "\\n")
                .replace("\\\\t", "\\t")
                .replace('\\\\"', '\\"')
                .replace("\\\\\\\\", "\\\\"),
            text,
            flags=re.DOTALL,
        )
        try:
            return json.loads(fixed)
        except json.JSONDecodeError:
            pass

        # 최종 시도: 줄 단위로 문제 위치 파악 후 수동 수정
        logger.warning("JSON 파싱 실패, 줄 단위 복구 시도")
        lines = text.split("\n")
        for i, line in enumerate(lines):
            # "value" 필드에서 이스케이프 안 된 큰따옴표 수정
            # "value": "some "broken" text" 패턴
            match = re.match(r'^(\s*"value"\s*:\s*")(.*)(")(.*)$', line)
            if match:
                inner = match.group(2)
                # 내부 큰따옴표 이스케이프
                inner = inner.replace('\\"', '\x00')
                inner = inner.replace('"', '\\"')
                inner = inner.replace('\x00', '\\"')
                lines[i] = match.group(1) + inner + match.group(3) + match.group(4)

        text = "\n".join(lines)
        text = re.sub(r",\s*([}\]])", r"\1", text)
        return json.loads(text)

    def validate_api_key(self) -> bool:
        """API 키 유효성 검사."""
        try:
            self.client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=10,
                messages=[{"role": "user", "content": "test"}],
            )
            return True
        except anthropic.AuthenticationError:
            return False
        except Exception:
            return True  # 인증 외 오류는 키 자체는 유효할 수 있음


def validate_ocr_response(ocr_result: dict) -> OCRQuality:
    """OCR 응답의 유효성을 검증.

    Args:
        ocr_result: OCR 엔진이 반환한 dict

    Returns:
        OCRQuality 검증 결과
    """
    quality = OCRQuality()

    # 1) 기본 구조 검증
    if not isinstance(ocr_result, dict):
        quality.valid = False
        quality.warnings.append("OCR 응답이 dict가 아닙니다.")
        return quality

    questions = ocr_result.get("questions", [])
    quality.question_count = len(questions)

    # 2) 문제 0개 → 경고
    if quality.question_count == 0:
        quality.warnings.append("인식된 문제가 없습니다. 이미지를 확인하세요.")

    # 3) 수식 개수 집계 + LaTeX 괄호 짝 검증
    for q in questions:
        _validate_question_latex(q, quality)

    return quality


def _validate_question_latex(q_data: dict, quality: OCRQuality) -> None:
    """문제 내 LaTeX 수식의 괄호 짝을 검증."""
    for block in q_data.get("contents", []):
        if block.get("type") in ("equation", "equation_block"):
            quality.equation_count += 1
            _check_latex_brackets(block.get("value", ""), quality)

    for choice in q_data.get("choices", []):
        for block in choice.get("contents", []):
            if block.get("type") in ("equation", "equation_block"):
                quality.equation_count += 1
                _check_latex_brackets(block.get("value", ""), quality)

    for sub in q_data.get("sub_questions", []):
        _validate_question_latex(sub, quality)


def _check_latex_brackets(latex: str, quality: OCRQuality) -> None:
    """LaTeX 문자열의 중괄호 짝이 맞는지 확인."""
    depth = 0
    for ch in latex:
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
        if depth < 0:
            quality.warnings.append(
                f"LaTeX 괄호 불일치 (닫는 괄호 초과): {latex[:50]}..."
            )
            return
    if depth != 0:
        quality.warnings.append(
            f"LaTeX 괄호 불일치 (여는 괄호 {depth}개 초과): {latex[:50]}..."
        )
