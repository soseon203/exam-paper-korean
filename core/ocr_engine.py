"""Claude Vision API를 이용한 수학 시험지 OCR 엔진."""

from __future__ import annotations

import json
import logging
import re
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

## 핵심 원칙 (가장 중요!)
1. 이미지의 텍스트를 **한 글자씩 정확하게** 읽으세요. 추측·의역·요약 금지.
2. 원본 문장을 그대로 복사하듯이 적으세요. 단어를 바꾸거나 빼지 마세요.
3. 수식도 이미지에 보이는 그대로 추출. 숫자·계수·지수를 절대 바꾸지 마세요.
4. 한글 음절을 하나라도 빠뜨리거나 바꾸면 안 됩니다:
   - "거듭제곱" ≠ "기하적금" (X)  /  "옳은" ≠ "올은" (X)
   - "거실" ≠ "가설" (X)  /  "회전축" ≠ "위중" (X)
   - "알맞은" ≠ "오는" (X)  /  "민성이는" ≠ "기여는" (X)

## 문제 번호 규칙 (매우 중요!)
- 문제의 **주 번호**(1., 2., ... 20., 21.)를 반드시 number에 기록하세요.
- "19. [서술형 3]"이면 number는 **19**입니다 (3이 아닙니다!).
- "20. [서술형 4]"이면 number는 **20**입니다 (4가 아닙니다!).
- [서술형 N] 레이블은 contents의 텍스트에 포함하세요:
  {"type":"text","value":"[서술형 3] 아래 그림은..."}
- 페이지 상단의 학교명·과목명은 header에만 넣고, questions에 넣지 마세요.

## 출력 규칙
1. **JSON 형식으로만** 응답. 순수 JSON만 출력하세요.
2. 수식은 **LaTeX 형식**으로 변환. 인라인 수식은 type="equation", 독립행은 type="equation_block".
3. 선택지 번호는 ①②③④⑤를 1,2,3,4,5로 변환.
4. 배점은 score에 숫자만 (예: 3, 4, 5, 8).

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
        {"number": 1, "contents": [{"type": "equation", "value": "\\\\frac{5}{6}"}]}
      ],
      "sub_questions": []
    }
  ]
}
```

## 수식/텍스트 분리 규칙
- **equation**: 영문 변수, 숫자, 수학 기호(+, -, =, ×, ÷), 분수, 지수 등 순수 수학 표현
- **text**: 한글 텍스트, 한글 괄호 내용("(가)", "(나)"), 조사, 문장부호
- 수식+한글이 섞인 문장은 반드시 분리:
  올바른 예: {"type":"equation","value":"a > 0"}, {"type":"text","value":"이고 "}, {"type":"equation","value":"b"}, {"type":"text","value":"는 정수일 때"}
- 쉼표로 구분된 독립 수식은 개별 블록으로:
  {"type":"equation","value":"A=2^6"}, {"type":"text","value":", "}, {"type":"equation","value":"B=3^6"}
- □(빈칸) → {"type":"equation","value":"\\\\square"}

## 괄호 종류 구분 (매우 중요!)
- 소괄호 ( ), 중괄호 \\{ \\}, 대괄호 [ ]를 정확히 구분하세요.
- 중첩 괄호 문제에서 괄호 종류가 다른 것은 의도적입니다:
  - 올바른 예: 6x - [3y + 2x - \\{3x + \\square - (5x - 7y)\\}] (O)
  - 잘못된 예: 6x - (3y + 2x - (3x + \\square - (5x - 7y))) (X) — 모두 ()로 바꾸면 안 됩니다!

## 변수·기호 정확도 (매우 중요!)
- **x와 z를 혼동하지 마세요.** 같은 수식에 x가 있으면, 다른 곳의 같은 글자도 x입니다.
- **÷(나눗셈)와 +(덧셈)을 혼동하지 마세요.** ÷는 가로줄 위아래에 점이 있습니다.
- ≠ (\neq), ≤ (\leq), ≥ (\geq), < (\lt), > (\gt)를 정확히 구분.
- 여러 변수(x, y, z)가 있는 수식에서 **변수를 누락하지 마세요**:
  - (x^a y^b z^c)^d에서 z^c를 빠뜨리면 안 됩니다!

## 지수(위첨자) 정확도 (매우 중요!)
- 지수는 글자가 작아서 오인식이 빈번합니다. 확대해서 확인하세요.
- **한 자릿수와 두 자릿수를 혼동하면 안 됩니다**: 2^{48} ≠ 2^{6}, x^{15} ≠ x^{5}
- 같은 문제의 여러 선택지에서 지수가 모두 같으면 오인식일 가능성이 높습니다.
- 각 선택지의 수식이 서로 **달라야** 합니다. 동일하면 오인식입니다.

## 순환소수
- 순환마디(점)는 LaTeX \\dot{}으로: 0.\\dot{2}\\dot{4} (24 순환)
- "순환소수"라는 단어가 나오면 소수에 반드시 순환마디 점이 있습니다.

## 조건 박스·표 (매우 중요!)
- 테두리/박스 안의 내용(조건, 정의 등)은 **절대 누락하지 마세요.**
- 박스 안에 (가), (나) 등이 있으면 sub_questions로 처리.
- 표(격자/그리드)가 있으면 type="table"로 추출:
```json
{"type": "table", "value": "", "rows": [
  ["열1", "열2", "열3"],
  ["값1", "값2", "값3"],
  ["값4", "값5", "값6"]
]}
```
- 표 안의 수식은 LaTeX로 변환하여 셀에 넣으세요.

## 강조 표시
- 밑줄 강조 텍스트: __텍스트__ 형식으로 감싸세요.
  예: {"type":"text","value":"옳지 __않은__ 것은?"}

## 한글 정확도
- 한글의 **모든 음절을 빠짐없이** 추출. 글자를 누락·치환하면 안 됩니다.
- 조사(은/는/이/가/을/를/의)와 접미사(들, 째, 개)를 빠뜨리지 마세요.
- 숫자·분수를 한글로 오인식하지 마세요: "1.1" ≠ "기", \\frac{1}{5} ≠ "다"

## 서술형 문제
- 서술형은 choices를 빈 배열로.
- "(단, 풀이 과정을 반드시 적으시오.)" 등의 부가 지시문도 빠짐없이 추출.
- 그림/도형 설명이 있으면 텍스트로 추출 (이미지 자체는 type="image").

## 주의사항
- 배점이 있으면 score에 숫자로 기록.
- 이미지에 문제가 여러 개 있으면 모두 추출.
- **이미지의 모든 텍스트를 빠짐없이 추출하세요.**
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
        try:
            return self._extract_json(response_text)
        except json.JSONDecodeError as e:
            # JSON 파싱 완전 실패 시 1회 재시도
            logger.warning("JSON 파싱 실패 (1차), 재시도: %s", e)
            message2 = self.client.messages.create(
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
            return self._extract_json(message2.content[0].text)

    def _extract_json(self, text: str) -> dict:
        """응답에서 JSON 추출 (LaTeX 수식이 포함된 경우도 처리).

        LLM이 반환하는 JSON은 LaTeX 역슬래시, 이스케이프 누락 등으로
        파싱 실패가 빈번합니다. 여러 단계의 복구를 시도합니다.
        """

        # ── 1단계: JSON 블록 추출 ──
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

        # ── 1.5단계: LaTeX-JSON 이스케이프 충돌 방지 ──
        # \f(rac), \b(eta), \n(eq), \r(ight), \t(imes) 등은
        # JSON 이스케이프(\f=form-feed, \b=backspace 등)와 충돌하므로
        # LaTeX 명령어인 경우(\+알파벳 연속) 이중 이스케이프로 보호
        text = re.sub(r'(?<!\\)\\([bfnrt])(?=[a-zA-Z])', r'\\\\\1', text)

        # ── 2단계: 직접 파싱 ──
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # ── 3단계: 문자열 내부 역슬래시 이스케이프 복구 ──
        fixed = self._fix_json_backslashes(text)
        try:
            return json.loads(fixed)
        except json.JSONDecodeError:
            pass

        # ── 4단계: 문자열 값을 보호하면서 구조 복구 ──
        logger.warning("JSON 파싱 실패, 문자열 보호 복구 시도")
        repaired = self._repair_json_strings(text)
        try:
            return json.loads(repaired)
        except json.JSONDecodeError:
            pass

        # ── 5단계: 줄 단위 복구 (최후 수단) ──
        logger.warning("JSON 파싱 재실패, 줄 단위 복구 시도")
        lines = text.split("\n")
        for i, line in enumerate(lines):
            # "value" 필드에서 이스케이프 안 된 큰따옴표 수정
            match = re.match(r'^(\s*"value"\s*:\s*")(.*)(")(.*)$', line)
            if match:
                inner = match.group(2)
                inner = inner.replace('\\"', '\x00')
                inner = inner.replace('"', '\\"')
                inner = inner.replace('\x00', '\\"')
                lines[i] = match.group(1) + inner + match.group(3) + match.group(4)

        text = "\n".join(lines)
        text = re.sub(r",\s*([}\]])", r"\1", text)
        return json.loads(text)

    @staticmethod
    def _fix_json_backslashes(text: str) -> str:
        """JSON 문자열 내부의 이스케이프 안 된 역슬래시를 수정 (1차 시도).

        전체 텍스트에서 \\X (X가 유효 JSON 이스케이프가 아닌 것)를 \\\\X로 변환.
        JSON 구조 바깥에는 역슬래시가 없으므로 전체 텍스트에 적용해도 안전.
        """
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        # \u는 유지 (\uXXXX 유니코드 이스케이프일 수 있으므로 — 4단계에서 정밀 처리)
        return re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', text)

    @staticmethod
    def _repair_json_strings(text: str) -> str:
        """JSON 문자열 값 내부의 깨진 부분을 복구.

        문자열 리터럴을 하나씩 추출하면서:
        - 유효하지 않은 이스케이프(\\frac, \\{ 등)를 이중 이스케이프
        - \\u + 비16진수를 이중 이스케이프 (\\underset 등 LaTeX)
        - 문자열 내 줄바꿈/탭을 이스케이프
        - 내부 따옴표를 감지하여 이스케이프
        """
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

        result = []
        i = 0
        while i < len(text):
            if text[i] == '"':
                j = i + 1
                parts = []
                while j < len(text):
                    ch = text[j]
                    if ch == '\\' and j + 1 < len(text):
                        nc = text[j + 1]
                        if nc == 'u':
                            # \uXXXX 유니코드 이스케이프 검증
                            hex_part = text[j + 2:j + 6]
                            if (len(hex_part) == 4
                                    and all(c in '0123456789abcdefABCDEF'
                                            for c in hex_part)):
                                parts.append('\\u')
                                parts.append(hex_part)
                                j += 6
                            else:
                                # \underset 등 LaTeX → 이중 이스케이프
                                parts.append('\\\\u')
                                j += 2
                        elif nc in '"\\/bfnrt':
                            parts.append(ch)
                            parts.append(nc)
                            j += 2
                        else:
                            # \frac, \{ 등 → 이중 이스케이프
                            parts.append('\\\\')
                            parts.append(nc)
                            j += 2
                        continue
                    if ch == '"':
                        # 진짜 문자열 끝인지 확인
                        k = j + 1
                        while k < len(text) and text[k] in ' \t\r\n':
                            k += 1
                        if k >= len(text) or text[k] in ',}]:"':
                            break  # 구조 문자 → 진짜 끝
                        # 내부 따옴표 → 이스케이프
                        parts.append('\\"')
                        j += 1
                        continue
                    if ch == '\n':
                        parts.append('\\n')
                        j += 1
                        continue
                    if ch == '\r':
                        j += 1
                        continue
                    parts.append(ch)
                    j += 1

                result.append('"')
                result.append(''.join(parts))
                result.append('"')
                i = j + 1
            else:
                result.append(text[i])
                i += 1

        return ''.join(result)

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
