"""Claude OCR 응답을 구조화 데이터(ExamDocument)로 변환하는 파서."""

from __future__ import annotations

import re

from models.exam_document import (
    ContentBlock,
    ContentType,
    Choice,
    Question,
    ExamPage,
    ExamDocument,
)

# 텍스트 내 인라인 LaTeX $...$ 감지 패턴
_INLINE_LATEX_RE = re.compile(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)")


def parse_ocr_response(ocr_result: dict, page_number: int) -> ExamPage:
    """OCR 결과 dict를 ExamPage 객체로 변환.

    Args:
        ocr_result: Claude OCR 응답 JSON dict
        page_number: 페이지 번호 (1부터)

    Returns:
        ExamPage 객체
    """
    page = ExamPage(page_number=page_number)
    page.header_text = ocr_result.get("header", "")

    for q_data in ocr_result.get("questions", []):
        question = _parse_question(q_data)
        page.questions.append(question)

    return page


def _parse_question(q_data: dict) -> Question:
    """문제 dict를 Question 객체로 변환."""
    question = Question(
        number=q_data.get("number", 0),
        score=q_data.get("score"),
    )

    # 문제 본문
    for block_data in q_data.get("contents", []):
        result = _parse_content_block(block_data)
        if isinstance(result, list):
            question.contents.extend(result)
        elif result:
            question.contents.append(result)

    # 쉼표로 구분된 독립 수식 분리 (안전 폴백)
    question.contents = _split_comma_equations(question.contents)

    # 선택지
    for choice_data in q_data.get("choices", []):
        choice = _parse_choice(choice_data)
        if choice:
            question.choices.append(choice)

    # 소문항
    for sub_data in q_data.get("sub_questions", []):
        sub = _parse_question(sub_data)
        question.sub_questions.append(sub)

    return question


def _parse_choice(choice_data: dict) -> Choice | None:
    """선택지 dict를 Choice 객체로 변환."""
    number = choice_data.get("number", 0)
    if not number:
        return None

    choice = Choice(number=number)
    for block_data in choice_data.get("contents", []):
        result = _parse_content_block(block_data)
        if isinstance(result, list):
            choice.contents.extend(result)
        elif result:
            choice.contents.append(result)

    # 쉼표로 구분된 독립 수식 분리
    choice.contents = _split_comma_equations(choice.contents)

    return choice


def _parse_content_block(block_data: dict) -> ContentBlock | None:
    """콘텐츠 블록 dict를 ContentBlock 객체로 변환.

    텍스트 블록 안에 $...$ 인라인 LaTeX가 포함된 경우
    텍스트+수식으로 분리하여 리스트로 반환하므로,
    호출부에서 리스트 여부를 확인해야 합니다.
    """
    type_str = block_data.get("type", "")
    value = block_data.get("value", "")

    if not value and type_str != "image":
        return None

    type_map = {
        "text": ContentType.TEXT,
        "equation": ContentType.EQUATION,
        "equation_block": ContentType.EQUATION_BLOCK,
        "image": ContentType.IMAGE,
        "table": ContentType.TABLE,
    }

    content_type = type_map.get(type_str)
    if content_type is None:
        content_type = ContentType.TEXT

    # 표(table) 블록 처리
    if content_type == ContentType.TABLE:
        rows = block_data.get("rows", [])
        return ContentBlock(type=ContentType.TABLE, value=value, rows=rows)

    # 텍스트 블록에 __밑줄__ 마크업이 있으면 분리
    if content_type == ContentType.TEXT and "__" in value:
        split = _split_underline_markup(value)
        if len(split) > 1:
            return split  # type: ignore[return-value]

    # 텍스트 블록에 $...$ 인라인 LaTeX가 있으면 분리
    if content_type == ContentType.TEXT and "$" in value:
        split = _split_inline_latex(value)
        if len(split) > 1:
            return split  # type: ignore[return-value]

    # 텍스트 블록에 수식 패턴이 섞여 있으면 분리
    if content_type == ContentType.TEXT:
        split = _split_mixed_text_equation(value)
        if len(split) > 1:
            return split  # type: ignore[return-value]

    return ContentBlock(type=content_type, value=value)


def _split_comma_equations(blocks: list[ContentBlock]) -> list[ContentBlock]:
    """쉼표로 구분된 독립 수식을 개별 블록으로 분리.

    예: "A=2^6, B=3^6" → equation("A=2^6") + text(", ") + equation("B=3^6")
    괄호·중괄호 안의 쉼표는 분리하지 않습니다.
    """
    result: list[ContentBlock] = []
    for block in blocks:
        if block.type == ContentType.EQUATION and "," in block.value:
            parts = _split_at_top_level_commas(block.value)
            valid = [p.strip() for p in parts if p.strip()]
            if len(valid) > 1:
                for i, part in enumerate(valid):
                    if i > 0:
                        result.append(
                            ContentBlock(type=ContentType.TEXT, value=", ")
                        )
                    result.append(
                        ContentBlock(type=ContentType.EQUATION, value=part)
                    )
                continue
        result.append(block)
    return result


def _split_at_top_level_commas(s: str) -> list[str]:
    """최상위 레벨의 쉼표에서 분리 (괄호·중괄호 안 쉼표 무시)."""
    parts: list[str] = []
    depth = 0
    current: list[str] = []
    for ch in s:
        if ch in "({[":
            depth += 1
        elif ch in ")}]":
            depth = max(0, depth - 1)
        elif ch == "," and depth == 0:
            parts.append("".join(current))
            current = []
            continue
        current.append(ch)
    parts.append("".join(current))
    return parts


# 텍스트 안에서 수식 구간을 감지하는 패턴
# 영문 변수/숫자 + 수학 연산자(=, >, <, +, -, ×, ÷, ≤, ≥, ≠) 조합
# 예: "a > 0", "b", "x = 3", "2x + 1"
_MATH_EXPR_RE = re.compile(
    r'(?<![a-zA-Z])'          # 앞에 영문자 없음 (단어 중간 방지)
    r'('
    r'[a-zA-Z0-9]+'           # 시작: 변수/숫자
    r'(?:'
    r'\s*[=><+\-×÷≤≥≠^_]\s*'  # 수학 연산자
    r'[a-zA-Z0-9]+'           # 뒤따르는 변수/숫자
    r')*'
    r')'
    r'(?![a-zA-Z])'           # 뒤에 영문자 없음
)


def _split_mixed_text_equation(text: str) -> list[ContentBlock]:
    """텍스트 안에 섞인 수식 패턴(영문 변수, 부등호 등)을 분리.

    예: "(a > 0, b는 정수)에서"
    → text("(") + eq("a > 0") + text(", ") + eq("b") + text("는 정수)에서")
    """
    # 한글이 전혀 없으면 분리 불필요 (순수 텍스트거나 이미 수식)
    if not re.search(r'[\uac00-\ud7a3]', text):
        return [ContentBlock(type=ContentType.TEXT, value=text)]

    # 수식 후보가 없으면 분리 불필요
    if not re.search(r'[a-zA-Z]', text):
        return [ContentBlock(type=ContentType.TEXT, value=text)]

    blocks: list[ContentBlock] = []
    last_end = 0

    for m in _MATH_EXPR_RE.finditer(text):
        expr = m.group(1).strip()
        # 너무 긴 영문 단어는 수식이 아님 (예: "정수")
        if len(expr) > 20:
            continue
        # 한글이 포함된 매치는 건너뜀
        if re.search(r'[\uac00-\ud7a3]', expr):
            continue
        # 단독 숫자 1자리는 문맥에 따라 건너뜀 (문제번호 등)
        # 단, 연산자가 포함되어 있으면 수식으로 처리
        if expr.isdigit() and len(expr) == 1:
            continue

        before = text[last_end:m.start()]
        if before:
            blocks.append(ContentBlock(type=ContentType.TEXT, value=before))
        blocks.append(ContentBlock(type=ContentType.EQUATION, value=expr))
        last_end = m.end()

    after = text[last_end:]
    if after:
        blocks.append(ContentBlock(type=ContentType.TEXT, value=after))

    return blocks if len(blocks) > 1 else [ContentBlock(type=ContentType.TEXT, value=text)]


def _split_inline_latex(text: str) -> list[ContentBlock]:
    """텍스트에서 $...$ 인라인 LaTeX를 분리하여 ContentBlock 리스트로 반환."""
    blocks: list[ContentBlock] = []
    last_end = 0

    for m in _INLINE_LATEX_RE.finditer(text):
        # 수식 앞 텍스트
        before = text[last_end:m.start()]
        if before:
            blocks.append(ContentBlock(type=ContentType.TEXT, value=before))
        # 수식
        latex = m.group(1).strip()
        if latex:
            blocks.append(ContentBlock(type=ContentType.EQUATION, value=latex))
        last_end = m.end()

    # 마지막 텍스트
    after = text[last_end:]
    if after:
        blocks.append(ContentBlock(type=ContentType.TEXT, value=after))

    return blocks if blocks else [ContentBlock(type=ContentType.TEXT, value=text)]


# __밑줄__ 마크업 감지 패턴
_UNDERLINE_RE = re.compile(r"__(.+?)__")


def _split_underline_markup(text: str) -> list[ContentBlock]:
    """텍스트에서 __밑줄__ 마크업을 분리하여 ContentBlock 리스트로 반환.

    예: "옳지 __않은__ 것은?"
    → text("옳지 ") + text("않은", underline=True) + text(" 것은?")
    """
    blocks: list[ContentBlock] = []
    last_end = 0

    for m in _UNDERLINE_RE.finditer(text):
        before = text[last_end:m.start()]
        if before:
            blocks.append(ContentBlock(type=ContentType.TEXT, value=before))
        inner = m.group(1)
        if inner:
            blocks.append(
                ContentBlock(type=ContentType.TEXT, value=inner, underline=True)
            )
        last_end = m.end()

    after = text[last_end:]
    if after:
        blocks.append(ContentBlock(type=ContentType.TEXT, value=after))

    return blocks if blocks else [ContentBlock(type=ContentType.TEXT, value=text)]


def build_document(
    pages: list[ExamPage],
    title: str = "",
    subject: str = "",
    grade: str = "",
) -> ExamDocument:
    """ExamPage 리스트로 ExamDocument 생성."""
    doc = ExamDocument(title=title, subject=subject, grade=grade)
    doc.pages = pages

    # 헤더에서 제목/과목 자동 추출 시도
    if pages and not title:
        header = pages[0].header_text
        if header:
            doc.title = header

    return doc
