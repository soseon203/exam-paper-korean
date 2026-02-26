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
    }

    content_type = type_map.get(type_str)
    if content_type is None:
        content_type = ContentType.TEXT

    # 텍스트 블록에 $...$ 인라인 LaTeX가 있으면 분리
    if content_type == ContentType.TEXT and "$" in value:
        split = _split_inline_latex(value)
        if len(split) > 1:
            return split  # type: ignore[return-value]

    return ContentBlock(type=content_type, value=value)


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
