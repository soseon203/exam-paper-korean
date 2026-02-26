from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class ContentType(Enum):
    """콘텐츠 블록 유형."""
    TEXT = "text"
    EQUATION = "equation"          # 인라인 수식
    EQUATION_BLOCK = "equation_block"  # 블록(독립행) 수식
    IMAGE = "image"


@dataclass
class ContentBlock:
    """문서 내 개별 콘텐츠 블록."""
    type: ContentType
    value: str  # TEXT: 텍스트, EQUATION/EQUATION_BLOCK: LaTeX, IMAGE: 파일경로
    hwp_equation: Optional[str] = None  # 변환된 HWP 수식 스크립트

    @property
    def is_equation(self) -> bool:
        return self.type in (ContentType.EQUATION, ContentType.EQUATION_BLOCK)


@dataclass
class Choice:
    """선택지 (보기)."""
    number: int       # 1~5
    contents: list[ContentBlock] = field(default_factory=list)


@dataclass
class Question:
    """시험 문제 하나."""
    number: int
    score: Optional[int] = None
    contents: list[ContentBlock] = field(default_factory=list)   # 문제 본문
    choices: list[Choice] = field(default_factory=list)          # 선택지
    sub_questions: list[Question] = field(default_factory=list)  # 소문항


@dataclass
class ExamPage:
    """시험지 한 페이지."""
    page_number: int
    questions: list[Question] = field(default_factory=list)
    header_text: str = ""   # 페이지 상단 (과목명, 학년 등)


@dataclass
class ExamDocument:
    """전체 시험 문서."""
    title: str = ""
    subject: str = ""
    grade: str = ""
    pages: list[ExamPage] = field(default_factory=list)

    @property
    def all_questions(self) -> list[Question]:
        """모든 페이지의 문제 목록."""
        questions = []
        for page in self.pages:
            questions.extend(page.questions)
        return questions
