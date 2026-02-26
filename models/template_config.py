"""HWPX 템플릿 설정 모델.

템플릿 HWPX에서 추출한 페이지·스타일 설정을 담습니다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class TemplateConfig:
    """HWPX 템플릿에서 추출한 설정."""

    # 원본 템플릿 파일 경로
    source_path: Path

    # 페이지 설정
    page_width: Optional[int] = None    # HPF 단위
    page_height: Optional[int] = None
    landscape: bool = False

    # 여백 (HPF 단위)
    margin_top: Optional[int] = None
    margin_bottom: Optional[int] = None
    margin_left: Optional[int] = None
    margin_right: Optional[int] = None
    margin_header: Optional[int] = None
    margin_footer: Optional[int] = None

    # secPr XML 원본 (section0.xml의 첫 문단에서 추출)
    sec_pr_xml: Optional[bytes] = None

    # header.xml 원본 바이트 (글꼴, 스타일 정의 통째 보존)
    header_xml_bytes: Optional[bytes] = None

    # 검증 결과
    has_valid_section: bool = False
    has_valid_header: bool = False

    @property
    def is_valid(self) -> bool:
        """템플릿이 유효한지 확인."""
        return self.has_valid_section and self.has_valid_header

    @property
    def summary(self) -> str:
        """템플릿 정보 요약 문자열."""
        parts = [f"템플릿: {self.source_path.name}"]
        if self.page_width and self.page_height:
            parts.append(f"페이지: {self.page_width}x{self.page_height}")
        if self.landscape:
            parts.append("가로 방향")
        return " | ".join(parts)
