"""HWPX 생성 전 미리보기 다이얼로그.

페이지별 원본 이미지 썸네일, 인식 결과 요약, 품질 경고를 표시하고
사용자가 변환 진행/취소를 결정할 수 있게 합니다.
"""

from __future__ import annotations

from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.quality_checker import ImageQuality
from core.ocr_engine import OCRQuality
from models.exam_document import ExamPage


class PageInfo:
    """한 페이지의 미리보기 정보를 모아두는 컨테이너."""

    def __init__(
        self,
        page_number: int,
        image: Image.Image,
        exam_page: ExamPage,
        image_quality: ImageQuality,
        ocr_quality: OCRQuality,
    ):
        self.page_number = page_number
        self.image = image
        self.exam_page = exam_page
        self.image_quality = image_quality
        self.ocr_quality = ocr_quality


class PreviewDialog(QDialog):
    """변환 미리보기 다이얼로그."""

    def __init__(self, pages: list[PageInfo], parent=None):
        super().__init__(parent)
        self.pages = pages
        self._accepted = False
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("변환 미리보기")
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout(self)

        # ── 요약 헤더 ──
        total_questions = sum(len(p.exam_page.questions) for p in self.pages)
        total_equations = sum(p.ocr_quality.equation_count for p in self.pages)
        total_warnings = sum(
            len(p.image_quality.warnings) + len(p.ocr_quality.warnings)
            for p in self.pages
        )

        summary_text = (
            f"총 {len(self.pages)}페이지 | "
            f"문제 {total_questions}개 | "
            f"수식 {total_equations}개"
        )
        if total_warnings > 0:
            summary_text += f" | 경고 {total_warnings}건"

        summary_label = QLabel(summary_text)
        summary_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 8px;")
        layout.addWidget(summary_label)

        # ── 페이지 목록 (스크롤) ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        for page_info in self.pages:
            page_widget = self._create_page_widget(page_info)
            scroll_layout.addWidget(page_widget)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        # ── 버튼 ──
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("취소")
        cancel_btn.setFixedSize(120, 40)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        proceed_btn = QPushButton("변환 진행")
        proceed_btn.setFixedSize(120, 40)
        proceed_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: #0d6efd;"
            "  color: white;"
            "  border: none;"
            "  border-radius: 6px;"
            "  font-size: 14px;"
            "  font-weight: bold;"
            "}"
            "QPushButton:hover { background-color: #0b5ed7; }"
        )
        proceed_btn.clicked.connect(self.accept)
        btn_layout.addWidget(proceed_btn)

        layout.addLayout(btn_layout)

    def _create_page_widget(self, info: PageInfo) -> QWidget:
        """한 페이지의 미리보기 위젯 생성."""
        widget = QWidget()
        widget.setStyleSheet(
            "QWidget { border: 1px solid #ddd; border-radius: 6px; "
            "padding: 8px; margin: 4px; background: #fafafa; }"
        )
        h_layout = QHBoxLayout(widget)

        # 썸네일
        thumb_label = QLabel()
        pixmap = self._pil_to_pixmap(info.image, max_height=160)
        thumb_label.setPixmap(pixmap)
        thumb_label.setFixedSize(pixmap.width(), pixmap.height())
        h_layout.addWidget(thumb_label)

        # 정보 영역
        info_layout = QVBoxLayout()

        # 페이지 번호 + 문제 수
        q_count = len(info.exam_page.questions)
        eq_count = info.ocr_quality.equation_count
        title = QLabel(
            f"<b>페이지 {info.page_number}</b> — "
            f"문제 {q_count}개, 수식 {eq_count}개"
        )
        title.setStyleSheet("font-size: 13px; border: none;")
        info_layout.addWidget(title)

        # 품질 점수
        score = info.image_quality.score
        color = "#198754" if info.image_quality.passed else "#dc3545"
        score_label = QLabel(
            f'품질 점수: <span style="color:{color}; font-weight:bold;">'
            f"{score:.0f}/100</span>"
        )
        score_label.setStyleSheet("font-size: 12px; border: none;")
        info_layout.addWidget(score_label)

        # 경고 목록
        all_warnings = info.image_quality.warnings + info.ocr_quality.warnings
        if all_warnings:
            for warn in all_warnings:
                warn_label = QLabel(f"⚠ {warn}")
                warn_label.setStyleSheet(
                    "color: #856404; background: #fff3cd; "
                    "border-radius: 3px; padding: 2px 6px; "
                    "font-size: 11px; border: none;"
                )
                warn_label.setWordWrap(True)
                info_layout.addWidget(warn_label)
        else:
            ok_label = QLabel("품질 양호")
            ok_label.setStyleSheet(
                "color: #0f5132; background: #d1e7dd; "
                "border-radius: 3px; padding: 2px 6px; "
                "font-size: 11px; border: none;"
            )
            info_layout.addWidget(ok_label)

        info_layout.addStretch()
        h_layout.addLayout(info_layout, stretch=1)
        return widget

    @staticmethod
    def _pil_to_pixmap(image: Image.Image, max_height: int = 160) -> QPixmap:
        """PIL Image → QPixmap (썸네일 크기로 축소)."""
        img = image.convert("RGB")
        # 축소 비율 계산
        w, h = img.size
        if h > max_height:
            ratio = max_height / h
            img = img.resize((int(w * ratio), max_height), Image.LANCZOS)

        data = img.tobytes("raw", "RGB")
        qimg = QImage(data, img.width, img.height, 3 * img.width, QImage.Format.Format_RGB888)
        return QPixmap.fromImage(qimg)
