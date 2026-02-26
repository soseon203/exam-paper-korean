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

    _BTN_HEIGHT = 36

    def _setup_ui(self):
        self.setWindowTitle("변환 미리보기")
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(24, 20, 24, 20)

        # ── 요약 헤더 ──
        total_questions = sum(len(p.exam_page.questions) for p in self.pages)
        total_equations = sum(p.ocr_quality.equation_count for p in self.pages)
        total_warnings = sum(
            len(p.image_quality.warnings) + len(p.ocr_quality.warnings)
            for p in self.pages
        )

        header_label = QLabel("변환 결과 미리보기")
        header_label.setStyleSheet(
            "font-size: 16px; font-weight: 600; color: #1d2939; padding: 0; margin: 0;"
        )
        layout.addWidget(header_label)
        layout.addSpacing(8)

        # 요약 뱃지들
        badges = []
        badges.append(
            f'<span style="background:#eff8ff; color:#1570ef; '
            f'padding:3px 10px; border-radius:10px; font-size:12px;">'
            f'{len(self.pages)}페이지</span>'
        )
        badges.append(
            f'<span style="background:#f0f9ff; color:#026aa2; '
            f'padding:3px 10px; border-radius:10px; font-size:12px;">'
            f'문제 {total_questions}개</span>'
        )
        badges.append(
            f'<span style="background:#f0f9ff; color:#026aa2; '
            f'padding:3px 10px; border-radius:10px; font-size:12px;">'
            f'수식 {total_equations}개</span>'
        )
        if total_warnings > 0:
            badges.append(
                f'<span style="background:#fffaeb; color:#b54708; '
                f'padding:3px 10px; border-radius:10px; font-size:12px;">'
                f'경고 {total_warnings}건</span>'
            )

        summary_label = QLabel("  ".join(badges))
        summary_label.setTextFormat(Qt.TextFormat.RichText)
        summary_label.setStyleSheet("padding: 0; margin: 0;")
        layout.addWidget(summary_label)

        layout.addSpacing(16)

        # ── 페이지 목록 (스크롤) ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(
            "QScrollArea { border: 1px solid #e4e7ec; border-radius: 8px; }"
        )
        scroll_widget = QWidget()
        scroll_widget.setStyleSheet("background: #f9fafb;")
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setContentsMargins(12, 12, 12, 12)
        scroll_layout.setSpacing(8)

        for page_info in self.pages:
            page_widget = self._create_page_widget(page_info)
            scroll_layout.addWidget(page_widget)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        layout.addSpacing(16)

        # ── 버튼 ──
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        cancel_btn = QPushButton("취소")
        cancel_btn.setFixedSize(80, self._BTN_HEIGHT)
        cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        proceed_btn = QPushButton("변환 진행")
        proceed_btn.setFixedSize(120, self._BTN_HEIGHT)
        proceed_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        proceed_btn.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "    stop:0 #1570ef, stop:1 #0d6efd);"
            "  color: white;"
            "  border: none;"
            "  border-radius: 6px;"
            "  font-size: 13px;"
            "  font-weight: 600;"
            "}"
            "QPushButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "    stop:0 #0d6efd, stop:1 #0b5ed7);"
            "}"
        )
        proceed_btn.clicked.connect(self.accept)
        btn_layout.addWidget(proceed_btn)

        layout.addLayout(btn_layout)

    def _create_page_widget(self, info: PageInfo) -> QWidget:
        """한 페이지의 미리보기 위젯 생성."""
        widget = QWidget()
        widget.setStyleSheet(
            "QWidget#pageCard {"
            "  border: 1px solid #e4e7ec;"
            "  border-radius: 8px;"
            "  padding: 12px;"
            "  background: #ffffff;"
            "}"
        )
        widget.setObjectName("pageCard")
        h_layout = QHBoxLayout(widget)
        h_layout.setSpacing(16)

        # 썸네일
        thumb_label = QLabel()
        thumb_label.setObjectName("thumb")
        pixmap = self._pil_to_pixmap(info.image, max_height=140)
        thumb_label.setPixmap(pixmap)
        thumb_label.setFixedSize(pixmap.width(), pixmap.height())
        thumb_label.setStyleSheet(
            "QLabel#thumb {"
            "  border: 1px solid #e4e7ec;"
            "  border-radius: 4px;"
            "  background: #ffffff;"
            "}"
        )
        h_layout.addWidget(thumb_label)

        # 정보 영역
        info_layout = QVBoxLayout()
        info_layout.setSpacing(6)

        # 페이지 번호 + 문제 수
        q_count = len(info.exam_page.questions)
        eq_count = info.ocr_quality.equation_count
        title = QLabel(
            f"<b>페이지 {info.page_number}</b>"
            f'<span style="color:#667085;"> — '
            f"문제 {q_count}개, 수식 {eq_count}개</span>"
        )
        title.setStyleSheet("font-size: 13px; border: none; color: #1d2939;")
        info_layout.addWidget(title)

        # 품질 점수
        score = info.image_quality.score
        if info.image_quality.passed:
            score_color = "#027a48"
            score_bg = "#ecfdf3"
        else:
            score_color = "#b42318"
            score_bg = "#fef3f2"
        score_label = QLabel(
            f'<span style="color:{score_color}; font-weight:600;">'
            f"{score:.0f}/100</span>"
        )
        score_label.setStyleSheet(
            f"font-size: 12px; border: none; background: {score_bg};"
            f"border-radius: 10px; padding: 2px 10px;"
        )
        score_label.setFixedWidth(70)
        info_layout.addWidget(score_label)

        # 경고 목록
        all_warnings = info.image_quality.warnings + info.ocr_quality.warnings
        if all_warnings:
            for warn in all_warnings:
                warn_label = QLabel(f"  {warn}")
                warn_label.setStyleSheet(
                    "color: #b54708; background: #fffaeb; "
                    "border-radius: 4px; padding: 3px 8px; "
                    "font-size: 11px; border: none;"
                )
                warn_label.setWordWrap(True)
                info_layout.addWidget(warn_label)
        else:
            ok_label = QLabel("  품질 양호")
            ok_label.setStyleSheet(
                "color: #027a48; background: #ecfdf3; "
                "border-radius: 4px; padding: 3px 8px; "
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
