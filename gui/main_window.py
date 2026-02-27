"""PySide6 메인 윈도우 모듈.

드래그앤드롭, 파일 선택, 변환 진행률, 설정 관리를 포함합니다.
"""

from __future__ import annotations

import logging
import sys
import threading
import traceback
from pathlib import Path

from PySide6.QtCore import Qt, Signal, QThread, QObject
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.pdf_handler import (
    get_supported_extensions,
    is_pdf,
    load_image,
    pdf_to_images,
)
from core.ocr_engine import OCREngine, validate_ocr_response
from core.quality_checker import check_image_quality
from core.content_parser import parse_ocr_response, build_document
from core.hwpx_writer import write_exam_to_hwpx
from core.template_loader import load_template
from gui.preview_dialog import PreviewDialog, PageInfo
from utils.config import get_output_dir

logger = logging.getLogger(__name__)


# ─── 백그라운드 변환 워커 ────────────────────────────────────

class ConversionWorker(QObject):
    """백그라운드 스레드에서 변환 수행."""

    progress = Signal(int, str)    # (진행률%, 메시지)
    finished = Signal(str)         # 결과 파일 경로
    error = Signal(str)            # 에러 메시지
    quality_warning = Signal(int, str)   # (페이지번호, 경고메시지)
    ocr_warning = Signal(int, str)       # (페이지번호, 경고메시지)
    # 미리보기 요청: 워커가 GUI 스레드에 미리보기 표시를 요청
    preview_requested = Signal(list)     # list[PageInfo]

    def __init__(
        self,
        file_path: str,
        output_path: str,
        api_key: str,
        template_path: str | None = None,
    ):
        super().__init__()
        self.file_path = file_path
        self.output_path = output_path
        self.api_key = api_key
        self.template_path = template_path
        self._cancelled = False
        # 미리보기 응답 동기화용
        self._preview_event = threading.Event()
        self._preview_approved = False

    def cancel(self):
        self._cancelled = True
        # 미리보기 대기 중이면 깨우기
        self._preview_event.set()

    def set_preview_result(self, approved: bool):
        """GUI 스레드에서 미리보기 결과를 설정."""
        self._preview_approved = approved
        self._preview_event.set()

    def run(self):
        try:
            self._do_conversion()
        except Exception as e:
            logger.exception("변환 중 오류 발생")
            self.error.emit(f"변환 실패: {e}\n\n{traceback.format_exc()}")

    def _do_conversion(self):
        file_path = Path(self.file_path)

        # Step 1: 이미지 로드
        self.progress.emit(5, "파일 로드 중...")
        if is_pdf(file_path):
            images = pdf_to_images(file_path)
        else:
            images = [load_image(file_path)]

        total_pages = len(images)
        self.progress.emit(10, f"{total_pages}페이지 로드 완료")

        # ── Gate 1: 이미지 품질 검사 ──
        self.progress.emit(10, "이미지 품질 검사 중...")
        valid_indices: list[int] = []

        for i, img in enumerate(images):
            if self._cancelled:
                self.error.emit("사용자에 의해 취소되었습니다.")
                return

            quality = check_image_quality(img)
            if not quality.passed:
                warn_msg = (
                    f"페이지 {i + 1} 품질 불합격 (점수 {quality.score:.0f}): "
                    + "; ".join(quality.warnings)
                )
                self.quality_warning.emit(i + 1, warn_msg)
                logger.warning("건너뛰기 — %s", warn_msg)
            else:
                if quality.warnings:
                    warn_msg = (
                        f"페이지 {i + 1} 경고: "
                        + "; ".join(quality.warnings)
                    )
                    self.quality_warning.emit(i + 1, warn_msg)
                valid_indices.append(i)

        if not valid_indices:
            self.error.emit("모든 페이지가 품질 검사에 불합격했습니다.")
            return

        skipped = total_pages - len(valid_indices)
        if skipped > 0:
            self.progress.emit(
                12, f"{skipped}페이지 건너뜀, {len(valid_indices)}페이지 처리 예정"
            )

        # Step 2: OCR 처리
        engine = OCREngine(api_key=self.api_key)
        pages = []
        page_infos: list[PageInfo] = []

        for seq, idx in enumerate(valid_indices):
            if self._cancelled:
                self.error.emit("사용자에 의해 취소되었습니다.")
                return

            img = images[idx]
            page_num = idx + 1
            pct = 15 + int((seq / len(valid_indices)) * 60)
            self.progress.emit(pct, f"OCR 처리 중... ({seq + 1}/{len(valid_indices)})")

            ocr_result = engine.recognize_page(img)

            # ── Gate 2: OCR 응답 검증 ──
            ocr_quality = validate_ocr_response(ocr_result)
            if ocr_quality.warnings:
                warn_msg = (
                    f"페이지 {page_num} OCR 경고: "
                    + "; ".join(ocr_quality.warnings)
                )
                self.ocr_warning.emit(page_num, warn_msg)

            page = parse_ocr_response(ocr_result, page_number=page_num)
            pages.append(page)

            # 미리보기용 정보 수집
            img_quality = check_image_quality(img)
            page_infos.append(PageInfo(
                page_number=page_num,
                image=img,
                exam_page=page,
                image_quality=img_quality,
                ocr_quality=ocr_quality,
            ))

        self.progress.emit(78, "OCR 완료, 미리보기 준비 중...")

        # ── Gate 3: 미리보기 다이얼로그 (GUI 스레드에서 실행) ──
        self._preview_event.clear()
        self._preview_approved = False
        self.preview_requested.emit(page_infos)

        # 사용자 응답 대기
        self._preview_event.wait()

        if self._cancelled or not self._preview_approved:
            self.error.emit("사용자에 의해 취소되었습니다.")
            return

        self.progress.emit(80, "문서 구성 중...")

        # Step 3: 문서 구성
        document = build_document(pages)

        # Step 4: HWPX 생성
        self.progress.emit(90, "HWPX 파일 생성 중...")
        result_path = write_exam_to_hwpx(
            document, self.output_path, template_path=self.template_path
        )

        self.progress.emit(100, "변환 완료!")
        self.finished.emit(str(result_path))


# ─── 메인 윈도우 ────────────────────────────────────────────

class MainWindow(QMainWindow):
    """수학 시험지 → HWPX 변환 메인 윈도우."""

    def __init__(self):
        super().__init__()
        self._worker: ConversionWorker | None = None
        self._thread: QThread | None = None
        self._setup_ui()

    # ── 통일된 크기 상수 ──
    _BTN_HEIGHT = 36
    _PRIMARY_BTN_HEIGHT = 42
    _ICON_SIZE = 16

    def _setup_ui(self):
        self.setWindowTitle("수학 시험지 한글화 변환기")
        self.setMinimumSize(720, 600)
        self.setAcceptDrops(True)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(0)
        layout.setContentsMargins(24, 20, 24, 20)

        # ── 섹션 1: API 키 ──
        section_label = QLabel("API 설정")
        section_label.setStyleSheet(
            "font-size: 11px; font-weight: 600; color: #667085;"
            "text-transform: uppercase; letter-spacing: 1px;"
            "padding: 0; margin: 0;"
        )
        layout.addWidget(section_label)
        layout.addSpacing(6)

        api_layout = QHBoxLayout()
        api_layout.setSpacing(10)
        api_label = QLabel("API 키")
        api_label.setFixedWidth(48)
        api_label.setStyleSheet("font-size: 13px; color: #344054; font-weight: 500;")
        self._api_key_input = QLineEdit()
        self._api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        self._api_key_input.setPlaceholderText("Anthropic API 키를 입력하세요")
        self._api_key_input.setFixedHeight(self._BTN_HEIGHT)
        self._load_api_key()
        api_layout.addWidget(api_label)
        api_layout.addWidget(self._api_key_input)
        layout.addLayout(api_layout)

        # 구분선
        layout.addSpacing(16)
        self._add_separator(layout)
        layout.addSpacing(16)

        # ── 섹션 2: 파일 선택 ──
        section_label2 = QLabel("입력 파일")
        section_label2.setStyleSheet(
            "font-size: 11px; font-weight: 600; color: #667085;"
            "text-transform: uppercase; letter-spacing: 1px;"
            "padding: 0; margin: 0;"
        )
        layout.addWidget(section_label2)
        layout.addSpacing(6)

        self._file_path_label = QLabel("파일을 드래그하거나 선택하세요")
        self._file_path_label.setStyleSheet(
            "QLabel {"
            "  border: 2px dashed #d0d5dd;"
            "  border-radius: 10px;"
            "  padding: 24px;"
            "  background: #f9fafb;"
            "  color: #667085;"
            "  font-size: 13px;"
            "}"
        )
        self._file_path_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._file_path_label.setMinimumHeight(80)
        layout.addWidget(self._file_path_label)
        layout.addSpacing(8)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        self._browse_btn = QPushButton("파일 선택")
        self._browse_btn.clicked.connect(self._browse_file)
        self._browse_btn.setFixedHeight(self._BTN_HEIGHT)
        self._browse_btn.setFixedWidth(120)
        btn_layout.addWidget(self._browse_btn)

        out_label = QLabel("출력")
        out_label.setFixedWidth(30)
        out_label.setStyleSheet("font-size: 13px; color: #344054; font-weight: 500;")
        self._output_input = QLineEdit()
        self._output_input.setPlaceholderText("출력 파일 경로 (자동 생성)")
        self._output_input.setFixedHeight(self._BTN_HEIGHT)
        self._output_browse_btn = QPushButton("경로")
        self._output_browse_btn.setFixedHeight(self._BTN_HEIGHT)
        self._output_browse_btn.setFixedWidth(80)
        self._output_browse_btn.setToolTip("출력 경로 직접 지정")
        self._output_browse_btn.clicked.connect(self._browse_output)
        btn_layout.addWidget(out_label)
        btn_layout.addWidget(self._output_input)
        btn_layout.addWidget(self._output_browse_btn)
        layout.addLayout(btn_layout)
        layout.addSpacing(6)

        # ── 양식 파일(템플릿) 선택 ──
        template_layout = QHBoxLayout()
        template_layout.setSpacing(10)
        self._template_btn = QPushButton("양식 파일 선택")
        self._template_btn.setFixedHeight(self._BTN_HEIGHT)
        self._template_btn.setFixedWidth(120)
        self._template_btn.setToolTip(
            "한글(.hwpx/.hwp) 양식 파일을 선택하면\n해당 서식(여백, 글꼴, 스타일)이 적용됩니다"
        )
        self._template_btn.clicked.connect(self._browse_template)
        template_layout.addWidget(self._template_btn)

        self._template_label = QLabel("양식 없음 (기본 서식)")
        self._template_label.setStyleSheet("color: #98a2b3; font-size: 12px;")
        template_layout.addWidget(self._template_label, 1)

        self._template_clear_btn = QPushButton("해제")
        self._template_clear_btn.setFixedHeight(self._BTN_HEIGHT)
        self._template_clear_btn.setFixedWidth(80)
        self._template_clear_btn.setEnabled(False)
        self._template_clear_btn.clicked.connect(self._clear_template)
        template_layout.addWidget(self._template_clear_btn)
        layout.addLayout(template_layout)

        # 구분선
        layout.addSpacing(16)
        self._add_separator(layout)
        layout.addSpacing(16)

        # ── 섹션 3: 변환 ──
        convert_layout = QHBoxLayout()
        convert_layout.setSpacing(10)
        self._convert_btn = QPushButton("변환 시작")
        self._convert_btn.setFixedHeight(self._PRIMARY_BTN_HEIGHT)
        self._convert_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._convert_btn.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "    stop:0 #1570ef, stop:1 #0d6efd);"
            "  color: white;"
            "  border: none;"
            "  border-radius: 8px;"
            "  font-size: 14px;"
            "  font-weight: 600;"
            "  padding: 0 20px;"
            "}"
            "QPushButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "    stop:0 #0d6efd, stop:1 #0b5ed7);"
            "}"
            "QPushButton:pressed { background: #0b5ed7; }"
            "QPushButton:disabled {"
            "  background: #e4e7ec;"
            "  color: #98a2b3;"
            "}"
        )
        self._convert_btn.clicked.connect(self._start_conversion)
        convert_layout.addWidget(self._convert_btn)

        self._cancel_btn = QPushButton("취소")
        self._cancel_btn.setFixedHeight(self._PRIMARY_BTN_HEIGHT)
        self._cancel_btn.setFixedWidth(80)
        self._cancel_btn.setEnabled(False)
        self._cancel_btn.setStyleSheet(
            "QPushButton {"
            "  border: 1px solid #fda29b;"
            "  border-radius: 8px;"
            "  color: #b42318;"
            "  background: #ffffff;"
            "  font-weight: 500;"
            "}"
            "QPushButton:hover { background: #fef3f2; }"
            "QPushButton:disabled {"
            "  border: 1px solid #e4e7ec;"
            "  color: #98a2b3;"
            "  background: #f2f4f7;"
            "}"
        )
        self._cancel_btn.clicked.connect(self._cancel_conversion)
        convert_layout.addWidget(self._cancel_btn)
        layout.addLayout(convert_layout)
        layout.addSpacing(12)

        # ── 진행률 ──
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)
        self._progress_bar.setTextVisible(True)
        self._progress_bar.setFixedHeight(22)
        layout.addWidget(self._progress_bar)
        layout.addSpacing(4)

        self._status_label = QLabel("대기 중")
        self._status_label.setStyleSheet("color: #667085; font-size: 12px;")
        layout.addWidget(self._status_label)
        layout.addSpacing(12)

        # ── 로그 출력 ──
        log_label = QLabel("변환 로그")
        log_label.setStyleSheet(
            "font-size: 11px; font-weight: 600; color: #667085;"
            "text-transform: uppercase; letter-spacing: 1px;"
            "padding: 0; margin: 0;"
        )
        layout.addWidget(log_label)
        layout.addSpacing(6)

        self._log_output = QPlainTextEdit()
        self._log_output.setReadOnly(True)
        self._log_output.setFont(QFont("Consolas", 9))
        self._log_output.setMaximumBlockCount(500)
        self._log_output.setPlaceholderText("변환 로그가 여기에 표시됩니다...")
        layout.addWidget(self._log_output)

        self._selected_file: str | None = None
        self._selected_template: str | None = None

    @staticmethod
    def _add_separator(layout: QVBoxLayout):
        """얇은 구분선 추가."""
        sep = QWidget()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background-color: #e4e7ec;")
        layout.addWidget(sep)

    def _load_api_key(self):
        """설정에서 API 키 로드."""
        try:
            from utils.config import get_api_key
            key = get_api_key()
            self._api_key_input.setText(key)
        except ValueError:
            pass

    def _log(self, msg: str):
        """로그 메시지 추가."""
        self._log_output.appendPlainText(msg)

    # ── 드래그앤드롭 ──

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and self._is_supported_file(urls[0].toLocalFile()):
                event.acceptProposedAction()
                self._file_path_label.setStyleSheet(
                    "QLabel {"
                    "  border: 2px dashed #0d6efd;"
                    "  border-radius: 10px;"
                    "  padding: 24px;"
                    "  background: #eff8ff;"
                    "  color: #1570ef;"
                    "  font-size: 13px;"
                    "  font-weight: 500;"
                    "}"
                )

    def dragLeaveEvent(self, event):
        self._reset_drop_style()

    def dropEvent(self, event: QDropEvent):
        self._reset_drop_style()
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if self._is_supported_file(file_path):
                self._set_file(file_path)

    def _reset_drop_style(self):
        self._file_path_label.setStyleSheet(
            "QLabel {"
            "  border: 2px dashed #d0d5dd;"
            "  border-radius: 10px;"
            "  padding: 24px;"
            "  background: #f9fafb;"
            "  color: #667085;"
            "  font-size: 13px;"
            "}"
        )

    @staticmethod
    def _is_supported_file(path: str) -> bool:
        ext = Path(path).suffix.lower()
        return ext in get_supported_extensions()

    def _set_file(self, path: str):
        self._selected_file = path
        name = Path(path).name
        self._file_path_label.setText(name)
        self._file_path_label.setStyleSheet(
            "QLabel {"
            "  border: 2px solid #32d583;"
            "  border-radius: 10px;"
            "  padding: 24px;"
            "  background: #ecfdf3;"
            "  color: #027a48;"
            "  font-size: 13px;"
            "  font-weight: 600;"
            "}"
        )

        # 출력 경로 자동 설정
        out_name = Path(path).stem + "_변환.hwpx"
        out_path = get_output_dir() / out_name
        self._output_input.setText(str(out_path))
        self._log(f"파일 선택: {path}")

    # ── 파일 선택 ──

    def _browse_file(self):
        exts = " ".join(f"*{e}" for e in sorted(get_supported_extensions()))
        path, _ = QFileDialog.getOpenFileName(
            self,
            "시험지 파일 선택",
            "",
            f"지원 파일 ({exts});;모든 파일 (*.*)",
        )
        if path:
            self._set_file(path)

    def _browse_output(self):
        """출력 경로 직접 지정."""
        current = self._output_input.text().strip()
        start_dir = str(Path(current).parent) if current else ""
        path, _ = QFileDialog.getSaveFileName(
            self,
            "출력 파일 경로 지정",
            current or start_dir,
            "HWPX 파일 (*.hwpx);;모든 파일 (*.*)",
        )
        if path:
            self._output_input.setText(path)
            self._log(f"출력 경로 지정: {path}")

    # ── 양식 파일(템플릿) ──

    def _browse_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "양식 파일 선택",
            "",
            "한글 파일 (*.hwpx *.hwp);;HWPX 파일 (*.hwpx);;HWP 파일 (*.hwp);;모든 파일 (*.*)",
        )
        if path:
            self._set_template(path)

    def _set_template(self, path: str):
        """템플릿 파일 설정 및 유효성 검사."""
        try:
            config = load_template(path)
        except (FileNotFoundError, ValueError) as e:
            QMessageBox.warning(
                self, "양식 파일 오류", f"양식 파일을 사용할 수 없습니다.\n\n{e}"
            )
            return

        self._selected_template = path
        name = Path(path).name
        self._template_label.setText(name)
        self._template_label.setStyleSheet(
            "color: #027a48; font-size: 12px; font-weight: 600;"
        )
        self._template_clear_btn.setEnabled(True)
        self._log(f"양식 파일 설정: {path}")
        self._log(f"  {config.summary}")

    def _clear_template(self):
        """템플릿 선택 해제."""
        self._selected_template = None
        self._template_label.setText("양식 없음 (기본 서식)")
        self._template_label.setStyleSheet("color: #98a2b3; font-size: 12px;")
        self._template_clear_btn.setEnabled(False)
        self._log("양식 파일 해제됨")

    # ── 변환 ──

    def _start_conversion(self):
        if not self._selected_file:
            QMessageBox.warning(self, "알림", "변환할 파일을 선택하세요.")
            return

        api_key = self._api_key_input.text().strip()
        if not api_key or api_key == "your-api-key-here":
            QMessageBox.warning(self, "알림", "Anthropic API 키를 입력하세요.")
            return

        output_path = self._output_input.text().strip()
        if not output_path:
            QMessageBox.warning(self, "알림", "출력 경로를 지정하세요.")
            return

        # 출력 파일 이미 존재 시 경고
        if Path(output_path).exists():
            reply = QMessageBox.warning(
                self,
                "파일 덮어쓰기 확인",
                f"이미 같은 이름의 파일이 존재합니다.\n\n"
                f"{Path(output_path).name}\n\n"
                f"덮어쓰시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                self._log("변환 취소: 파일 덮어쓰기 거부")
                return

        # 출력 디렉토리 생성
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        self._set_ui_converting(True)
        self._progress_bar.setValue(0)
        self._log("=" * 40)
        self._log("변환 시작...")

        # 워커 스레드 생성
        self._thread = QThread()
        self._worker = ConversionWorker(
            self._selected_file, output_path, api_key,
            template_path=self._selected_template,
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.quality_warning.connect(self._on_quality_warning)
        self._worker.ocr_warning.connect(self._on_ocr_warning)
        self._worker.preview_requested.connect(self._on_preview_requested)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)

        self._thread.start()

    def _cancel_conversion(self):
        if self._worker:
            self._worker.cancel()
            self._log("취소 요청됨...")

    def _on_progress(self, percent: int, message: str):
        self._progress_bar.setValue(percent)
        self._status_label.setText(message)
        self._log(f"[{percent}%] {message}")

    def _on_finished(self, result_path: str):
        self._set_ui_converting(False)
        self._log(f"변환 완료: {result_path}")
        self._status_label.setText("변환 완료!")

        QMessageBox.information(
            self,
            "변환 완료",
            f"HWPX 파일이 생성되었습니다.\n\n{result_path}",
        )

    def _on_error(self, error_msg: str):
        self._set_ui_converting(False)
        self._log(f"오류: {error_msg}")
        self._status_label.setText("오류 발생")
        self._progress_bar.setValue(0)

        QMessageBox.critical(self, "변환 오류", error_msg)

    def _cleanup_thread(self):
        self._thread = None
        self._worker = None

    def _on_quality_warning(self, page_num: int, message: str):
        self._log(f"[품질] {message}")

    def _on_ocr_warning(self, page_num: int, message: str):
        self._log(f"[OCR] {message}")

    def _on_preview_requested(self, page_infos: list):
        """워커에서 미리보기 요청 → GUI 스레드에서 다이얼로그 표시."""
        dialog = PreviewDialog(page_infos, parent=self)
        result = dialog.exec()
        approved = result == PreviewDialog.DialogCode.Accepted
        if self._worker:
            self._worker.set_preview_result(approved)

    def _set_ui_converting(self, converting: bool):
        self._convert_btn.setEnabled(not converting)
        self._cancel_btn.setEnabled(converting)
        self._browse_btn.setEnabled(not converting)
        self._output_browse_btn.setEnabled(not converting)
        self._api_key_input.setEnabled(not converting)
        self._template_btn.setEnabled(not converting)
        self._template_clear_btn.setEnabled(
            not converting and self._selected_template is not None
        )
