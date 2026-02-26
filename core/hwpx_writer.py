"""HWPX 문서 생성 모듈.

python-hwpx로 문서 골격을 생성하고, lxml로 수식 XML을 직접 삽입합니다.
템플릿 HWPX가 주어지면 해당 파일을 복사하여 서식·글꼴을 유지합니다.
"""

from __future__ import annotations

import logging
import random
import re
import shutil
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

from models.exam_document import (
    ExamDocument,
    ExamPage,
    Question,
    Choice,
    ContentBlock,
    ContentType,
)
from models.template_config import TemplateConfig
from core.latex_to_hwpeq import latex_to_hwpeq, latex_to_image

logger = logging.getLogger(__name__)

# HWPX 네임스페이스
NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
}

# 원문자 ①②③④⑤
CIRCLE_NUMBERS = {1: "\u2460", 2: "\u2461", 3: "\u2462", 4: "\u2463", 5: "\u2464"}

# 우측정렬 문단 속성 ID (배점용, header.xml에 동적 추가)
_RIGHT_ALIGN_PR_ID = "100"


def _qn(prefix: str, local: str) -> str:
    """Clark notation으로 네임스페이스 태그 생성."""
    return f"{{{NS[prefix]}}}{local}"


def _random_id() -> str:
    """랜덤 ID 생성 (HWPX 요소용)."""
    return str(random.randint(100000000, 4294967295))


# 기호로 렌더링되는 HWP 명령어 (1문자 기호로 치환됨)
# 긴 이름을 먼저 배치하여 부분 일치 방지 (예: "varepsilon" 전에 "epsilon" 처리 방지)
_SYMBOL_KEYWORDS = [
    # latex_to_hwpeq.py가 생성하는 대문자 키워드
    "PLUSMINUS", "MINUSPLUS", "SMALLUNION", "SMALLINTER",
    "APPROX", "PROPTO", "LAPLACE", "BULLET", "TRIANGLE",
    "DIAMOND", "SQUARE",
    "EQUIV", "SIMEQ", "ASYMP", "DOTEQ",
    "TIMES", "CDOT", "EXIST",
    "WEDGE", "LNOT", "OPLUS", "OTIMES",
    "VDASH", "MODELS",
    "PREC", "SUCC", "CONG", "OWNS", "CIRC", "STAR",
    "DIV", "LEQ", "GEQ", "SIM", "VEE", "BOT", "TOP",
    # 소문자 그리스 문자
    "varepsilon", "vartheta", "varphi",
    "epsilon", "upsilon", "lambda",
    "alpha", "beta", "gamma", "delta", "zeta", "eta", "theta",
    "iota", "kappa", "mu", "nu", "xi", "pi", "rho", "sigma",
    "tau", "phi", "chi", "psi", "omega",
    # 기타 소문자 키워드
    "partial", "therefore", "because", "forall", "exists",
    "emptyset", "subseteq", "supseteq", "subset", "supset",
    "notin", "parallel",
    "infty", "prime", "dprime", "angle", "nabla", "bullet",
    "approx", "propto", "equiv", "neq", "leq", "geq", "sim",
    "cdot", "times", "div", "pm", "mp", "inf", "in",
]

# 대형 연산자 (1문자 넓은 기호)
_LARGE_OP_KEYWORDS = ["SUM", "PROD", "OINT", "DINT", "TINT", "INT"]

# 구조 명령어 (렌더링에 기여하지 않음)
_STRUCT_KEYWORDS = [
    "eqalign", "matrix", "cases", "pile",
    "sqrt", "root", "of",
    "over", "atop",
    "from", "left", "right",
    "roman", "bold", "ital",
    "to",
]


def _extract_latex_brace(s: str) -> tuple[str, str]:
    """LaTeX 문자열에서 첫 번째 {content}를 추출.

    Returns:
        (content, rest_of_string)
    """
    s = s.lstrip()
    if not s or s[0] != "{":
        return ("", s)
    depth = 0
    for i, ch in enumerate(s):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return (s[1:i], s[i + 1 :])
    return (s[1:], "")


def _measure_latex_size(latex: str) -> tuple[int, int] | None:
    """matplotlib로 LaTeX 수식의 렌더링 크기를 측정.

    분수(\\frac) 포함 시 분자·분모를 개별 측정하여 합성합니다.
    HWP 수식 렌더러는 분수 내용을 약 75% 크기로 렌더링하므로,
    분자·분모를 fontsize×0.75로 측정해서 정확한 너비를 구합니다.

    Returns:
        (width, height) in hwpunit, or None if measurement fails.
    """
    try:
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        from matplotlib.figure import Figure

        fig = Figure(figsize=(10, 2), dpi=72)
        canvas = FigureCanvasAgg(fig)
        renderer = canvas.get_renderer()

        def mpl_width(expr: str, fontsize: float = 10.0) -> int:
            """LaTeX 식의 렌더링 너비를 hwpunit로 반환."""
            if not expr.strip():
                return 0
            t = fig.text(0, 0, f"${expr}$", fontsize=fontsize)
            bb = t.get_window_extent(renderer)
            w = int(bb.width * 100)
            t.remove()
            return w

        # 분수 포함 여부 확인
        frac_match = re.search(r"\\(?:d?frac)\s*\{", latex)
        if frac_match:
            # prefix: \frac 앞의 내용
            prefix = latex[: frac_match.start()].strip()
            rest = latex[frac_match.end() - 1 :]  # '{' 포함

            # 분자·분모 추출
            numerator, after_num = _extract_latex_brace(rest)
            denominator, suffix = _extract_latex_brace(after_num)
            suffix = suffix.strip()

            # 각 부분 측정: 분수 내용은 85% 크기 (HWP 수식 렌더러 기준)
            FRAC_SCALE = 0.85
            w_prefix = mpl_width(prefix) if prefix else 0
            w_num = mpl_width(numerator, 10 * FRAC_SCALE)
            w_den = mpl_width(denominator, 10 * FRAC_SCALE)
            w_suffix = mpl_width(suffix) if suffix else 0

            # 합성: prefix + space + fraction + space + suffix
            w_frac = max(w_num, w_den) + 300  # fraction bar padding
            w = w_prefix
            if w_prefix:
                w += 200  # prefix와 분수 사이 간격
            w += w_frac
            if w_suffix:
                w += 200  # 분수와 suffix 사이 간격
                w += w_suffix

            h = 2400  # 분수 기본 높이
            if "\\sqrt" in latex:
                h = max(h, 3200)

            return (max(w, 800), h)

        # 분수 없는 경우: 직접 측정
        w = mpl_width(latex)
        h = 1200
        if "\\sqrt" in latex:
            h = 1600

        return (max(w, 800), h)
    except Exception:
        return None


def _estimate_equation_size(hwp_eq_script: str) -> tuple[int, int]:
    """수식 크기 추정 (폴백용, matplotlib 측정 실패 시 사용).

    HWP 수식 스크립트에서 가시 문자를 세어 크기를 추정합니다.
    """
    has_fraction = "over" in hwp_eq_script or "atop" in hwp_eq_script
    has_sqrt = "sqrt" in hwp_eq_script or "root" in hwp_eq_script

    if has_fraction:
        parts = hwp_eq_script.replace("atop", "over").split("over")
        visible_len = max(_visual_char_count(p) for p in parts)
    else:
        visible_len = _visual_char_count(hwp_eq_script)

    CHAR_WIDTH = 650
    PADDING = 200
    width = max(int(visible_len * CHAR_WIDTH) + PADDING, 800)

    if has_fraction:
        width = int(width * 1.4)
    if has_sqrt:
        width = int(width * 1.4)

    height = 1200
    if has_fraction:
        height = 2400
    if has_sqrt:
        height = max(height, 1600)
    if has_fraction and has_sqrt:
        height = max(height, 3200)

    return (width, height)


def _visual_char_count(text: str) -> float:
    """HWP 수식 스크립트의 가시 문자 수 (폴백용)."""
    s = text
    for kw in _SYMBOL_KEYWORDS:
        s = s.replace(kw, "G")
    for kw in _LARGE_OP_KEYWORDS:
        s = s.replace(kw, "W")
    for cmd in _STRUCT_KEYWORDS:
        s = s.replace(cmd, "")

    sup_sub_chars = 0
    for m in re.finditer(r'[\^_]\{([^{}]*)\}', s):
        content = m.group(1).replace(" ", "")
        sup_sub_chars += len(content)
    s_no_brace = re.sub(r'[\^_]\{[^{}]*\}', '', s)
    for m in re.finditer(r'[\^_](\S)', s_no_brace):
        sup_sub_chars += 1

    for ch in "{}^_":
        s = s.replace(ch, "")
    total = len(s.replace(" ", "").strip())

    base = total - sup_sub_chars
    return base + sup_sub_chars * 0.5


class HWPXWriter:
    """HWPX 문서 생성기."""

    def __init__(self):
        self._image_counter = 0
        self._eq_counter = 0
        self._embedded_images: dict[str, bytes] = {}  # bindata id → image bytes

    def write(
        self,
        document: ExamDocument,
        output_path: str | Path,
        template: Optional[TemplateConfig] = None,
    ) -> Path:
        """ExamDocument를 HWPX 파일로 저장.

        Args:
            document: 변환할 시험 문서
            output_path: 출력 파일 경로
            template: 템플릿 설정 (None이면 기본 서식)

        Returns:
            저장된 파일 경로
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if template and template.is_valid:
            return self._write_with_template(document, output_path, template)
        else:
            return self._write_default(document, output_path)

    def _write_default(self, document: ExamDocument, output_path: Path) -> Path:
        """기본 서식으로 HWPX 파일 생성 (기존 동작)."""
        from hwpx import HwpxDocument
        doc = HwpxDocument.new()

        # 첫 번째 섹션 가져오기
        section = doc.sections[0]
        sec_elem = section.element

        # 기존 문단 처리: secPr이 포함된 첫 문단은 유지하고 나머지 제거
        existing_paras = sec_elem.findall(_qn("hp", "p"))
        for i, p in enumerate(existing_paras):
            if i == 0:
                # 첫 문단(secPr 포함)은 유지 - 텍스트 run만 정리
                for run in p.findall(_qn("hp", "run")):
                    # secPr이 없는 빈 run 제거
                    if run.find(_qn("hp", "secPr")) is None:
                        t_elem = run.find(_qn("hp", "t"))
                        if t_elem is not None and not (t_elem.text and t_elem.text.strip()):
                            p.remove(run)
            else:
                sec_elem.remove(p)

        # 문서 제목 삽입
        if document.title:
            self._add_title_paragraph(sec_elem, document.title)

        # 페이지별 내용 삽입
        for page in document.pages:
            self._write_page(sec_elem, page)

        # 모든 문단에 linesegarray 보장 (한글 필수 요소)
        self._ensure_linesegarray(sec_elem)

        # 섹션 변경 표시 후 저장
        section.mark_dirty()
        doc.save_to_path(str(output_path))

        # 우측정렬 paraPr 추가 (배점용)
        self._inject_right_align_parapr(output_path)

        # 수식 이미지가 있으면 ZIP에 추가
        if self._embedded_images:
            self._inject_images_to_zip(output_path)

        logger.info("HWPX 파일 저장 완료: %s", output_path)
        return output_path

    def _write_with_template(
        self,
        document: ExamDocument,
        output_path: Path,
        template: TemplateConfig,
    ) -> Path:
        """템플릿 HWPX를 복사하여 내용만 교체.

        전략:
        1. 템플릿 HWPX를 output 경로에 복사
        2. ZIP 내부의 section0.xml에서 내용 문단만 제거 (secPr 보존)
        3. 새 OCR 결과를 section0.xml에 삽입
        4. header.xml은 그대로 유지 → 글꼴·스타일 정의 보존
        """
        logger.info("템플릿 기반 변환: %s", template.source_path.name)

        # 1. 템플릿 파일을 output에 복사
        shutil.copy2(str(template.source_path), str(output_path))

        # 2. ZIP에서 section0.xml 읽기
        with zipfile.ZipFile(str(output_path), "r") as zf:
            section_bytes = zf.read("Contents/section0.xml")

        # 3. section0.xml 파싱 → 내용 교체
        root = etree.fromstring(section_bytes)
        sec_elem = root  # section0.xml의 루트가 곧 섹션 요소

        # secPr이 포함된 첫 문단 찾기 & 보존
        first_para_with_secpr = None
        all_paras = list(sec_elem.findall(_qn("hp", "p")))

        for p in all_paras:
            has_secpr = False
            for elem in p.iter():
                if elem.tag and "secPr" in elem.tag:
                    has_secpr = True
                    break
            if has_secpr and first_para_with_secpr is None:
                first_para_with_secpr = p
                # secPr 문단의 텍스트 run만 제거 (secPr 자체는 보존)
                for run in list(p.findall(_qn("hp", "run"))):
                    has_secpr_child = False
                    for child in run.iter():
                        if child.tag and "secPr" in child.tag:
                            has_secpr_child = True
                            break
                    if not has_secpr_child:
                        p.remove(run)
                # linesegarray 제거 (새로 생성할 것)
                for lsa in list(p.findall(_qn("hp", "linesegarray"))):
                    p.remove(lsa)
            else:
                sec_elem.remove(p)

        # 4. 새 내용 삽입
        if document.title:
            self._add_title_paragraph(sec_elem, document.title)

        for page in document.pages:
            self._write_page(sec_elem, page)

        # 5. linesegarray 보장
        self._ensure_linesegarray(sec_elem)

        # 6. 수정된 section0.xml을 ZIP에 다시 기록
        new_section_bytes = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        self._replace_in_zip(output_path, "Contents/section0.xml", new_section_bytes)

        # 우측정렬 paraPr 추가 (배점용)
        self._inject_right_align_parapr(output_path)

        # 수식 이미지가 있으면 ZIP에 추가
        if self._embedded_images:
            self._inject_images_to_zip(output_path)

        logger.info("템플릿 기반 HWPX 파일 저장 완료: %s", output_path)
        return output_path

    @staticmethod
    def _replace_in_zip(zip_path: Path, entry_name: str, new_data: bytes):
        """ZIP 파일 내 특정 엔트리를 교체."""
        temp_path = zip_path.with_suffix(".hwpx.tmp")

        with zipfile.ZipFile(str(zip_path), "r") as zin:
            with zipfile.ZipFile(str(temp_path), "w") as zout:
                for item in zin.infolist():
                    if item.filename == entry_name:
                        zout.writestr(item, new_data)
                    else:
                        zout.writestr(item, zin.read(item.filename))

        shutil.move(str(temp_path), str(zip_path))

    def _add_title_paragraph(self, sec_elem: etree._Element, title: str):
        """제목 문단 추가."""
        p_elem = self._create_paragraph(sec_elem)
        run = self._create_run(p_elem, char_pr_id="1")  # Bold style
        self._set_run_text(run, title)

    def _write_page(self, sec_elem: etree._Element, page: ExamPage):
        """한 페이지 내용을 섹션에 삽입."""
        # 페이지 헤더
        if page.header_text and page.page_number > 1:
            p_elem = self._create_paragraph(sec_elem)
            run = self._create_run(p_elem)
            self._set_run_text(run, page.header_text)

        # 문제별 삽입
        for question in page.questions:
            self._write_question(sec_elem, question)

    def _write_question(self, sec_elem: etree._Element, question: Question):
        """문제 하나를 삽입."""
        # 문제 본문 첫 줄에 번호 포함
        p_elem = self._create_paragraph(sec_elem)

        # 번호 run
        run = self._create_run(p_elem, char_pr_id="1")
        self._set_run_text(run, f"{question.number}. ")

        # 본문 내용
        for block in question.contents:
            self._write_content_block(sec_elem, p_elem, block)

        # 배점을 별도 우측정렬 문단으로 추가
        if question.score:
            score_para = self._create_paragraph(
                sec_elem, para_pr_id=_RIGHT_ALIGN_PR_ID
            )
            run = self._create_run(score_para)
            self._set_run_text(run, f"[{question.score}점]")

        # 선택지
        for choice in question.choices:
            self._write_choice(sec_elem, choice)

        # 소문항
        for sub in question.sub_questions:
            self._write_question(sec_elem, sub)

        # 문제 간 빈 줄
        self._create_paragraph(sec_elem)

    def _write_choice(self, sec_elem: etree._Element, choice: Choice):
        """선택지 하나를 삽입."""
        p_elem = self._create_paragraph(sec_elem)

        # 원문자 번호
        circle = CIRCLE_NUMBERS.get(choice.number, f"({choice.number})")
        run = self._create_run(p_elem)
        self._set_run_text(run, f"  {circle} ")

        for block in choice.contents:
            self._write_content_block(sec_elem, p_elem, block)

    def _write_content_block(
        self,
        sec_elem: etree._Element,
        current_para: etree._Element,
        block: ContentBlock,
    ):
        """콘텐츠 블록을 삽입."""
        if block.type == ContentType.TEXT:
            run = self._create_run(current_para)
            self._set_run_text(run, block.value)

        elif block.type == ContentType.EQUATION:
            # 인라인 수식
            self._insert_equation(current_para, block.value)

        elif block.type == ContentType.EQUATION_BLOCK:
            # 블록 수식: 새 문단에 삽입
            eq_para = self._create_paragraph(sec_elem)
            self._insert_equation(eq_para, block.value)

    def _insert_equation(self, p_elem: etree._Element, latex: str):
        """수식을 문단에 삽입.

        1차: matplotlib로 수식 크기 측정 → 네이티브 HWP 수식 삽입
        2차: 측정 실패 시 휴리스틱 추정으로 폴백
        3차: 수식 변환 실패 시 이미지 폴백
        """
        try:
            hwp_eq = latex_to_hwpeq(latex)
            # matplotlib로 실제 수식 크기 측정 (정확), 실패 시 None → 폴백
            measured_size = _measure_latex_size(latex)
            self._inject_equation_xml(p_elem, hwp_eq, size=measured_size)
        except Exception as e:
            logger.warning("수식 변환 실패, 이미지 폴백: %s (%s)", latex, e)
            try:
                img_data = latex_to_image(latex)
                self._inject_equation_image(p_elem, img_data, latex)
            except Exception as e2:
                logger.error("수식 이미지 생성도 실패: %s (%s)", latex, e2)
                # 최후 수단: LaTeX 텍스트 그대로 삽입
                run = self._create_run(p_elem)
                self._set_run_text(run, f"[{latex}]")

    def _inject_equation_xml(
        self,
        p_elem: etree._Element,
        hwp_eq_script: str,
        size: tuple[int, int] | None = None,
    ):
        """네이티브 HWPX 수식 XML을 문단에 삽입.

        Args:
            p_elem: 부모 문단 요소
            hwp_eq_script: HWP 수식 스크립트
            size: (width, height) in hwpunit. None이면 휴리스틱 추정.
        """
        self._eq_counter += 1

        if size is None:
            size = _estimate_equation_size(hwp_eq_script)

        run = etree.SubElement(p_elem, _qn("hp", "run"))
        run.set("charPrIDRef", "0")

        eq = etree.SubElement(run, _qn("hp", "equation"))
        eq.set("id", _random_id())
        eq.set("zOrder", str(self._eq_counter))
        eq.set("numberingType", "EQUATION")
        eq.set("textWrap", "TOP_AND_BOTTOM")
        eq.set("textFlow", "BOTH_SIDES")
        eq.set("lock", "0")
        eq.set("dropcapstyle", "None")
        eq.set("version", "Equation Version 60")
        eq.set("baseLine", "85")
        eq.set("textColor", "#000000")
        eq.set("baseUnit", "1000")
        eq.set("lineMode", "CHAR")
        eq.set("font", "HYhwpEQ")

        # ShapeSize — matplotlib 측정 또는 휴리스틱 추정
        est_width, est_height = size
        sz = etree.SubElement(eq, _qn("hp", "sz"))
        sz.set("width", str(est_width))
        sz.set("height", str(est_height))
        sz.set("widthRelTo", "ABSOLUTE")
        sz.set("heightRelTo", "ABSOLUTE")
        sz.set("protect", "0")

        # ShapePosition — 글자처럼 취급 (인라인)
        pos = etree.SubElement(eq, _qn("hp", "pos"))
        pos.set("treatAsChar", "1")
        pos.set("affectLSpacing", "1")
        pos.set("flowWithText", "1")
        pos.set("allowOverlap", "0")
        pos.set("holdAnchorAndSO", "0")
        pos.set("vertRelTo", "PARA")
        pos.set("horzRelTo", "PARA")
        pos.set("vertAlign", "TOP")
        pos.set("horzAlign", "LEFT")
        pos.set("vertOffset", "0")
        pos.set("horzOffset", "0")

        # 외부 여백 (인라인 수식-텍스트 간 간격)
        out_margin = etree.SubElement(eq, _qn("hp", "outMargin"))
        out_margin.set("left", "170")
        out_margin.set("right", "170")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")

        # 수식 주석
        shape_comment = etree.SubElement(eq, _qn("hp", "shapeComment"))
        shape_comment.text = "수식입니다."

        # 수식 스크립트
        script = etree.SubElement(eq, _qn("hp", "script"))
        script.text = hwp_eq_script

    def _inject_equation_image(
        self, p_elem: etree._Element, img_data: bytes, alt_text: str
    ):
        """수식 이미지를 문단에 인라인 삽입 (폴백).

        이미지를 bindata로 ZIP에 포함하고 참조합니다.
        """
        self._image_counter += 1
        img_id = f"eq_img_{self._image_counter}"
        filename = f"{img_id}.png"

        # 이미지 바이트 저장 (나중에 ZIP에 추가)
        self._embedded_images[filename] = img_data

        # 인라인 이미지 XML
        run = etree.SubElement(p_elem, _qn("hp", "run"))
        run.set("charPrIDRef", "0")

        ctrl = etree.SubElement(run, _qn("hp", "ctrl"))
        pic = etree.SubElement(ctrl, _qn("hp", "pic"))
        pic.set("id", _random_id())
        pic.set("width", "5000")  # 대략적 크기 (HWP unit)
        pic.set("height", "2000")

        img_rect = etree.SubElement(pic, _qn("hp", "imgRect"))
        img_rect.set("x", "0")
        img_rect.set("y", "0")
        img_rect.set("cx", "5000")
        img_rect.set("cy", "2000")

        img_clip = etree.SubElement(pic, _qn("hp", "imgClip"))
        img_clip.set("left", "0")
        img_clip.set("top", "0")
        img_clip.set("right", "0")
        img_clip.set("bottom", "0")

        img_data_elem = etree.SubElement(pic, _qn("hp", "imgData"))
        img_data_elem.text = f"BinData/{filename}"

    def _inject_images_to_zip(self, hwpx_path: Path):
        """저장된 HWPX ZIP 파일에 수식 이미지를 추가."""
        temp_path = hwpx_path.with_suffix(".hwpx.tmp")

        with zipfile.ZipFile(str(hwpx_path), "r") as zin:
            with zipfile.ZipFile(str(temp_path), "w") as zout:
                # 기존 파일 복사
                for item in zin.infolist():
                    zout.writestr(item, zin.read(item.filename))

                # 이미지 추가
                for filename, data in self._embedded_images.items():
                    zout.writestr(f"BinData/{filename}", data)

        # 원본 교체
        shutil.move(str(temp_path), str(hwpx_path))

    @staticmethod
    def _inject_right_align_parapr(zip_path: Path):
        """header.xml에 우측정렬 문단 속성(paraPr)을 추가.

        배점 표시용 RIGHT 정렬 paraPr를 header.xml의 paraProperties에 추가합니다.
        """
        HH = "http://www.hancom.co.kr/hwpml/2011/head"
        HC = "http://www.hancom.co.kr/hwpml/2011/core"
        HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"

        with zipfile.ZipFile(str(zip_path), "r") as zf:
            header_bytes = zf.read("Contents/header.xml")

        root = etree.fromstring(header_bytes)

        # paraProperties 찾기
        para_props = root.find(f".//{{{HH}}}paraProperties")
        if para_props is None:
            return

        # 이미 같은 ID가 있으면 스킵
        for pp in para_props.findall(f"{{{HH}}}paraPr"):
            if pp.get("id") == _RIGHT_ALIGN_PR_ID:
                return

        # 기존 paraPr id=0을 기반으로 복제 후 정렬만 RIGHT로 변경
        base_pr = para_props.find(f"{{{HH}}}paraPr[@id='0']")
        if base_pr is None:
            return

        import copy
        new_pr = copy.deepcopy(base_pr)
        new_pr.set("id", _RIGHT_ALIGN_PR_ID)

        # align 요소의 horizontal을 RIGHT로 변경
        align_elem = new_pr.find(f"{{{HH}}}align")
        if align_elem is not None:
            align_elem.set("horizontal", "RIGHT")

        para_props.append(new_pr)

        # itemCnt 업데이트
        old_cnt = int(para_props.get("itemCnt", "0"))
        para_props.set("itemCnt", str(old_cnt + 1))

        new_header = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        HWPXWriter._replace_in_zip(zip_path, "Contents/header.xml", new_header)

    @staticmethod
    def _ensure_linesegarray(sec_elem: etree._Element):
        """모든 문단에 linesegarray가 있는지 확인하고, 없으면 추가.

        한글(HWP)은 모든 <hp:p>에 <hp:linesegarray>를 필수로 요구합니다.
        linesegarray는 문단의 마지막 자식으로 위치해야 합니다.
        """
        for p in sec_elem.findall(_qn("hp", "p")):
            if p.find(_qn("hp", "linesegarray")) is None:
                lsa = etree.SubElement(p, _qn("hp", "linesegarray"))
                ls = etree.SubElement(lsa, _qn("hp", "lineseg"))
                ls.set("textpos", "0")
                ls.set("vertpos", "0")
                ls.set("vertsize", "1000")
                ls.set("textheight", "1000")
                ls.set("baseline", "850")
                ls.set("spacing", "600")
                ls.set("horzpos", "0")
                ls.set("horzsize", "42520")
                ls.set("flags", "393216")

    # ─── 저수준 XML 헬퍼 ─────────────────────────────────────

    @staticmethod
    def _create_paragraph(
        parent: etree._Element, para_pr_id: str = "0", style_id: str = "0"
    ) -> etree._Element:
        """문단 요소 생성."""
        p = etree.SubElement(parent, _qn("hp", "p"))
        p.set("id", _random_id())
        p.set("paraPrIDRef", para_pr_id)
        p.set("styleIDRef", style_id)
        p.set("pageBreak", "0")
        p.set("columnBreak", "0")
        p.set("merged", "0")
        return p

    @staticmethod
    def _create_run(
        p_elem: etree._Element, char_pr_id: str = "0"
    ) -> etree._Element:
        """텍스트 run 요소 생성."""
        run = etree.SubElement(p_elem, _qn("hp", "run"))
        run.set("charPrIDRef", char_pr_id)
        return run

    @staticmethod
    def _set_run_text(run: etree._Element, text: str):
        """run 요소에 텍스트 설정."""
        t = etree.SubElement(run, _qn("hp", "t"))
        t.text = text


def write_exam_to_hwpx(
    document: ExamDocument,
    output_path: str | Path,
    template_path: str | Path | None = None,
) -> Path:
    """편의 함수: ExamDocument를 HWPX 파일로 저장.

    Args:
        document: 변환할 시험 문서
        output_path: 출력 파일 경로
        template_path: 양식 HWPX 파일 경로 (None이면 기본 서식)

    Returns:
        저장된 파일 경로
    """
    template = None
    if template_path:
        from core.template_loader import load_template
        template = load_template(template_path)

    writer = HWPXWriter()
    return writer.write(document, output_path, template=template)
