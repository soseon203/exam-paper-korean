"""HWPX 문서 생성 모듈.

python-hwpx로 문서 골격을 생성하고, lxml로 수식 XML을 직접 삽입합니다.
템플릿 HWPX가 주어지면 해당 파일을 복사하여 서식·글꼴을 유지합니다.
"""

from __future__ import annotations

import logging
import random
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


def _qn(prefix: str, local: str) -> str:
    """Clark notation으로 네임스페이스 태그 생성."""
    return f"{{{NS[prefix]}}}{local}"


def _random_id() -> str:
    """랜덤 ID 생성 (HWPX 요소용)."""
    return str(random.randint(100000000, 4294967295))


class HWPXWriter:
    """HWPX 문서 생성기."""

    def __init__(self):
        self._image_counter = 0
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
        # 문제 번호 + 배점 행
        header_parts = [f"{question.number}."]
        if question.score:
            header_parts.append(f" [{question.score}점]")

        # 문제 본문 첫 줄에 번호 포함
        p_elem = self._create_paragraph(sec_elem)

        # 번호 + 배점 run
        run = self._create_run(p_elem, char_pr_id="1")
        self._set_run_text(run, " ".join(header_parts) + " ")

        # 본문 내용
        for block in question.contents:
            self._write_content_block(sec_elem, p_elem, block)

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

        1차: 네이티브 HWP 수식 스크립트 삽입 시도
        2차: 수식 이미지 폴백
        """
        try:
            hwp_eq = latex_to_hwpeq(latex)
            self._inject_equation_xml(p_elem, hwp_eq)
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

    def _inject_equation_xml(self, p_elem: etree._Element, hwp_eq_script: str):
        """네이티브 HWPX 수식 XML을 문단에 삽입.

        HWPX의 수식은 다음 구조:
        <hp:run>
          <hp:ctrl>
            <hp:eqEdit charPrIDRef="0" version="2">
              <hp:script>[수식 스크립트]</hp:script>
            </hp:eqEdit>
          </hp:ctrl>
        </hp:run>
        """
        run = etree.SubElement(p_elem, _qn("hp", "run"))
        run.set("charPrIDRef", "0")

        ctrl = etree.SubElement(run, _qn("hp", "ctrl"))

        eq_edit = etree.SubElement(ctrl, _qn("hp", "eqEdit"))
        eq_edit.set("charPrIDRef", "0")
        eq_edit.set("version", "2")

        # 수식 크기 속성
        eq_sz = etree.SubElement(eq_edit, _qn("hp", "sz"))
        eq_sz.set("baseUnit", "1000")
        eq_sz.set("subscript", "500")
        eq_sz.set("superscript", "500")
        eq_sz.set("bigSymbol", "1500")
        eq_sz.set("bigSymbolSmall", "1000")

        script = etree.SubElement(eq_edit, _qn("hp", "script"))
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
