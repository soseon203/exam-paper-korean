"""HWPX 템플릿 로더.

HWPX ZIP 파일에서 페이지 설정과 스타일 정보를 추출합니다.
복사 기반 접근: 템플릿 HWPX를 통째로 복사하여 section0.xml 내용만 교체하므로,
header.xml의 글꼴·스타일 정의가 자동으로 보존됩니다.

.hwp 파일이 주어지면 한글 COM 자동화로 .hwpx로 변환 후 로드합니다.
"""

from __future__ import annotations

import logging
import os
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from models.template_config import TemplateConfig

logger = logging.getLogger(__name__)

# HWPX 네임스페이스
NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
}

# 템플릿에서 필수로 존재해야 하는 파일
REQUIRED_FILES = {"Contents/header.xml", "Contents/section0.xml"}


def load_template(hwpx_path: str | Path) -> TemplateConfig:
    """HWPX 또는 HWP 템플릿 파일을 로드하여 TemplateConfig를 반환.

    .hwp 파일이 주어지면 한글 COM 자동화로 .hwpx로 변환 후 로드합니다.

    Args:
        hwpx_path: HWPX 또는 HWP 템플릿 파일 경로

    Returns:
        추출된 TemplateConfig

    Raises:
        FileNotFoundError: 파일이 없을 때
        ValueError: 유효하지 않은 파일일 때
        RuntimeError: HWP 변환 실패 시
    """
    hwpx_path = Path(hwpx_path)
    if not hwpx_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {hwpx_path}")

    ext = hwpx_path.suffix.lower()
    if ext == ".hwp":
        # HWP → HWPX 변환
        hwpx_path = convert_hwp_to_hwpx(hwpx_path)
    elif ext != ".hwpx":
        raise ValueError(f"지원하지 않는 파일 형식입니다: {ext} (.hwp 또는 .hwpx만 가능)")

    config = TemplateConfig(source_path=hwpx_path)

    try:
        with zipfile.ZipFile(str(hwpx_path), "r") as zf:
            # 필수 파일 존재 확인
            zip_names = set(zf.namelist())
            missing = REQUIRED_FILES - zip_names
            if missing:
                raise ValueError(
                    f"필수 파일이 없습니다: {', '.join(missing)}"
                )

            # header.xml 추출 (글꼴·스타일 정의 전체)
            _extract_header(zf, config)

            # section0.xml에서 페이지 설정(secPr) 추출
            _extract_section_settings(zf, config)

    except zipfile.BadZipFile:
        raise ValueError(f"유효하지 않은 ZIP/HWPX 파일입니다: {hwpx_path}")

    if not config.is_valid:
        reasons = []
        if not config.has_valid_header:
            reasons.append("header.xml 파싱 실패")
        if not config.has_valid_section:
            reasons.append("section0.xml 페이지 설정 없음")
        raise ValueError(
            f"유효하지 않은 템플릿: {'; '.join(reasons)}"
        )

    logger.info("템플릿 로드 완료: %s", config.summary)
    return config


def _extract_header(zf: zipfile.ZipFile, config: TemplateConfig):
    """header.xml에서 글꼴·스타일 정보 추출.

    header.xml 전체를 바이트로 보존하여 나중에 그대로 교체할 수 있도록 합니다.
    이렇게 하면 fontfaces, charProperties, paraProperties, styles가
    모두 자동으로 유지됩니다.
    """
    try:
        header_bytes = zf.read("Contents/header.xml")
        # 파싱 가능한지 검증
        root = etree.fromstring(header_bytes)

        # 최소한의 구조 확인: head 요소가 존재하는지
        if root.tag and "head" in root.tag.lower() or len(root) > 0:
            config.header_xml_bytes = header_bytes
            config.has_valid_header = True
            logger.debug("header.xml 추출 완료 (%d bytes)", len(header_bytes))
        else:
            logger.warning("header.xml 구조가 예상과 다릅니다")

    except Exception as e:
        logger.warning("header.xml 추출 실패: %s", e)


def _extract_section_settings(zf: zipfile.ZipFile, config: TemplateConfig):
    """section0.xml에서 secPr(페이지 설정) 추출."""
    try:
        section_bytes = zf.read("Contents/section0.xml")
        root = etree.fromstring(section_bytes)

        # secPr 요소 찾기 (여러 네임스페이스 경로 시도)
        sec_pr = _find_sec_pr(root)
        if sec_pr is None:
            logger.warning("section0.xml에서 secPr을 찾을 수 없습니다")
            return

        config.sec_pr_xml = etree.tostring(sec_pr, encoding="unicode").encode("utf-8")
        config.has_valid_section = True

        # 페이지 크기·여백 추출 (있으면)
        _parse_page_properties(sec_pr, config)

        logger.debug("secPr 추출 완료")

    except Exception as e:
        logger.warning("section0.xml 파싱 실패: %s", e)


def _find_sec_pr(root: etree._Element) -> etree._Element | None:
    """secPr 요소를 다양한 경로에서 검색."""
    # 직접 검색
    for ns_prefix in ("hp", "hs"):
        ns_uri = NS.get(ns_prefix, "")
        sec_pr = root.find(f".//{{{ns_uri}}}secPr")
        if sec_pr is not None:
            return sec_pr

    # 첫 번째 문단 내에서 검색
    for ns_prefix in ("hp", "hs"):
        ns_uri = NS.get(ns_prefix, "")
        for p in root.findall(f".//{{{ns_uri}}}p"):
            sec_pr = p.find(f".//{{{ns_uri}}}secPr")
            if sec_pr is not None:
                return sec_pr

    # 와일드카드 검색
    for elem in root.iter():
        if elem.tag and "secPr" in elem.tag:
            return elem

    return None


def _parse_page_properties(sec_pr: etree._Element, config: TemplateConfig):
    """secPr에서 페이지 크기·여백 정보 추출."""
    # pagePr 찾기
    page_pr = None
    for elem in sec_pr.iter():
        if elem.tag and "pagePr" in elem.tag:
            page_pr = elem
            break

    if page_pr is None:
        return

    # 속성 추출
    config.landscape = page_pr.get("landscape", "0") == "1"

    # 페이지 크기
    for sz in page_pr.iter():
        if sz.tag and "pageSz" in sz.tag:
            w = sz.get("width") or sz.get("w")
            h = sz.get("height") or sz.get("h")
            if w:
                config.page_width = int(w)
            if h:
                config.page_height = int(h)
            break

    # 여백
    for margin in page_pr.iter():
        if margin.tag and "pageMargin" in margin.tag:
            for attr, field_name in [
                ("top", "margin_top"),
                ("bottom", "margin_bottom"),
                ("left", "margin_left"),
                ("right", "margin_right"),
                ("header", "margin_header"),
                ("footer", "margin_footer"),
            ]:
                val = margin.get(attr)
                if val:
                    setattr(config, field_name, int(val))
            break


# ─── HWP → HWPX 변환 (한글 COM 자동화) ─────────────────────


def convert_hwp_to_hwpx(hwp_path: str | Path) -> Path:
    """한글 COM 자동화를 사용하여 .hwp를 .hwpx로 변환.

    한컴오피스가 설치된 Windows에서만 동작합니다.

    Args:
        hwp_path: 변환할 .hwp 파일 경로

    Returns:
        변환된 .hwpx 파일 경로

    Raises:
        RuntimeError: 한글이 설치되지 않았거나 변환 실패 시
    """
    if sys.platform != "win32":
        raise RuntimeError("HWP → HWPX 변환은 Windows에서만 가능합니다.")

    hwp_path = Path(hwp_path).resolve()
    if not hwp_path.exists():
        raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {hwp_path}")

    # 변환된 hwpx를 원본 옆에 저장
    hwpx_path = hwp_path.with_suffix(".hwpx")

    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32가 설치되어 있지 않습니다.\n"
            "설치: pip install pywin32"
        )

    hwp = None
    try:
        logger.info("한글 COM 객체 생성 중...")
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")

        # 보안 모듈 등록 (자동화 허용)
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

        # 백그라운드 실행 (창 숨기기)
        hwp.XHwpWindows.Item(0).Visible = False

        # HAction 방식으로 파일 열기
        logger.info("HWP 파일 열기: %s", hwp_path)
        hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(hwp_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

        # HAction 방식으로 HWPX 저장
        logger.info("HWPX로 저장: %s", hwpx_path)
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(hwpx_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        logger.info("HWP → HWPX 변환 완료: %s", hwpx_path)

    except Exception as e:
        if "HWPFrame.HwpObject" in str(e) or "Class not registered" in str(e):
            raise RuntimeError(
                "한컴오피스(한글)가 설치되어 있지 않습니다.\n"
                "한글이 설치된 PC에서 실행하거나, "
                "한글에서 직접 .hwpx로 저장해주세요."
            )
        raise RuntimeError(f"HWP → HWPX 변환 실패: {e}")

    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)  # 변경사항 무시하고 닫기
                hwp.Quit()
            except Exception:
                pass

    if not hwpx_path.exists():
        raise RuntimeError("변환된 HWPX 파일이 생성되지 않았습니다.")

    return hwpx_path
