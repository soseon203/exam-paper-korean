"""PDF를 페이지별 이미지로 변환하는 모듈."""

from __future__ import annotations

import io
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image

from utils.config import PDF_DPI, MAX_IMAGE_SIZE


def pdf_to_images(pdf_path: str | Path) -> list[Image.Image]:
    """PDF 파일을 페이지별 PIL Image 리스트로 변환.

    Args:
        pdf_path: PDF 파일 경로

    Returns:
        페이지별 PIL Image 리스트
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

    doc = fitz.open(str(pdf_path))
    images = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        # 300 DPI로 렌더링
        zoom = PDF_DPI / 72  # 72 DPI가 기본
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        img = _resize_if_needed(img)
        images.append(img)

    doc.close()
    return images


def load_image(image_path: str | Path) -> Image.Image:
    """이미지 파일을 PIL Image로 로드.

    Args:
        image_path: 이미지 파일 경로

    Returns:
        PIL Image
    """
    image_path = Path(image_path)
    if not image_path.exists():
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image_path}")

    img = Image.open(str(image_path))
    if img.mode != "RGB":
        img = img.convert("RGB")
    img = _resize_if_needed(img)
    return img


def _resize_if_needed(img: Image.Image) -> Image.Image:
    """이미지가 최대 크기를 초과하면 리사이즈."""
    w, h = img.size
    if w > MAX_IMAGE_SIZE or h > MAX_IMAGE_SIZE:
        ratio = min(MAX_IMAGE_SIZE / w, MAX_IMAGE_SIZE / h)
        new_size = (int(w * ratio), int(h * ratio))
        img = img.resize(new_size, Image.LANCZOS)
    return img


def image_to_base64(img: Image.Image, format: str = "PNG") -> str:
    """PIL Image를 base64 문자열로 변환."""
    import base64

    buffer = io.BytesIO()
    img.save(buffer, format=format)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def get_supported_extensions() -> set[str]:
    """지원하는 파일 확장자 집합 반환."""
    return {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}


def is_pdf(path: str | Path) -> bool:
    """PDF 파일 여부 확인."""
    return Path(path).suffix.lower() == ".pdf"
