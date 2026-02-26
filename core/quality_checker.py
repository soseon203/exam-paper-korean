"""이미지 품질 사전 검사 모듈.

API 호출 전에 이미지 품질을 검사하여 흐린 스캔, 빈 페이지 등
불필요한 과금을 방지합니다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
from PIL import Image

from utils.config import (
    QC_MIN_WIDTH,
    QC_MIN_HEIGHT,
    QC_BLUR_THRESHOLD,
    QC_BLANK_THRESHOLD,
    QC_CONTRAST_THRESHOLD,
    QC_PASS_SCORE,
)


@dataclass
class ImageQuality:
    """이미지 품질 검사 결과."""

    score: float = 100.0          # 0~100 품질 점수
    passed: bool = True           # 합격 여부
    warnings: list[str] = field(default_factory=list)

    # 세부 측정값
    width: int = 0
    height: int = 0
    blur_score: float = 0.0       # Laplacian 분산 (높을수록 선명)
    blank_ratio: float = 0.0      # 비백색 픽셀 비율 (%)
    contrast_std: float = 0.0     # 히스토그램 표준편차


def check_image_quality(image: Image.Image) -> ImageQuality:
    """이미지 품질을 종합 검사.

    Args:
        image: 검사할 PIL Image

    Returns:
        ImageQuality 결과
    """
    result = ImageQuality()
    arr = np.array(image.convert("RGB"))

    # 1) 해상도 체크
    result.width, result.height = image.size
    _check_resolution(result)

    # 2) 그레이스케일 변환 (블러/대비 측정용)
    gray = np.array(image.convert("L"), dtype=np.float64)

    # 3) 흐림 감지
    result.blur_score = _compute_laplacian_variance(gray)
    _check_blur(result)

    # 4) 빈 페이지 감지
    result.blank_ratio = _compute_non_white_ratio(arr)
    _check_blank(result)

    # 5) 대비 부족 감지
    result.contrast_std = float(np.std(gray))
    _check_contrast(result)

    # 최종 합격 판정
    result.passed = result.score >= QC_PASS_SCORE
    return result


# ─── 개별 검사 함수 ──────────────────────────────────────────


def _check_resolution(result: ImageQuality) -> None:
    if result.width < QC_MIN_WIDTH or result.height < QC_MIN_HEIGHT:
        penalty = 30.0
        result.score -= penalty
        result.warnings.append(
            f"해상도 부족: {result.width}x{result.height} "
            f"(최소 {QC_MIN_WIDTH}x{QC_MIN_HEIGHT})"
        )


def _compute_laplacian_variance(gray: np.ndarray) -> float:
    """Laplacian 필터의 분산으로 선명도 측정."""
    # 간단한 3x3 Laplacian 커널 적용 (scipy/cv2 없이 구현)
    kernel = np.array([[0, 1, 0],
                       [1, -4, 1],
                       [0, 1, 0]], dtype=np.float64)
    h, w = gray.shape
    # 패딩된 이미지
    padded = np.pad(gray, 1, mode="edge")
    laplacian = np.zeros_like(gray)
    for dy in range(3):
        for dx in range(3):
            laplacian += kernel[dy, dx] * padded[dy:dy + h, dx:dx + w]
    return float(np.var(laplacian))


def _check_blur(result: ImageQuality) -> None:
    if result.blur_score < QC_BLUR_THRESHOLD:
        penalty = 40.0
        result.score -= penalty
        result.warnings.append(
            f"이미지가 흐림: 선명도 {result.blur_score:.1f} "
            f"(기준 {QC_BLUR_THRESHOLD})"
        )


def _compute_non_white_ratio(arr: np.ndarray) -> float:
    """비백색 픽셀 비율(%) 계산."""
    # 각 픽셀의 밝기 (R+G+B)/3
    brightness = arr.mean(axis=2)
    # 240 이상을 '백색'으로 간주
    non_white = np.sum(brightness < 240)
    total = brightness.size
    return float(non_white / total * 100)


def _check_blank(result: ImageQuality) -> None:
    if result.blank_ratio < QC_BLANK_THRESHOLD:
        penalty = 50.0
        result.score -= penalty
        result.warnings.append(
            f"빈 페이지 의심: 내용 비율 {result.blank_ratio:.2f}% "
            f"(기준 {QC_BLANK_THRESHOLD}%)"
        )


def _check_contrast(result: ImageQuality) -> None:
    if result.contrast_std < QC_CONTRAST_THRESHOLD:
        penalty = 25.0
        result.score -= penalty
        result.warnings.append(
            f"대비 부족: 표준편차 {result.contrast_std:.1f} "
            f"(기준 {QC_CONTRAST_THRESHOLD})"
        )
