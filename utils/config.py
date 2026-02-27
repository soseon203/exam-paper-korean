import os
from pathlib import Path
from dotenv import load_dotenv

# 프로젝트 루트 디렉토리
PROJECT_ROOT = Path(__file__).parent.parent

# .env 파일 로드
load_dotenv(PROJECT_ROOT / ".env")


def get_api_key() -> str:
    """Anthropic API 키 반환."""
    key = os.getenv("ANTHROPIC_API_KEY", "")
    if not key or key == "your-api-key-here":
        raise ValueError(
            "ANTHROPIC_API_KEY가 설정되지 않았습니다. "
            ".env 파일에 API 키를 입력해주세요."
        )
    return key


def get_output_dir() -> Path:
    """기본 출력 디렉토리 반환."""
    output_dir = Path(os.getenv("OUTPUT_DIR", str(PROJECT_ROOT / "output")))
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


# Claude 모델 설정
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-6")
CLAUDE_MAX_TOKENS = int(os.getenv("CLAUDE_MAX_TOKENS", "8192"))

# PDF 변환 DPI
PDF_DPI = int(os.getenv("PDF_DPI", "300"))

# 이미지 최대 크기 (Claude Vision API 제한)
MAX_IMAGE_SIZE = int(os.getenv("MAX_IMAGE_SIZE", "4096"))

# ─── 이미지 품질 검사 임계값 ───────────────────────────────
# 최소 해상도 (가로 또는 세로가 이 값 미만이면 경고)
QC_MIN_WIDTH = int(os.getenv("QC_MIN_WIDTH", "500"))
QC_MIN_HEIGHT = int(os.getenv("QC_MIN_HEIGHT", "500"))

# 흐림(블러) 감지: Laplacian 분산이 이 값 미만이면 흐린 이미지로 판정
QC_BLUR_THRESHOLD = float(os.getenv("QC_BLUR_THRESHOLD", "100.0"))

# 빈 페이지 감지: 비백색 픽셀 비율(%)이 이 값 미만이면 빈 페이지로 판정
QC_BLANK_THRESHOLD = float(os.getenv("QC_BLANK_THRESHOLD", "1.0"))

# 대비 부족 감지: 히스토그램 표준편차가 이 값 미만이면 대비 부족
QC_CONTRAST_THRESHOLD = float(os.getenv("QC_CONTRAST_THRESHOLD", "30.0"))

# 품질 점수 최소 합격선 (0~100)
QC_PASS_SCORE = float(os.getenv("QC_PASS_SCORE", "40.0"))
