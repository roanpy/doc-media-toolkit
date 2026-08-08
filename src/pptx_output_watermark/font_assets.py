from __future__ import annotations

import sys
from pathlib import Path

BUNDLED_WATERMARK_FONT_FILE = "NotoSansSC[wght].ttf"
BUNDLED_WATERMARK_FONT_FAMILY = "Noto Sans SC"


def resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    source_root = Path(__file__).resolve().parents[2]
    return source_root if (source_root / "assets").is_dir() else Path(sys.prefix)


def bundled_watermark_font_path() -> Path:
    return resource_root() / "assets" / "fonts" / BUNDLED_WATERMARK_FONT_FILE


def bundled_watermark_font_candidates() -> list[Path]:
    bundled = bundled_watermark_font_path()
    return [bundled] if bundled.exists() else []
