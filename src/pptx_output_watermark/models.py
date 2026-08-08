from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .pptx_fonts import default_source_font_family

DEFAULT_WATERMARK_TEXT = "企业专属，注意保密"


@dataclass(slots=True)
class WatermarkOptions:
    enabled: bool = False
    kind: str = "text"
    text: str = DEFAULT_WATERMARK_TEXT
    image_path: Path | None = None
    image_width: int = 180
    angle: float = 315.0
    color: str = "#D9D9D9"
    opacity: float = 0.18
    font_size: int = 28
    spacing: int = 320
    margin: int = 120
    bold: bool = True


@dataclass(slots=True)
class ExportOptions:
    input_path: Path
    output_format: str = "pptx"
    output_mode: str = "editable"
    output_path: Path | None = None
    preserve_videos_in_image_pptx: bool = False
    video_encoder: str = "auto"
    video_quality_profile: str = "high"
    dpi: int = 240
    jpeg_quality: int = 85
    keep_artifacts: bool = False
    replace_source_fonts: bool = False
    source_font_family: str = field(default_factory=default_source_font_family)
    source_font_names: tuple[str, ...] = ()
    watermark: WatermarkOptions = field(default_factory=WatermarkOptions)


@dataclass(slots=True)
class ExportArtifacts:
    temp_pdf: Path | None = None
    rendered_dir: Path | None = None
    watermarked_dir: Path | None = None
    preprocessed_dir: Path | None = None
    video_work_dir: Path | None = None
