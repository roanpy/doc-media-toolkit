from __future__ import annotations

import argparse
from pathlib import Path

from .dependencies import dependency_statuses, missing_dependency_message
from .export_pipeline import export_document
from .models import DEFAULT_WATERMARK_TEXT, ExportOptions, WatermarkOptions
from .pptx_fonts import default_source_font_family
from .runtime_temp import cleanup_stale_runtime_entries
from pptx_tools.ui_theme import format_user_file_size


def format_file_size(size_bytes: int) -> str:
    return format_user_file_size(size_bytes)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Export DOCX/PDF/PPTX with optional watermarking, or watermark "
            "standalone image/video files directly. DOCX/PDF output is always PDF; "
            "PPTX can output PDF or PPTX; standalone media keeps its native output type."
        )
    )
    parser.add_argument(
        "input_document",
        type=Path,
        help="Input DOCX/PDF/PPTX/image/video path",
    )
    parser.add_argument(
        "--output-format",
        choices=["pptx", "pdf"],
        default="pptx",
        help="Output format",
    )
    parser.add_argument(
        "--output-mode",
        choices=["editable", "image"],
        default="editable",
        help="Editable keeps vector/text where possible; image flattens each page.",
    )
    parser.add_argument("--output", type=Path, help="Explicit output path")
    parser.add_argument("--watermark", action="store_true", help="Enable watermark")
    parser.add_argument(
        "--watermark-type",
        choices=["text", "image"],
        default="text",
        help="Watermark type",
    )
    parser.add_argument(
        "--watermark-text",
        default=DEFAULT_WATERMARK_TEXT,
        help="Watermark text",
    )
    parser.add_argument(
        "--watermark-image",
        type=Path,
        help="PNG image path for image watermark",
    )
    parser.add_argument(
        "--watermark-image-width",
        type=int,
        default=180,
        help="Image watermark width in pixels/points before rotation",
    )
    parser.add_argument(
        "--watermark-angle",
        type=float,
        default=315.0,
        help="Watermark angle in degrees",
    )
    parser.add_argument(
        "--watermark-color",
        default="#D9D9D9",
        help="Watermark hex color",
    )
    parser.add_argument(
        "--watermark-opacity",
        type=float,
        default=0.18,
        help="Watermark opacity between 0 and 1",
    )
    parser.add_argument(
        "--watermark-font-size",
        type=int,
        default=28,
        help="Watermark font size",
    )
    parser.add_argument(
        "--watermark-spacing",
        type=int,
        default=320,
        help="Watermark spacing",
    )
    parser.add_argument(
        "--watermark-margin",
        type=int,
        default=120,
        help="Extra margin between repeated watermarks",
    )
    parser.add_argument(
        "--no-watermark-bold",
        action="store_true",
        help="Disable the default bold watermark rendering.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=240,
        help="Render DPI for image mode",
    )
    parser.add_argument(
        "--jpeg-quality",
        type=int,
        default=85,
        help="JPEG quality for rendered page images",
    )
    parser.add_argument(
        "--image-pptx-keep-videos",
        action="store_true",
        help="For image-based PPTX export, reinsert embedded videos and watermark them with ffmpeg.",
    )
    parser.add_argument(
        "--video-encoder",
        choices=["auto", "cpu", "gpu"],
        default="auto",
        help="Video encoder for image-based PPTX video reinsertion. auto prefers GPU and falls back to CPU.",
    )
    parser.add_argument(
        "--keep-artifacts",
        action="store_true",
        help="Keep rendered page images and intermediate folders",
    )
    parser.add_argument(
        "--check-deps",
        action="store_true",
        help="Validate runtime dependencies for the selected export mode and exit.",
    )
    parser.add_argument(
        "--replace-source-fonts",
        action="store_true",
        help="Preprocess a temporary PPTX copy and replace explicit source fonts before export.",
    )
    parser.add_argument(
        "--source-font-family",
        default=None,
        help=(
            "Replacement family used with --replace-source-fonts. "
            "Defaults to a platform system font."
        ),
    )
    return parser.parse_args()


def main() -> int:
    cleanup_stale_runtime_entries()
    args = parse_args()
    if args.watermark and args.watermark_type == "image" and not args.watermark_image:
        raise SystemExit("--watermark-image is required when --watermark-type=image")
    options = ExportOptions(
        input_path=args.input_document,
        output_format=args.output_format,
        output_mode=args.output_mode,
        output_path=args.output,
        preserve_videos_in_image_pptx=args.image_pptx_keep_videos,
        video_encoder=args.video_encoder,
        dpi=args.dpi,
        jpeg_quality=args.jpeg_quality,
        keep_artifacts=args.keep_artifacts,
        replace_source_fonts=args.replace_source_fonts,
        source_font_family=args.source_font_family or default_source_font_family(),
        watermark=WatermarkOptions(
            enabled=args.watermark,
            kind=args.watermark_type,
            text=args.watermark_text,
            image_path=args.watermark_image,
            image_width=max(8, args.watermark_image_width),
            angle=args.watermark_angle,
            color=args.watermark_color,
            opacity=max(0.0, min(1.0, args.watermark_opacity)),
            font_size=max(6, args.watermark_font_size),
            spacing=max(40, args.watermark_spacing),
            margin=max(0, args.watermark_margin),
            bold=not args.no_watermark_bold,
        ),
    )
    if args.check_deps:
        statuses = dependency_statuses(options)
        for status in statuses:
            required = "required" if status.required else "optional"
            state = "ok" if status.available else "missing"
            path_suffix = f" path={status.path}" if status.path else ""
            detail_suffix = f" detail={status.detail}" if status.detail else ""
            print(f"{status.name}: {state} ({required}){path_suffix}{detail_suffix}")
        message = missing_dependency_message(options)
        return 1 if message else 0
    output_path = export_document(options)
    size_bytes = output_path.stat().st_size
    print(f"Output: {output_path}")
    print(f"Size: {format_file_size(size_bytes)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
