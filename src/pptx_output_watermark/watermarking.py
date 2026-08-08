from __future__ import annotations

import math
import os
import tempfile
from pathlib import Path

from PIL import Image, ImageColor, ImageDraw, ImageFont

from .font_assets import bundled_watermark_font_candidates
from .models import DEFAULT_WATERMARK_TEXT, WatermarkOptions

JPEG_OUTPUT_EXTENSIONS = {".jpg", ".jpeg", ".jpe"}
PNG_OUTPUT_EXTENSIONS = {".png"}
WEBP_OUTPUT_EXTENSIONS = {".webp"}


def render_rotation_angle(angle: float) -> float:
    """Convert UI clockwise angle to the image rendering direction."""
    return -float(angle)


def _bold_stroke_width(font_size: int) -> int:
    return max(1, min(2, font_size // 36))


def _load_font(font_size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    font_candidates = [
        *[str(path) for path in bundled_watermark_font_candidates()],
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "NotoSansSC[wght].ttf",
        "DejaVuSans.ttf",
        "arial.ttf",
    ]
    for candidate in font_candidates:
        try:
            return ImageFont.truetype(candidate, font_size)
        except Exception:
            continue
    return ImageFont.load_default()


def parse_hex_color(value: str) -> tuple[int, int, int]:
    rgb = ImageColor.getrgb(value)
    return rgb[:3]


def apply_image_opacity(image: Image.Image, opacity: float) -> Image.Image:
    rgba = image.convert("RGBA")
    alpha = rgba.getchannel("A")
    alpha = alpha.point(lambda value: int(value * max(0.0, min(1.0, opacity))))
    rgba.putalpha(alpha)
    return rgba


def _load_watermark_image(options: WatermarkOptions) -> Image.Image:
    if options.image_path is None:
        raise FileNotFoundError("Watermark image path is not set.")
    image_path = Path(options.image_path).expanduser()
    if not image_path.exists():
        raise FileNotFoundError(f"Watermark image not found: {image_path}")
    return Image.open(image_path).convert("RGBA")


def _build_text_watermark_overlay(
    size: tuple[int, int],
    options: WatermarkOptions,
) -> Image.Image:
    width, height = size
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    font = _load_font(options.font_size)
    alpha = max(0, min(255, int(round(options.opacity * 255))))
    color = (*parse_hex_color(options.color), alpha)
    text = options.text.strip() or DEFAULT_WATERMARK_TEXT
    stroke_width = _bold_stroke_width(options.font_size) if options.bold else 0

    bbox = draw.textbbox((0, 0), text, font=font, stroke_width=stroke_width)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    padding = max(4, stroke_width + 4)
    stamp_width = max(text_width + padding * 2, 1)
    stamp_height = max(text_height + padding * 2, 1)

    stamp = Image.new("RGBA", (stamp_width, stamp_height), (0, 0, 0, 0))
    stamp_draw = ImageDraw.Draw(stamp)
    stamp_draw.text(
        (padding - bbox[0], padding - bbox[1]),
        text,
        fill=color,
        font=font,
        stroke_width=stroke_width,
        stroke_fill=color,
    )
    rotated = stamp.rotate(
        render_rotation_angle(options.angle),
        expand=True,
        resample=Image.Resampling.BICUBIC,
    )

    step_x = max(int(options.spacing), rotated.width + int(options.margin), 1)
    step_y = max(int(options.spacing), rotated.height + int(options.margin), 1)
    for center_y in range(-step_y, height + step_y + 1, step_y):
        for center_x in range(-step_x, width + step_x + 1, step_x):
            overlay.alpha_composite(
                rotated,
                (center_x - rotated.width // 2, center_y - rotated.height // 2),
            )
    return overlay


def _build_image_watermark_overlay(
    size: tuple[int, int],
    options: WatermarkOptions,
) -> Image.Image:
    width, height = size
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    with _load_watermark_image(options) as source:
        image_width = max(8, int(options.image_width))
        ratio = image_width / max(1, source.width)
        image_height = max(1, int(round(source.height * ratio)))
        watermark = source.resize((image_width, image_height), Image.Resampling.LANCZOS)
        watermark = apply_image_opacity(watermark, options.opacity)

    rotated = watermark.rotate(
        render_rotation_angle(options.angle),
        expand=True,
        resample=Image.Resampling.BICUBIC,
    )

    step_x = max(int(options.spacing), rotated.width + int(options.margin), 1)
    step_y = max(int(options.spacing), rotated.height + int(options.margin), 1)
    for center_y in range(-step_y, height + step_y + 1, step_y):
        for center_x in range(-step_x, width + step_x + 1, step_x):
            overlay.alpha_composite(
                rotated,
                (center_x - rotated.width // 2, center_y - rotated.height // 2),
            )
    return overlay


def build_watermark_overlay(
    size: tuple[int, int],
    options: WatermarkOptions,
) -> Image.Image:
    if options.kind == "image":
        return _build_image_watermark_overlay(size, options)
    return _build_text_watermark_overlay(size, options)


def apply_watermark_to_image(
    image_path: Path,
    output_path: Path,
    options: WatermarkOptions,
    *,
    jpeg_quality: int = 95,
) -> Path:
    with Image.open(image_path) as base:
        rgba = base.convert("RGBA")
        overlay = build_watermark_overlay(rgba.size, options)
        combined = Image.alpha_composite(rgba, overlay)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        suffix = output_path.suffix.lower()
        normalized_quality = max(1, min(100, int(jpeg_quality)))
        if suffix in PNG_OUTPUT_EXTENSIONS:
            combined.save(output_path, format="PNG", optimize=True)
        elif suffix in WEBP_OUTPUT_EXTENSIONS:
            combined.convert("RGB").save(
                output_path,
                format="WEBP",
                quality=normalized_quality,
                method=6,
            )
        elif suffix in JPEG_OUTPUT_EXTENSIONS:
            combined.convert("RGB").save(
                output_path,
                format="JPEG",
                quality=normalized_quality,
                optimize=True,
                progressive=True,
            )
        else:
            combined.convert("RGB").save(
                output_path,
                format="JPEG",
                quality=normalized_quality,
                optimize=True,
                progressive=True,
            )
    return output_path


def write_watermark_overlay_image(
    size: tuple[int, int],
    options: WatermarkOptions,
) -> Path:
    fd, temp_name = tempfile.mkstemp(
        prefix="pptx_output_watermark_overlay_", suffix=".png"
    )
    os.close(fd)
    output = Path(temp_name)
    overlay = build_watermark_overlay(size, options)
    overlay.save(output, format="PNG")
    return output


def slide_emu_to_overlay_pixels(
    slide_width_emu: int, slide_height_emu: int
) -> tuple[int, int]:
    if slide_width_emu <= 0 or slide_height_emu <= 0:
        return (1920, 1080)
    scale = 1920.0 / float(slide_width_emu)
    return (
        max(1, int(math.ceil(slide_width_emu * scale))),
        max(1, int(math.ceil(slide_height_emu * scale))),
    )
