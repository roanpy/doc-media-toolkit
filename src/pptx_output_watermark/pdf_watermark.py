from __future__ import annotations

import io
from functools import lru_cache
from pathlib import Path

from pypdf import PdfWriter
from PIL import Image
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from .font_assets import bundled_watermark_font_candidates
from .models import DEFAULT_WATERMARK_TEXT, WatermarkOptions
from .pdf_io import open_pdf_reader
from .watermarking import apply_image_opacity, parse_hex_color, render_rotation_angle


def pdf_rotation_angle(angle: float) -> float:
    """Use the same visual direction as preview and image-based exports."""
    return render_rotation_angle(angle)


def _needs_unicode_font(text: str) -> bool:
    return any(ord(ch) > 127 for ch in text)


@lru_cache(maxsize=1)
def _register_unicode_font() -> str:
    font_candidates = [
        *[str(path) for path in bundled_watermark_font_candidates()],
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/msyh.ttf",
        "C:/Windows/Fonts/simhei.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/arphic/ukai.ttc",
    ]
    for font_path in font_candidates:
        if not Path(font_path).exists():
            continue
        try:
            font_name = "PPTXOutputWatermarkUnicode"
            pdfmetrics.registerFont(TTFont(font_name, font_path))
            return font_name
        except Exception:
            continue

    cid_font_name = "STSong-Light"
    try:
        pdfmetrics.getFont(cid_font_name)
    except KeyError:
        pdfmetrics.registerFont(UnicodeCIDFont(cid_font_name))
    return cid_font_name


def _resolve_font_name(text: str) -> str:
    if _needs_unicode_font(text):
        return _register_unicode_font()
    return "Helvetica"


def _draw_watermark_string(
    c: canvas.Canvas, text: str, font_size: int, bold: bool
) -> None:
    c.drawString(0, 0, text)
    if not bold:
        return

    offset = max(0.3, min(1.0, font_size / 56.0))
    c.drawString(offset, 0, text)
    c.drawString(0, offset, text)


def _build_watermark_pdf_bytes(
    *,
    width: float,
    height: float,
    options: WatermarkOptions,
) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    if options.kind == "image":
        _draw_image_watermarks(c, width, height, options)
        c.showPage()
        c.save()
        return buffer.getvalue()

    r, g, b = parse_hex_color(options.color)
    c.saveState()
    c.setFillColor(Color(r / 255.0, g / 255.0, b / 255.0, alpha=options.opacity))
    if hasattr(c, "setFillAlpha"):
        c.setFillAlpha(options.opacity)
    text = options.text.strip() or DEFAULT_WATERMARK_TEXT
    c.setFont(_resolve_font_name(text), max(6, int(options.font_size)))
    x = -width * 0.25
    while x < width * 1.25:
        y = -height * 0.25
        while y < height * 1.25:
            c.saveState()
            c.translate(x, y)
            c.rotate(pdf_rotation_angle(options.angle))
            _draw_watermark_string(
                c, text, max(6, int(options.font_size)), options.bold
            )
            c.restoreState()
            y += max(options.spacing, 80)
        x += max(options.spacing, 80)
    c.restoreState()
    c.showPage()
    c.save()
    return buffer.getvalue()


def _draw_image_watermarks(
    c: canvas.Canvas,
    width: float,
    height: float,
    options: WatermarkOptions,
) -> None:
    if options.image_path is None:
        raise FileNotFoundError("Watermark image path is not set.")
    image_path = Path(options.image_path).expanduser()
    if not image_path.exists():
        raise FileNotFoundError(f"Watermark image not found: {image_path}")

    with Image.open(image_path) as source:
        image_width = max(8.0, float(options.image_width))
        ratio = image_width / max(1, source.width)
        image_height = max(1.0, float(source.height) * ratio)
        resized = source.convert("RGBA").resize(
            (int(round(image_width)), int(round(image_height))),
            Image.Resampling.LANCZOS,
        )
        resized = apply_image_opacity(resized, options.opacity)
        png_buffer = io.BytesIO()
        resized.save(png_buffer, format="PNG")
    png_buffer.seek(0)
    reader = ImageReader(png_buffer)

    c.saveState()
    x = -width * 0.25
    while x < width * 1.25:
        y = -height * 0.25
        while y < height * 1.25:
            c.saveState()
            c.translate(x, y)
            c.rotate(pdf_rotation_angle(options.angle))
            c.drawImage(
                reader,
                -image_width / 2.0,
                -image_height / 2.0,
                width=image_width,
                height=image_height,
                mask="auto",
            )
            c.restoreState()
            y += max(float(options.spacing), image_height + 40.0)
        x += max(float(options.spacing), image_width + 40.0)
    c.restoreState()


def apply_watermark_to_pdf(
    input_pdf: Path,
    output_pdf: Path,
    options: WatermarkOptions,
) -> Path:
    reader = open_pdf_reader(str(input_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_reader = open_pdf_reader(
            io.BytesIO(
                _build_watermark_pdf_bytes(width=width, height=height, options=options)
            )
        )
        page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)
    with open(output_pdf, "wb") as handle:
        writer.write(handle)
    return output_pdf
