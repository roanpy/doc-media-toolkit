"""PDF-to-image rendering helpers using pypdfium2."""

from __future__ import annotations

from pathlib import Path
import math

import pypdfium2 as pdfium

from .image_utils import scale_image_to_limits
from .pdf_io import validate_pdf_page_size

DEFAULT_FULL_PAGE_RENDER_DPI = 240
DEFAULT_FULL_PAGE_JPEG_QUALITY = 85


def _bounded_render_scale(
    width: float,
    height: float,
    *,
    requested_scale: float,
    max_edge: int,
    max_pixels: int,
) -> float:
    validate_pdf_page_size(width, height)
    return min(
        requested_scale,
        max_edge / max(width, height),
        math.sqrt(max_pixels / (width * height)),
    )


def batch_render_pdf_slides(
    pdf_path: Path,
    *,
    num_slides: int,
    output_dir: Path,
    max_edge: int = 2048,
    max_pixels: int = 3_000_000,
    jpeg_quality: int = DEFAULT_FULL_PAGE_JPEG_QUALITY,
    dpi: int = DEFAULT_FULL_PAGE_RENDER_DPI,
) -> dict[int, Path]:
    """
    Renders pages of a PDF to JPEG images using pypdfium2.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    result: dict[int, Path] = {}

    scale = dpi / 72.0

    try:
        pdf = pdfium.PdfDocument(str(pdf_path))
    except Exception as e:
        raise RuntimeError(f"Failed to load PDF for rendering: {e}")

    pdf_page_count = len(pdf)
    render_count = min(num_slides, pdf_page_count)

    try:
        for page_index in range(render_count):
            page = pdf[page_index]

            width, height = page.get_size()
            page_scale = _bounded_render_scale(
                width,
                height,
                requested_scale=scale,
                max_edge=max_edge,
                max_pixels=max_pixels,
            )
            bitmap = page.render(
                scale=page_scale,
                rev_byteorder=False,  # RGB, not BGR
            )
            pil_img = bitmap.to_pil()

            normalized = (
                pil_img.convert("RGB") if pil_img.mode != "RGB" else pil_img.copy()
            )
            scaled = scale_image_to_limits(
                normalized,
                max_edge=max_edge,
                max_pixels=max_pixels,
            )

            final_path = output_dir / f"slide_{page_index:04d}.jpg"
            scaled.save(final_path, format="JPEG", quality=jpeg_quality)
            result[page_index] = final_path
    finally:
        pdf.close()

    if not result:
        raise RuntimeError("No rendered slide images were produced from PDF.")
    return result
