from __future__ import annotations

from pathlib import Path
from typing import IO, Any

from pypdf import PdfReader
from pypdf import filters as pypdf_filters


# LibreOffice can emit very large page/content streams for image-heavy decks.
# Recent pypdf releases cap declared stream lengths at 75 MB by default, which
# breaks otherwise valid editable-PDF watermark workflows on large presentations.
MAX_PDF_STREAM_LENGTH = 512_000_000
MAX_PDF_PAGE_POINTS = 14_400.0


def configure_pypdf_limits() -> None:
    current_declared_limit = getattr(pypdf_filters, "MAX_DECLARED_STREAM_LENGTH", 0)
    if current_declared_limit < MAX_PDF_STREAM_LENGTH:
        pypdf_filters.MAX_DECLARED_STREAM_LENGTH = MAX_PDF_STREAM_LENGTH


def validate_pdf_page_size(width: float, height: float) -> None:
    if (
        width <= 0
        or height <= 0
        or width > MAX_PDF_PAGE_POINTS
        or height > MAX_PDF_PAGE_POINTS
    ):
        raise ValueError(f"Unsupported PDF page size: {width:g} x {height:g} points")


def open_pdf_reader(stream: str | Path | IO[Any]) -> PdfReader:
    configure_pypdf_limits()
    return PdfReader(stream, strict=False)
