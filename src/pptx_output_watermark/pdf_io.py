from __future__ import annotations

from pathlib import Path
from typing import IO, Any

from pypdf import PdfReader
from pypdf import filters as pypdf_filters


# LibreOffice can emit very large page/content streams for image-heavy decks.
# Recent pypdf releases cap declared stream lengths at 75 MB by default, which
# breaks otherwise valid editable-PDF watermark workflows on large presentations.
MAX_PDF_STREAM_LENGTH = 512_000_000


def configure_pypdf_limits() -> None:
    current_declared_limit = getattr(pypdf_filters, "MAX_DECLARED_STREAM_LENGTH", 0)
    current_array_limit = getattr(
        pypdf_filters, "MAX_ARRAY_BASED_STREAM_OUTPUT_LENGTH", 0
    )
    if current_declared_limit < MAX_PDF_STREAM_LENGTH:
        pypdf_filters.MAX_DECLARED_STREAM_LENGTH = MAX_PDF_STREAM_LENGTH
    if current_array_limit < MAX_PDF_STREAM_LENGTH:
        pypdf_filters.MAX_ARRAY_BASED_STREAM_OUTPUT_LENGTH = MAX_PDF_STREAM_LENGTH


def open_pdf_reader(stream: str | Path | IO[Any]) -> PdfReader:
    configure_pypdf_limits()
    return PdfReader(stream, strict=False)
