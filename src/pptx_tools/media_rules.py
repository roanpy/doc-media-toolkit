from __future__ import annotations

import re
from pathlib import Path


_LEADING_EPOCH = re.compile(r"^(\d{10}|\d{13})[\s._-]+(.+)$")
_TRAILING_FAULT = re.compile(r"[\s._-]+fault$", re.IGNORECASE)
_CJK = re.compile(r"[\u3400-\u9fff]")


def safe_media_name(value: str, fallback: str = "media") -> str:
    """Return the existing cross-platform library name format."""
    value = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("._")
    return value[:80] or fallback


def normalize_import_name(value: str, fallback: str = "media") -> str:
    """Clean only unambiguous machine-generated noise from a new asset name."""
    name = value.strip()
    timestamp = ""
    match = _LEADING_EPOCH.fullmatch(name)
    if match:
        timestamp, name = match.groups()
    if _TRAILING_FAULT.search(name):
        status = "异常" if _CJK.search(name) else "fault"
        name = f"{_TRAILING_FAULT.sub('', name)}_{status}"
    name = safe_media_name(name, fallback)
    if timestamp:
        suffix = f"_{timestamp}"
        return f"{name[: 80 - len(suffix)]}{suffix}"
    return name


def normalize_media_category(value: str) -> Path:
    """Validate a portable, library-relative category path."""
    raw = value.strip().replace("\\", "/")
    if not raw:
        return Path()
    if raw.startswith("/") or re.match(r"^[A-Za-z]:", raw):
        raise ValueError("Media library category must be a relative folder path")
    raw = raw.rstrip("/")
    if any(part in {"", ".", ".."} for part in raw.split("/")):
        raise ValueError("Media library category must be a relative folder path")
    parts = [safe_media_name(part, "") for part in raw.split("/")]
    if any(not part for part in parts):
        raise ValueError("Media library category contains an invalid folder name")
    return Path(*parts)
