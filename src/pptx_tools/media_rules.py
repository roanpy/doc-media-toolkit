from __future__ import annotations

import re
from pathlib import Path


_LEADING_EPOCH = re.compile(r"^(\d{10}|\d{13})[\s._-]+(.+)$")
_TRAILING_FAULT = re.compile(r"[\s._-]+fault$", re.IGNORECASE)
_CJK = re.compile(r"[\u3400-\u9fff]")
_WINDOWS_RESERVED = re.compile(
    r"^(?:con|prn|aux|nul|com[1-9]|lpt[1-9]|com[¹²³]|lpt[¹²³])$",
    re.IGNORECASE,
)
MEDIA_NAME_MAX_LENGTH = 80
MEDIA_COMPONENT_MAX_BYTES = 255


def _truncate_utf8(value: str, max_bytes: int) -> str:
    """Truncate without splitting a Unicode code point."""
    if max_bytes <= 0:
        return ""
    if len(value.encode("utf-8")) <= max_bytes:
        return value
    encoded = value.encode("utf-8")[:max_bytes]
    return encoded.decode("utf-8", errors="ignore")


def safe_media_name(value: str, fallback: str = "media") -> str:
    """Return the existing cross-platform library name format."""
    clean = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("._")
    if not clean:
        if fallback == "":
            return ""
        clean = re.sub(r"[^\w.-]+", "_", fallback, flags=re.UNICODE).strip("._")
    if not clean:
        clean = "media"
    # Windows reserves these names even when an extension follows (CON.txt).
    device = clean.split(".", 1)[0].rstrip(" .")
    if _WINDOWS_RESERVED.fullmatch(device):
        clean = f"_{clean}"
    clean = _truncate_utf8(clean[:MEDIA_NAME_MAX_LENGTH], MEDIA_COMPONENT_MAX_BYTES)
    return clean.rstrip("._") or ("" if fallback == "" else "media")


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
        return fit_media_component(name[: MEDIA_NAME_MAX_LENGTH - len(suffix)], suffix)
    return name


def fit_media_component(value: str, suffix: str = "") -> str:
    """Fit a generated filename component while keeping its metadata suffix."""
    budget = MEDIA_COMPONENT_MAX_BYTES - len(suffix.encode("utf-8"))
    return f"{_truncate_utf8(value, budget)}{suffix}"


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
