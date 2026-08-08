from __future__ import annotations

import html
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path


_TYPEFACE_RE = re.compile(rb'(?P<prefix>\btypeface=")(?P<value>[^"]*)(?P<suffix>")')


@dataclass(frozen=True, slots=True)
class FontScanResult:
    referenced: tuple[str, ...]
    missing: tuple[str, ...]


def default_source_font_family() -> str:
    if sys.platform == "darwin":
        return "PingFang SC"
    if sys.platform == "win32":
        return "Microsoft YaHei"
    return "Noto Sans CJK SC"


def _normalize_font_name(value: str) -> str:
    return " ".join(value.strip().strip("'\"").split()).casefold()


def _is_theme_font(value: str) -> bool:
    stripped = value.strip()
    return not stripped or stripped.startswith("+")


def _typeface_value(raw_value: bytes) -> str:
    return " ".join(html.unescape(raw_value.decode("utf-8", errors="ignore")).split())


def _extract_typefaces(data: bytes) -> set[str]:
    fonts: set[str] = set()
    for match in _TYPEFACE_RE.finditer(data):
        font_name = _typeface_value(match.group("value"))
        if _is_theme_font(font_name):
            continue
        fonts.add(font_name)
    return fonts


def _extract_theme_typefaces(data: bytes) -> set[str]:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return _extract_typefaces(data)

    fonts: set[str] = set()
    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name not in {"latin", "ea", "cs"}:
            continue
        font_name = " ".join(html.unescape(element.get("typeface", "")).split())
        if _is_theme_font(font_name):
            continue
        fonts.add(font_name)
    return fonts


def extract_pptx_fonts(pptx_path: Path) -> tuple[str, ...]:
    fonts: set[str] = set()
    with zipfile.ZipFile(pptx_path, "r") as package:
        for name in package.namelist():
            if not name.startswith("ppt/") or not name.endswith(".xml"):
                continue
            data = package.read(name)
            if name.startswith("ppt/theme/"):
                fonts.update(_extract_theme_typefaces(data))
            else:
                fonts.update(_extract_typefaces(data))
    return tuple(sorted(fonts, key=str.casefold))


def scan_missing_fonts(pptx_path: Path, available_families: set[str]) -> FontScanResult:
    referenced = extract_pptx_fonts(pptx_path)
    normalized_available = {_normalize_font_name(name) for name in available_families}
    missing = tuple(
        font_name
        for font_name in referenced
        if _normalize_font_name(font_name) not in normalized_available
    )
    return FontScanResult(referenced=referenced, missing=missing)


def replace_pptx_fonts(
    input_pptx: Path,
    output_pptx: Path,
    *,
    replacement_family: str,
    font_names: tuple[str, ...] = (),
) -> Path:
    targets = {_normalize_font_name(name) for name in font_names}
    replace_all = not targets
    replacement = html.escape(replacement_family, quote=True).encode("utf-8")

    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    with (
        zipfile.ZipFile(input_pptx, "r") as source,
        zipfile.ZipFile(
            output_pptx,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as target,
    ):
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename.startswith("ppt/") and item.filename.endswith(".xml"):
                data = _replace_typefaces(data, replacement, targets, replace_all)
            target.writestr(item, data)

    shutil.copystat(input_pptx, output_pptx, follow_symlinks=True)
    return output_pptx


def _replace_typefaces(
    data: bytes,
    replacement: bytes,
    targets: set[str],
    replace_all: bool,
) -> bytes:
    def repl(match: re.Match[bytes]) -> bytes:
        raw_value = match.group("value")
        font_name = html.unescape(raw_value.decode("utf-8", errors="ignore"))
        if _is_theme_font(font_name):
            return match.group(0)
        if replace_all or _normalize_font_name(font_name) in targets:
            return match.group("prefix") + replacement + match.group("suffix")
        return match.group(0)

    return _TYPEFACE_RE.sub(repl, data)
