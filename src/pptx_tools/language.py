"""Small, shared language-selection helper for the desktop workspaces."""

from __future__ import annotations

import os
import re
import subprocess
import sys

from PySide6.QtCore import QLocale


def _macos_preferred_language() -> str:
    if sys.platform != "darwin":
        return ""
    try:
        result = subprocess.run(
            ["/usr/bin/defaults", "read", "-g", "AppleLanguages"],
            check=False,
            capture_output=True,
            text=True,
            timeout=2,
        )
    except (OSError, subprocess.SubprocessError):
        return ""
    if result.returncode != 0:
        return ""
    quoted = re.findall(r'"([^"\r\n]+)"', result.stdout)
    if quoted:
        return quoted[0]
    for line in result.stdout.splitlines():
        candidate = line.strip().strip(",")
        if re.fullmatch(r"[A-Za-z]{2,3}(?:[-_][A-Za-z0-9]+)*", candidate):
            return candidate
    return ""


def detect_language(environment_variable: str) -> str:
    """Return the explicit override, otherwise the system UI language.

    The application currently ships Simplified Chinese and English.  Any
    unsupported system locale intentionally falls back to English.
    """

    override = os.environ.get(environment_variable, "").strip().lower()
    if override.startswith("zh"):
        return "zh"
    if override.startswith("en"):
        return "en"
    system_locale = (
        (_macos_preferred_language() or QLocale.system().name()).strip().lower()
    )
    return "zh" if system_locale.startswith("zh") else "en"
