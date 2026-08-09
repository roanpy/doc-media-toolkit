"""Small, shared language-selection helper for the desktop workspaces."""

from __future__ import annotations

import os

from PySide6.QtCore import QLocale


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
    system_locale = QLocale.system().name().strip().lower()
    return "zh" if system_locale.startswith("zh") else "en"
