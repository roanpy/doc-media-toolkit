from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .process_utils import subprocess_text_kwargs

AUTOMATION_SETTINGS_URL = (
    "x-apple.systempreferences:com.apple.preference.security?Privacy_Automation"
)

_READINESS_APPLESCRIPT = """
on run argv
    if (count of argv) < 1 then
        error "Expected application name"
    end if
    set appName to item 1 of argv
    tell application appName
        run
        get name
    end tell
end run
"""


def _normalize_error(message: str) -> str:
    return " ".join(str(message or "").split()).strip()


def _is_permission_error(message: str) -> bool:
    lowered = message.lower()
    return "not authorized" in lowered or "not permitted" in lowered


def probe_app_automation(
    app_name: str,
    app_path: Path,
    *,
    timeout_seconds: float = 8.0,
) -> tuple[str, str]:
    if sys.platform != "darwin":
        return "unsupported", f"{app_name} automation is only supported on macOS."
    if not app_path.exists():
        return "missing", f"{app_name} is not installed."
    try:
        proc = subprocess.run(
            ["osascript", "-", app_name],
            input=_READINESS_APPLESCRIPT,
            capture_output=True,
            timeout=timeout_seconds,
            **subprocess_text_kwargs(),
        )
    except subprocess.TimeoutExpired:
        return (
            "error",
            f"{app_name} automation probe timed out after {timeout_seconds} seconds.",
        )
    except Exception as exc:
        return "error", f"{type(exc).__name__}: {exc}"

    if proc.returncode == 0:
        return "ready", ""
    message = _normalize_error(proc.stderr or proc.stdout)
    if _is_permission_error(message):
        return (
            "permission_denied",
            f"{message} Grant Automation permission for Doc Media Toolkit or Terminal to control {app_name} in System Settings.",
        )
    return (
        "error",
        f"osascript exit {proc.returncode}: {message or 'Unknown AppleScript error'}",
    )
