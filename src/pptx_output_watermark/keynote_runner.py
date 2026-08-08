"""Keynote AppleScript helpers for macOS."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .app_circuit import conversion_disabled as circuit_disabled
from .app_circuit import disable_conversion_temporarily as disable_circuit_temporarily
from .mac_automation import AUTOMATION_SETTINGS_URL, probe_app_automation
from .process_utils import run_process, subprocess_text_kwargs

KEYNOTE_APP_PATH = Path("/Applications/Keynote.app")
KEYNOTE_APP_STORE_URL = "macappstore://apps.apple.com/app/keynote/id409183694"
KEYNOTE_AUTOMATION_SETTINGS_URL = AUTOMATION_SETTINGS_URL
KEYNOTE_CIRCUIT_KEY = "keynote"

KEYNOTE_APPLESCRIPT = """
on run argv
    if (count of argv) < 2 then
        error "Expected 2 arguments: input_path output_path"
    end if
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    set inFile to POSIX file inPath
    set outFile to POSIX file outPath
    set myDoc to missing value

    tell application "Keynote"
        run
        try
            set myDoc to open inFile
            export myDoc to outFile as PDF
            close myDoc saving no
        on error errMsg number errNum
            try
                if myDoc is not missing value then close myDoc saving no
            end try
            error errMsg number errNum
        end try
    end tell
end run
"""


def keynote_available() -> bool:
    return sys.platform == "darwin" and KEYNOTE_APP_PATH.exists()


def keynote_automation_status() -> tuple[str, str]:
    return probe_app_automation("Keynote", KEYNOTE_APP_PATH)


def keynote_conversion_disabled() -> tuple[bool, str]:
    return circuit_disabled(KEYNOTE_CIRCUIT_KEY)


def disable_keynote_temporarily(reason: str) -> None:
    disable_circuit_temporarily(KEYNOTE_CIRCUIT_KEY, reason)


def _format_keynote_error(message: str) -> str:
    normalized = " ".join(str(message or "").split())
    lowered = normalized.lower()
    if "not authorized" in lowered or "not permitted" in lowered:
        return (
            f"{normalized} Grant Automation permission for Doc Media Toolkit or Terminal "
            "to control Keynote in System Settings."
        )
    return normalized or "Unknown AppleScript error"


def convert_via_keynote(
    input_pptx: Path, output_pdf: Path, timeout_seconds: float = 120.0
) -> tuple[bool, str]:
    """Convert a PPTX file to PDF using Keynote via AppleScript.

    Returns:
        tuple[bool, str]: (Success, detail or error message)
    """
    if sys.platform != "darwin":
        return False, "Keynote conversion is only supported on macOS."
    if not keynote_available():
        return False, "Keynote is not installed."
    disabled, disabled_reason = keynote_conversion_disabled()
    if disabled:
        return False, f"Keynote conversion is temporarily disabled: {disabled_reason}"

    try:
        input_str = str(input_pptx.resolve())
        output_str = str(output_pdf.resolve())

        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        output_pdf.unlink(missing_ok=True)

        proc = run_process(
            ["osascript", "-", input_str, output_str],
            input=KEYNOTE_APPLESCRIPT,
            capture_output=True,
            timeout=timeout_seconds,
            **subprocess_text_kwargs(),
        )

        if (
            proc.returncode == 0
            and output_pdf.exists()
            and output_pdf.stat().st_size > 0
        ):
            return True, ""

        error_msg = _format_keynote_error(proc.stderr or proc.stdout)
        disable_keynote_temporarily(error_msg)
        return False, f"osascript exit {proc.returncode}: {error_msg}"

    except subprocess.TimeoutExpired:
        reason = f"Keynote conversion timed out after {timeout_seconds} seconds."
        disable_keynote_temporarily(reason)
        return False, reason
    except Exception as exc:
        reason = f"{type(exc).__name__}: {exc}"
        disable_keynote_temporarily(reason)
        return False, reason
