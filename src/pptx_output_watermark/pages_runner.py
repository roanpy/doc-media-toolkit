"""Pages AppleScript helpers for macOS DOCX fallback."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .app_circuit import conversion_disabled as circuit_disabled
from .app_circuit import disable_conversion_temporarily as disable_circuit_temporarily
from .mac_automation import AUTOMATION_SETTINGS_URL, probe_app_automation
from .process_utils import run_process, subprocess_text_kwargs

PAGES_APP_PATH = Path("/Applications/Pages.app")
PAGES_APP_STORE_URL = "macappstore://apps.apple.com/app/pages/id409201541"
PAGES_AUTOMATION_SETTINGS_URL = AUTOMATION_SETTINGS_URL
PAGES_CIRCUIT_KEY = "pages"

PAGES_APPLESCRIPT = """
on run argv
    if (count of argv) < 2 then
        error "Expected 2 arguments: input_path output_path"
    end if
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    set inFile to POSIX file inPath
    set outFile to POSIX file outPath
    set myDoc to missing value

    tell application "Pages"
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


def pages_available() -> bool:
    return sys.platform == "darwin" and PAGES_APP_PATH.exists()


def pages_automation_status() -> tuple[str, str]:
    return probe_app_automation("Pages", PAGES_APP_PATH)


def pages_conversion_disabled() -> tuple[bool, str]:
    return circuit_disabled(PAGES_CIRCUIT_KEY)


def disable_pages_temporarily(reason: str) -> None:
    disable_circuit_temporarily(PAGES_CIRCUIT_KEY, reason)


def _format_pages_error(message: str) -> str:
    normalized = " ".join(str(message or "").split())
    lowered = normalized.lower()
    if "not authorized" in lowered or "not permitted" in lowered:
        return (
            f"{normalized} Grant Automation permission for Doc Media Toolkit or Terminal "
            "to control Pages in System Settings."
        )
    return normalized or "Unknown AppleScript error"


def convert_via_pages(
    input_docx: Path, output_pdf: Path, timeout_seconds: float = 120.0
) -> tuple[bool, str]:
    """Convert a DOCX file to PDF using Pages via AppleScript."""
    if sys.platform != "darwin":
        return False, "Pages conversion is only supported on macOS."
    if input_docx.suffix.lower() != ".docx":
        return False, "Pages fallback only supports DOCX input."
    if not pages_available():
        return False, "Pages is not installed."
    disabled, disabled_reason = pages_conversion_disabled()
    if disabled:
        return False, f"Pages conversion is temporarily disabled: {disabled_reason}"

    try:
        input_str = str(input_docx.resolve())
        output_str = str(output_pdf.resolve())

        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        output_pdf.unlink(missing_ok=True)

        proc = run_process(
            ["osascript", "-", input_str, output_str],
            input=PAGES_APPLESCRIPT,
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

        error_msg = _format_pages_error(proc.stderr or proc.stdout)
        disable_pages_temporarily(error_msg)
        return False, f"osascript exit {proc.returncode}: {error_msg}"

    except subprocess.TimeoutExpired:
        reason = f"Pages conversion timed out after {timeout_seconds} seconds."
        disable_pages_temporarily(reason)
        return False, reason
    except Exception as exc:
        reason = f"{type(exc).__name__}: {exc}"
        disable_pages_temporarily(reason)
        return False, reason
