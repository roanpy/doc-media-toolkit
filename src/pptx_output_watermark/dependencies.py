from __future__ import annotations

import os
import sys
from dataclasses import dataclass

from pptx_video_compactor import VIDEO_EXTENSIONS

from .ffmpeg_runtime import resolve_binary
from .keynote_runner import (
    KEYNOTE_APP_PATH,
    KEYNOTE_AUTOMATION_SETTINGS_URL,
    KEYNOTE_APP_STORE_URL,
    keynote_automation_status,
    keynote_available,
    keynote_conversion_disabled,
)
from .pages_runner import (
    PAGES_APP_PATH,
    PAGES_AUTOMATION_SETTINGS_URL,
    PAGES_APP_STORE_URL,
    pages_automation_status,
    pages_available,
    pages_conversion_disabled,
)
from .libreoffice_runner import conversion_disabled, resolve_soffice_path
from .presentation_rendering import check_document_com_engine_installed
from .models import ExportOptions

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"
FFMPEG_DOWNLOAD_URL = "https://ffmpeg.org/download.html"


@dataclass(slots=True)
class DependencyStatus:
    name: str
    required: bool
    available: bool
    path: str
    detail: str
    status_code: str = "ok"
    action_label: str = ""
    action_url: str = ""


def _needs_pdf_engine(options: ExportOptions) -> bool:
    input_suffix = options.input_path.suffix.lower()
    if input_suffix == ".pdf":
        return False
    if input_suffix in VIDEO_EXTENSIONS or input_suffix in {
        ".jpg",
        ".jpeg",
        ".jpe",
        ".png",
        ".webp",
    }:
        return False
    return not (options.output_format == "pptx" and options.output_mode == "editable")


def _needs_ffmpeg(options: ExportOptions) -> bool:
    if options.input_path.suffix.lower() in VIDEO_EXTENSIONS:
        return options.watermark.enabled
    return (
        options.output_format == "pptx"
        and options.output_mode == "image"
        and options.preserve_videos_in_image_pptx
    )


def dependency_statuses(options: ExportOptions) -> list[DependencyStatus]:
    statuses: list[DependencyStatus] = []
    needs_pdf_engine = _needs_pdf_engine(options)
    input_suffix = options.input_path.suffix.lower()

    has_com_app, com_app_name = check_document_com_engine_installed(input_suffix)
    if has_com_app:
        if "PowerPoint" in com_app_name:
            engine_name = "Microsoft PowerPoint"
        elif "Word" in com_app_name:
            engine_name = "Microsoft Word"
        else:
            engine_name = "WPS Office"
        statuses.append(
            DependencyStatus(
                name=engine_name,
                required=needs_pdf_engine,
                available=True,
                path="COM Automation",
                detail=f"Using native Windows COM automation ({com_app_name}).",
            )
        )
    else:
        soffice_path = resolve_soffice_path()
        soffice_exists = bool(soffice_path and os.path.exists(soffice_path))
        disabled, disabled_reason = conversion_disabled()

        if sys.platform == "darwin":
            if soffice_exists and not disabled:
                statuses.append(
                    DependencyStatus(
                        name="LibreOffice",
                        required=needs_pdf_engine,
                        available=True,
                        path=soffice_path or "Missing",
                        detail="Using LibreOffice as the primary document to PDF conversion engine.",
                        status_code="libreoffice_primary",
                    )
                )
            else:
                if input_suffix == ".pptx":
                    keynote_state, keynote_detail = keynote_automation_status()
                    keynote_disabled, keynote_disabled_reason = (
                        keynote_conversion_disabled()
                    )
                    if keynote_state == "ready" and not keynote_disabled:
                        if soffice_exists:
                            detail = (
                                "Using Keynote fallback because LibreOffice conversion is "
                                f"temporarily disabled: {disabled_reason}"
                            )
                        else:
                            detail = "Using Keynote fallback because LibreOffice is not installed."
                        statuses.append(
                            DependencyStatus(
                                name="Keynote",
                                required=needs_pdf_engine,
                                available=True,
                                path=str(KEYNOTE_APP_PATH),
                                detail=detail,
                                status_code="keynote_fallback",
                                action_label="Install LibreOffice",
                            )
                        )
                    else:
                        if keynote_disabled:
                            fallback_detail = (
                                "Keynote is installed but temporarily disabled: "
                                f"{keynote_disabled_reason}"
                            )
                            status_code = "keynote_disabled"
                            action_label = "Install LibreOffice"
                            action_url = ""
                        elif keynote_state == "permission_denied":
                            fallback_detail = keynote_detail
                            status_code = "keynote_permission_denied"
                            action_label = "Open Automation Settings"
                            action_url = KEYNOTE_AUTOMATION_SETTINGS_URL
                        elif keynote_state == "missing":
                            fallback_detail = "Keynote is not installed."
                            status_code = "keynote_missing"
                            action_label = "Get Keynote"
                            action_url = KEYNOTE_APP_STORE_URL
                        else:
                            fallback_detail = (
                                keynote_detail or "Keynote is unavailable."
                            )
                            status_code = "keynote_unavailable"
                            action_label = (
                                "Get Keynote"
                                if not keynote_available()
                                else "Install LibreOffice"
                            )
                            action_url = (
                                KEYNOTE_APP_STORE_URL if not keynote_available() else ""
                            )
                        statuses.append(
                            DependencyStatus(
                                name="LibreOffice / Keynote",
                                required=needs_pdf_engine,
                                available=False,
                                path=soffice_path
                                or str(
                                    KEYNOTE_APP_PATH
                                    if keynote_available()
                                    else "Missing"
                                ),
                                detail=(
                                    "LibreOffice is unavailable for PPTX export on macOS. "
                                    f"{fallback_detail}"
                                ),
                                status_code=status_code,
                                action_label=action_label,
                                action_url=action_url,
                            )
                        )
                else:
                    pages_state, pages_detail = pages_automation_status()
                    pages_disabled, pages_disabled_reason = pages_conversion_disabled()
                    if pages_state == "ready" and not pages_disabled:
                        if soffice_exists:
                            detail = (
                                "Using Pages fallback because LibreOffice conversion is "
                                f"temporarily disabled: {disabled_reason}"
                            )
                        else:
                            detail = "Using Pages fallback because LibreOffice is not installed."
                        statuses.append(
                            DependencyStatus(
                                name="Pages",
                                required=needs_pdf_engine,
                                available=True,
                                path=str(PAGES_APP_PATH),
                                detail=detail,
                                status_code="pages_fallback",
                                action_label="Install LibreOffice",
                            )
                        )
                    else:
                        if pages_disabled:
                            fallback_detail = (
                                "Pages is installed but temporarily disabled: "
                                f"{pages_disabled_reason}"
                            )
                            status_code = "pages_disabled"
                            action_label = "Install LibreOffice"
                            action_url = ""
                        elif pages_state == "permission_denied":
                            fallback_detail = pages_detail
                            status_code = "pages_permission_denied"
                            action_label = "Open Automation Settings"
                            action_url = PAGES_AUTOMATION_SETTINGS_URL
                        elif pages_state == "missing":
                            fallback_detail = "Pages is not installed."
                            status_code = "pages_missing"
                            action_label = "Get Pages"
                            action_url = PAGES_APP_STORE_URL
                        else:
                            fallback_detail = pages_detail or "Pages is unavailable."
                            status_code = "pages_unavailable"
                            action_label = (
                                "Get Pages"
                                if not pages_available()
                                else "Install LibreOffice"
                            )
                            action_url = (
                                PAGES_APP_STORE_URL if not pages_available() else ""
                            )
                        statuses.append(
                            DependencyStatus(
                                name="LibreOffice / Pages",
                                required=needs_pdf_engine,
                                available=False,
                                path=soffice_path
                                or str(
                                    PAGES_APP_PATH if pages_available() else "Missing"
                                ),
                                detail=(
                                    "LibreOffice is unavailable for DOCX export on macOS. "
                                    f"{fallback_detail}"
                                ),
                                status_code=status_code,
                                action_label=action_label,
                                action_url=action_url,
                            )
                        )
        elif soffice_exists:
            libreoffice_detail = "Using LibreOffice for PPTX to PDF conversion."
            if disabled:
                libreoffice_detail = (
                    f"LibreOffice conversion is temporarily disabled: {disabled_reason}"
                )
            statuses.append(
                DependencyStatus(
                    name="LibreOffice",
                    required=needs_pdf_engine,
                    available=not disabled,
                    path=soffice_path or "Missing",
                    detail=libreoffice_detail,
                    status_code="libreoffice_primary"
                    if not disabled
                    else "libreoffice_disabled",
                )
            )
        else:
            if sys.platform == "win32":
                status_name = "PowerPoint/WPS/LibreOffice Engine"
                libreoffice_detail = (
                    "PowerPoint/WPS COM and LibreOffice are both missing. "
                    "Install Microsoft Office, WPS, or LibreOffice."
                )
            else:
                status_name = "LibreOffice"
                libreoffice_detail = "LibreOffice soffice was not found."

            statuses.append(
                DependencyStatus(
                    name=status_name,
                    required=needs_pdf_engine,
                    available=False,
                    path="Missing",
                    detail=libreoffice_detail,
                    status_code="engine_missing",
                    action_label="Install LibreOffice",
                    action_url=LIBREOFFICE_DOWNLOAD_URL,
                )
            )

    ffmpeg_path = resolve_binary("ffmpeg")
    ffmpeg_exists = bool(ffmpeg_path and os.path.exists(ffmpeg_path))
    statuses.append(
        DependencyStatus(
            name="FFmpeg",
            required=_needs_ffmpeg(options),
            available=ffmpeg_exists,
            path=ffmpeg_path or "Missing",
            status_code="ffmpeg_available" if ffmpeg_exists else "ffmpeg_missing",
            detail=(
                "Required to watermark and reinsert videos for image-based PPTX export."
                if not ffmpeg_exists
                else "Used for embedded video watermarking and poster extraction."
            ),
            action_label="Install FFmpeg" if not ffmpeg_exists else "",
            action_url=FFMPEG_DOWNLOAD_URL if not ffmpeg_exists else "",
        )
    )

    return statuses


def missing_dependency_message(options: ExportOptions) -> str | None:
    missing = [
        status
        for status in dependency_statuses(options)
        if status.required and not status.available
    ]
    if not missing:
        return None
    lines = [
        "Missing required runtime dependencies for the selected export mode:",
    ]
    for status in missing:
        suffix = f" ({status.detail})" if status.detail else ""
        lines.append(f"- {status.name}{suffix}")
        if status.action_label:
            action = f"  Next step: {status.action_label}"
            if status.action_url:
                action += f" — {status.action_url}"
            lines.append(action)
    return "\n".join(lines)
