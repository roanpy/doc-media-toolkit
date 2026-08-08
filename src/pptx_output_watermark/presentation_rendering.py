"""Presentation-to-PDF rendering helpers."""

from __future__ import annotations

import os
import signal
import shutil
import subprocess
import sys
import time
from collections.abc import Callable
from contextlib import contextmanager
from pathlib import Path

from . import libreoffice_runner
from . import keynote_runner
from . import pages_runner
from .runtime_temp import create_runtime_temp_dir

Logger = Callable[[str], None]

COM_PDF_ENGINES: tuple[tuple[str, str], ...] = (
    ("Microsoft PowerPoint", "PowerPoint.Application"),
    ("WPS Office", "KWPP.Application"),
)

COM_WORD_PDF_ENGINES: tuple[tuple[str, str], ...] = (
    ("Microsoft Word", "Word.Application"),
    ("WPS Office", "KWPS.Application"),
)

COM_EXCEL_PDF_ENGINES: tuple[tuple[str, str], ...] = (
    ("Microsoft Excel", "Excel.Application"),
    ("WPS Office", "KET.Application"),
)


def _log(logger: Logger | None, message: str) -> None:
    if logger is not None:
        logger(message)


def _compact_detail(detail: str, *, limit: int = 220) -> str:
    normalized = " ".join(str(detail or "").split())
    if len(normalized) <= limit:
        return normalized
    return f"{normalized[: limit - 1]}…"


def _is_valid_pdf_with_eof(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            if f.read(4) != b"%PDF":
                return False
            f.seek(0, 2)
            size = f.tell()
            if size > 10:
                f.seek(max(0, size - 1024))
                tail = f.read()
                return b"%%EOF" in tail
    except Exception:
        pass
    return False


def _is_readable_pdf(path: Path) -> bool:
    if _is_valid_pdf_with_eof(path):
        return True
    try:
        from pypdf import PdfReader

        with open(path, "rb") as handle:
            reader = PdfReader(handle, strict=False)
            return len(reader.pages) > 0
    except Exception:
        return False


def _wait_for_readable_pdf(
    path: Path,
    *,
    timeout_seconds: float,
    poll_interval: float = 0.5,
) -> bool:
    deadline = time.time() + max(0.1, timeout_seconds)
    last_size = -1
    stable_ticks = 0
    while time.time() < deadline:
        try:
            if path.exists() and path.stat().st_size > 0:
                size = path.stat().st_size
                if size == last_size:
                    stable_ticks += 1
                else:
                    stable_ticks = 0
                    last_size = size
                if stable_ticks >= 2 and _is_readable_pdf(path):
                    return True
        except Exception:
            pass
        time.sleep(poll_interval)
    return path.exists() and _is_readable_pdf(path)


def _safe_which(command: str) -> str | None:
    try:
        return shutil.which(command)
    except Exception:
        return None


def _wps_executable_candidates() -> list[Path]:
    candidates: list[Path] = []
    for env_name in (
        "PPTX_TOOLS_WPS",
        "PPTX_TOOLS_WPP",
        "PPTX_OUTPUT_WATERMARK_WPS",
        "PPTX_OUTPUT_WATERMARK_WPP",
    ):
        override = os.environ.get(env_name, "").strip()
        if override:
            candidates.append(Path(override).expanduser())

    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "").strip()
        localappdata = os.environ.get("LOCALAPPDATA", "").strip()
        for root in (appdata, localappdata):
            if not root:
                continue
            root_path = Path(root)
            candidates.extend(
                [
                    root_path / "kingsoft" / "office6" / "wpp.exe",
                    root_path / "Kingsoft" / "office6" / "wpp.exe",
                    root_path / "Kingsoft" / "WPS Office" / "office6" / "wpp.exe",
                    root_path / "Kingsoft" / "WPS Office" / "ksolaunch.exe",
                ]
            )
            wps_root = root_path / "Kingsoft" / "WPS Office"
            if wps_root.exists():
                candidates.extend(wps_root.glob("**/office6/wpp.exe"))

        for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
            root = os.environ.get(env_name, "").strip()
            if not root:
                continue
            candidates.extend(
                [
                    Path(root) / "Kingsoft" / "WPS Office" / "office6" / "wpp.exe",
                    Path(root) / "WPS Office" / "office6" / "wpp.exe",
                ]
            )

    for command in ("wpp.exe", "wps.exe", "et.exe"):
        found = _safe_which(command)
        if found:
            candidates.append(Path(found))

    seen: set[Path] = set()
    result: list[Path] = []
    for candidate in candidates:
        try:
            resolved = candidate.expanduser()
        except Exception:
            continue
        if resolved in seen:
            continue
        seen.add(resolved)
        if resolved.exists():
            result.append(resolved)
    return result


def _wps_runtime_dirs() -> list[str]:
    dirs: list[str] = []
    for candidate in _wps_executable_candidates():
        directory = str(candidate.parent)
        if directory not in dirs:
            dirs.append(directory)
    return dirs


@contextmanager
def _temporary_path_entries(entries: list[str]):
    if not entries:
        yield
        return
    previous_path = os.environ.get("PATH", "")
    added_dll_dirs = []
    try:
        os.environ["PATH"] = os.pathsep.join([*entries, previous_path])
        if sys.platform == "win32" and hasattr(os, "add_dll_directory"):
            for entry in entries:
                try:
                    added_dll_dirs.append(os.add_dll_directory(entry))
                except Exception:
                    pass
        yield
    finally:
        os.environ["PATH"] = previous_path
        for handle in added_dll_dirs:
            try:
                handle.close()
            except Exception:
                pass


def _export_word_to_pdf(document, output_pdf: Path) -> bool:
    output_pdf = output_pdf.resolve()
    export_attempts = (
        lambda: document.ExportAsFixedFormat(
            str(output_pdf), 17
        ),  # wdExportFormatPDF = 17
        lambda: document.SaveAs(str(output_pdf), 17),
    )
    for export in export_attempts:
        try:
            output_pdf.unlink(missing_ok=True)
            export()
            if _wait_for_readable_pdf(output_pdf, timeout_seconds=30.0):
                return True
        except Exception:
            continue
    return False


def _open_word_com_document(app, input_docx: Path) -> tuple[object | None, str]:
    input_path = str(input_docx.resolve())
    open_attempts = (
        lambda: app.Documents.Open(
            input_path,
            ReadOnly=True,
            Visible=False,
        ),
        lambda: app.Documents.Open(
            input_path,
            ReadOnly=True,
            Visible=True,
        ),
        lambda: app.Documents.Open(input_path),
    )
    details: list[str] = []
    for index, open_attempt in enumerate(open_attempts, start=1):
        try:
            doc = open_attempt()
            if doc is not None:
                return doc, ""
            details.append(f"attempt {index}: returned None")
        except Exception as exc:
            details.append(f"attempt {index}: {type(exc).__name__}: {exc}")
    return None, "; ".join(details)


def _export_presentation_to_pdf(presentation, output_pdf: Path) -> bool:
    output_pdf = output_pdf.resolve()
    export_attempts = (
        lambda: presentation.ExportAsFixedFormat(str(output_pdf), 2, 1),
        lambda: presentation.ExportAsFixedFormat(str(output_pdf), 2),
        # ppSaveAsPDF = 32. Some WPS builds expose SaveAs more reliably than
        # the full PowerPoint ExportAsFixedFormat signature.
        lambda: presentation.SaveAs(str(output_pdf), 32),
    )
    for export in export_attempts:
        try:
            output_pdf.unlink(missing_ok=True)
            export()
            if _wait_for_readable_pdf(output_pdf, timeout_seconds=30.0):
                return True
        except Exception:
            continue
    return False


def _export_excel_to_pdf(workbook, output_pdf: Path) -> bool:
    output_pdf = output_pdf.resolve()
    export_attempts = (
        lambda: workbook.ExportAsFixedFormat(0, str(output_pdf)),
        lambda: workbook.ExportAsFixedFormat(Type=0, Filename=str(output_pdf)),
    )
    for export in export_attempts:
        try:
            output_pdf.unlink(missing_ok=True)
            export()
            if _wait_for_readable_pdf(output_pdf, timeout_seconds=30.0):
                return True
        except Exception:
            continue
    return False


def _open_excel_com_workbook(app, input_xlsx: Path) -> tuple[object | None, str]:
    input_path = str(input_xlsx.resolve())
    open_attempts = (
        lambda: app.Workbooks.Open(input_path, ReadOnly=True, UpdateLinks=0),
        lambda: app.Workbooks.Open(input_path, ReadOnly=True),
        lambda: app.Workbooks.Open(input_path),
    )
    details: list[str] = []
    for index, open_attempt in enumerate(open_attempts, start=1):
        try:
            workbook = open_attempt()
            if workbook is not None:
                return workbook, ""
            details.append(f"attempt {index}: returned None")
        except Exception as exc:
            details.append(f"attempt {index}: {type(exc).__name__}: {exc}")
    return None, "; ".join(details)


def _open_com_presentation(app, input_pptx: Path) -> tuple[object | None, str]:
    input_path = str(input_pptx.resolve())
    open_attempts = (
        lambda: app.Presentations.Open(
            input_path,
            ReadOnly=1,
            Untitled=0,
            WithWindow=0,
        ),
        # Some WPS builds fail hidden automation but can still export reliably
        # once the presentation is opened with a real window.
        lambda: app.Presentations.Open(
            input_path,
            ReadOnly=1,
            Untitled=0,
            WithWindow=1,
        ),
        lambda: app.Presentations.Open(input_path),
    )
    details: list[str] = []
    for index, open_attempt in enumerate(open_attempts, start=1):
        try:
            presentation = open_attempt()
            if presentation is not None:
                return presentation, ""
            details.append(f"attempt {index}: returned None")
        except Exception as exc:
            details.append(f"attempt {index}: {type(exc).__name__}: {exc}")
    return None, "; ".join(details)


def _convert_via_com_app(app_id: str, input_pptx: Path, output_pdf: Path) -> bool:
    success, _detail = _convert_via_com_app_with_detail(app_id, input_pptx, output_pdf)
    return success


def _convert_via_com_app_with_detail(
    app_id: str,
    input_pptx: Path,
    output_pdf: Path,
    is_word: bool = False,
    is_excel: bool = False,
) -> tuple[bool, str]:
    import comtypes.client

    app = None
    presentation = None
    success = False
    safe_path: Path | None = None
    post_close_validation = False
    failure_detail = "COM export did not produce a readable PDF"
    runtime_dirs = (
        _wps_runtime_dirs()
        if app_id.startswith("KW") or app_id == "KET.Application"
        else []
    )
    try:
        with _temporary_path_entries(runtime_dirs):
            app = comtypes.client.CreateObject(app_id)
            if app is None:
                return False, "COM CreateObject returned None"
            if is_excel:
                presentation, open_detail = _open_excel_com_workbook(app, input_pptx)
            elif is_word:
                presentation, open_detail = _open_word_com_document(app, input_pptx)
            else:
                presentation, open_detail = _open_com_presentation(app, input_pptx)
            if presentation is None:
                document_type = (
                    "Workbook"
                    if is_excel
                    else "Document"
                    if is_word
                    else "Presentation"
                )
                return False, f"{document_type} open failed: {open_detail}"
            if is_excel:
                exported = _export_excel_to_pdf(presentation, output_pdf)
            elif is_word:
                exported = _export_word_to_pdf(presentation, output_pdf)
            else:
                exported = _export_presentation_to_pdf(presentation, output_pdf)
            if exported:
                safe_path = output_pdf.with_name(output_pdf.name + ".safe")
                shutil.copy2(output_pdf, safe_path)
                success = True
            else:
                post_close_validation = (
                    output_pdf.exists() and output_pdf.stat().st_size > 0
                )
                failure_detail = "COM export did not produce a readable PDF"
    except Exception as exc:
        hint = ""
        if app_id == "KWPP.Application" and runtime_dirs:
            hint = f"; WPS runtime dirs: {', '.join(runtime_dirs)}"
        return False, f"{type(exc).__name__}: {exc}{hint}"
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    if (
        not success
        and post_close_validation
        and _wait_for_readable_pdf(output_pdf, timeout_seconds=8.0, poll_interval=0.4)
    ):
        try:
            safe_path = output_pdf.with_name(output_pdf.name + ".safe")
            shutil.copy2(output_pdf, safe_path)
            success = True
        except Exception as exc:
            return (
                False,
                f"COM export finished after shutdown but safe PDF restore failed: {type(exc).__name__}: {exc}",
            )

    if success:
        if safe_path is not None and safe_path.exists():
            try:
                shutil.copy2(safe_path, output_pdf)
                safe_path.unlink(missing_ok=True)
                return True, ""
            except Exception as exc:
                return (
                    False,
                    f"COM export succeeded but safe PDF restore failed: {type(exc).__name__}: {exc}",
                )
        if output_pdf.exists():
            return True, ""
        return (
            False,
            "COM export succeeded but output PDF was removed before it could be used",
        )

    if safe_path is not None:
        safe_path.unlink(missing_ok=True)
    return False, failure_detail


def _registered_com_progid(progid: str) -> bool:
    try:
        import winreg
    except Exception:
        return False

    candidate_keys = (
        (winreg.HKEY_CLASSES_ROOT, progid),
        (winreg.HKEY_CLASSES_ROOT, f"{progid}\\CurVer"),
        (winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{progid}"),
        (winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{progid}\\CurVer"),
        (winreg.HKEY_LOCAL_MACHINE, f"Software\\Classes\\{progid}"),
        (winreg.HKEY_LOCAL_MACHINE, f"Software\\Classes\\{progid}\\CurVer"),
        (winreg.HKEY_LOCAL_MACHINE, f"Software\\WOW6432Node\\Classes\\{progid}"),
        (
            winreg.HKEY_LOCAL_MACHINE,
            f"Software\\WOW6432Node\\Classes\\{progid}\\CurVer",
        ),
    )
    for root, subkey in candidate_keys:
        try:
            with winreg.OpenKey(root, subkey) as _:
                return True
        except Exception:
            continue
    return False


def check_document_com_engine_installed(
    input_suffix: str = ".pptx",
) -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, ""
    suffix = input_suffix.lower()
    engines = (
        COM_WORD_PDF_ENGINES
        if suffix in {".docx", ".docm"}
        else COM_EXCEL_PDF_ENGINES
        if suffix in {".xlsx", ".xlsm"}
        else COM_PDF_ENGINES
    )
    for _, progid in engines:
        if _registered_com_progid(progid):
            return True, progid
    return False, ""


def check_office_com_engine_installed() -> tuple[bool, str]:
    return check_document_com_engine_installed(".pptx")


def check_powerpoint_installed() -> tuple[bool, str]:
    return check_office_com_engine_installed()


def _convert_via_powerpoint_com(input_pptx: Path, output_pdf: Path) -> bool:
    return _convert_via_office_com_with_errors(input_pptx, output_pdf, None)


def _convert_via_office_com_with_errors(
    input_path: Path,
    output_pdf: Path,
    errors: list[str] | None,
    logger: Logger | None = None,
) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import comtypes.client  # noqa: F401
    except ImportError as exc:
        if errors is not None:
            errors.append(f"COM automation unavailable: {exc}")
        _log(
            logger,
            "Windows COM automation is unavailable; trying LibreOffice fallback.",
        )
        return False

    suffix = input_path.suffix.lower()
    is_word = suffix in {".docx", ".docm"}
    is_excel = suffix in {".xlsx", ".xlsm"}
    engines = (
        COM_WORD_PDF_ENGINES
        if is_word
        else COM_EXCEL_PDF_ENGINES
        if is_excel
        else COM_PDF_ENGINES
    )

    for engine_name, app_id in engines:
        _log(logger, f"Trying PDF export engine: {engine_name} COM")
        success, detail = _convert_via_com_app_with_detail(
            app_id,
            input_path,
            output_pdf,
            is_word=is_word,
            is_excel=is_excel,
        )
        if success:
            _log(logger, f"PDF export engine: {engine_name} COM")
            return True
        if errors is not None:
            errors.append(f"{engine_name} ({app_id}): {detail}")
        if detail:
            _log(
                logger,
                f"PDF export engine failed: {engine_name} COM ({_compact_detail(detail)})",
            )
    return False


def _format_returncode(returncode: int | None) -> str:
    if returncode is None:
        return "returncode=unknown"
    if returncode < 0:
        signal_number = abs(returncode)
        try:
            signal_name = signal.Signals(signal_number).name
        except ValueError:
            signal_name = "UNKNOWN_SIGNAL"
        return f"signal={signal_number} ({signal_name})"
    if sys.platform == "win32" and returncode:
        unsigned_code = returncode & 0xFFFFFFFF
        if unsigned_code == 0xC0000142:
            return (
                "returncode=3221225794 (0xC0000142, LibreOffice initialization failed; "
                "prefer soffice.com over soffice.exe or reinstall LibreOffice)"
            )
        return f"returncode={returncode} (0x{unsigned_code:08X})"
    return f"returncode={returncode}"


def _format_convert_error(proc: subprocess.CompletedProcess) -> str:
    code_detail = _format_returncode(proc.returncode)
    message = (proc.stderr or proc.stdout or "").strip()
    if message:
        return f"{code_detail}: {message}"
    return code_detail


def convert_document_to_pdf(
    input_path: Path,
    *,
    timeout_seconds: int = 120,
    logger: Logger | None = None,
) -> Path:
    temp_dir = create_runtime_temp_dir(
        "pptx_output_watermark_pdf_",
        purpose="document_to_pdf_output",
    )
    pdf_name = input_path.with_suffix(".pdf").name
    output_path = temp_dir / pdf_name

    com_errors: list[str] = []
    keynote_detail: str = ""
    pages_detail: str = ""
    if sys.platform == "win32":
        if _convert_via_office_com_with_errors(
            input_path,
            output_path,
            com_errors,
            logger=logger,
        ):
            if output_path.exists():
                return output_path
            else:
                com_errors.append(
                    f"COM engine reported success, but output file {output_path.name} was missing afterwards."
                )
    soffice_path = libreoffice_runner.resolve_soffice_path()
    soffice_exists = os.path.exists(soffice_path)
    disabled, disabled_reason = libreoffice_runner.conversion_disabled()

    libreoffice_failed = False
    last_error = ""

    if soffice_exists and not disabled:
        try:
            input_suffix = input_path.suffix.lower()
            if input_suffix == ".docx":
                filters = ("pdf:writer_pdf_Export", "pdf")
            elif input_suffix == ".pptx":
                filters = ("pdf:impress_pdf_Export", "pdf")
            elif input_suffix in {".xlsx", ".xlsm"}:
                filters = ("pdf:calc_pdf_Export", "pdf")
            else:
                filters = ("pdf",)
            for export_filter in filters:
                profile_dir = str(
                    create_runtime_temp_dir(
                        "pptx_output_watermark_lo_profile_",
                        purpose="libreoffice_pdf_export_profile",
                    )
                )
                try:
                    if sys.platform == "darwin":
                        _log(logger, "Trying PDF export engine: LibreOffice (macOS)")
                    elif sys.platform == "win32":
                        _log(logger, "Trying PDF export engine: LibreOffice fallback")
                    else:
                        _log(logger, "Trying PDF export engine: LibreOffice")
                    args = [
                        "--headless",
                        "--invisible",
                        "--nologo",
                        "--nodefault",
                        "--nofirststartwizard",
                        "--nolockcheck",
                        "--norestore",
                        "--convert-to",
                        export_filter,
                        "--outdir",
                        str(temp_dir),
                        str(input_path),
                    ]
                    proc = libreoffice_runner.run_convert_command(
                        soffice_path,
                        args,
                        profile_dir=profile_dir,
                        timeout_seconds=max(10, int(timeout_seconds)),
                    )
                finally:
                    shutil.rmtree(profile_dir, ignore_errors=True)

                output_path = temp_dir / pdf_name
                if proc.returncode == 0 and _wait_for_readable_pdf(
                    output_path, timeout_seconds=10.0
                ):
                    if sys.platform == "darwin":
                        _log(logger, "PDF export engine: LibreOffice (macOS)")
                    elif sys.platform == "win32":
                        _log(logger, "PDF export engine: LibreOffice fallback")
                    else:
                        _log(logger, "PDF export engine: LibreOffice")
                    return output_path
                if proc.returncode == 0:
                    last_error = (
                        "LibreOffice returned 0 but did not produce a readable PDF."
                    )
                else:
                    last_error = _format_convert_error(proc)
                _log(
                    logger,
                    f"LibreOffice PDF export failed: {_compact_detail(last_error)}",
                )
                if proc.returncode < 0 or sys.platform == "darwin":
                    break
            if sys.platform == "darwin":
                libreoffice_runner.disable_conversion_temporarily(
                    last_error or "macOS LibreOffice conversion produced no PDF"
                )
            libreoffice_failed = True
        except subprocess.TimeoutExpired:
            libreoffice_runner.disable_conversion_temporarily(
                f"timeout>{int(timeout_seconds)}s"
            )
            last_error = (
                f"LibreOffice PDF export timed out after {int(timeout_seconds)}s"
            )
            libreoffice_failed = True
    else:
        if not soffice_exists:
            last_error = "LibreOffice soffice was not found."
        else:
            last_error = (
                f"LibreOffice conversion is temporarily disabled: {disabled_reason}"
            )
        libreoffice_failed = True

    if libreoffice_failed and sys.platform == "darwin":
        input_suffix = input_path.suffix.lower()
        keynote_supported = input_suffix == ".pptx"
        pages_supported = input_suffix == ".docx"
        if keynote_supported:
            keynote_disabled, keynote_disabled_reason = (
                keynote_runner.keynote_conversion_disabled()
            )
            keynote_state, automation_detail = (
                keynote_runner.keynote_automation_status()
            )
            if keynote_disabled:
                keynote_detail = (
                    "Keynote conversion is temporarily disabled: "
                    f"{keynote_disabled_reason}"
                )
                _log(logger, f"Keynote fallback skipped: {keynote_detail}")
            elif keynote_state == "permission_denied":
                keynote_detail = automation_detail
                _log(
                    logger,
                    f"Keynote fallback skipped: {_compact_detail(keynote_detail)}",
                )
            elif keynote_state == "ready":
                _log(logger, "Trying PDF export engine: Keynote (macOS)")
                success, keynote_detail = keynote_runner.convert_via_keynote(
                    input_path, output_path, timeout_seconds=timeout_seconds
                )
                if success and _wait_for_readable_pdf(
                    output_path, timeout_seconds=10.0
                ):
                    _log(logger, "PDF export engine: Keynote (macOS)")
                    return output_path
                if success:
                    keynote_detail = "Keynote produced a PDF that was not readable."
                    keynote_runner.disable_keynote_temporarily(keynote_detail)
                _log(
                    logger,
                    f"Keynote PDF export failed: {_compact_detail(keynote_detail)}",
                )
            else:
                keynote_detail = automation_detail or "Keynote is not installed."
                _log(
                    logger,
                    f"Keynote fallback skipped: {_compact_detail(keynote_detail)}",
                )
        elif pages_supported:
            pages_disabled, pages_disabled_reason = (
                pages_runner.pages_conversion_disabled()
            )
            pages_state, automation_detail = pages_runner.pages_automation_status()
            if pages_disabled:
                pages_detail = (
                    f"Pages conversion is temporarily disabled: {pages_disabled_reason}"
                )
                _log(logger, f"Pages fallback skipped: {pages_detail}")
            elif pages_state == "permission_denied":
                pages_detail = automation_detail
                _log(logger, f"Pages fallback skipped: {_compact_detail(pages_detail)}")
            elif pages_state == "ready":
                _log(logger, "Trying PDF export engine: Pages (macOS)")
                success, pages_detail = pages_runner.convert_via_pages(
                    input_path, output_path, timeout_seconds=timeout_seconds
                )
                if success and _wait_for_readable_pdf(
                    output_path, timeout_seconds=10.0
                ):
                    _log(logger, "PDF export engine: Pages (macOS)")
                    return output_path
                if success:
                    pages_detail = "Pages produced a PDF that was not readable."
                    pages_runner.disable_pages_temporarily(pages_detail)
                _log(
                    logger, f"Pages PDF export failed: {_compact_detail(pages_detail)}"
                )
            else:
                pages_detail = automation_detail or "Pages is not installed."
                _log(logger, f"Pages fallback skipped: {_compact_detail(pages_detail)}")
        elif not keynote_supported:
            keynote_detail = "Keynote only supports PPTX fallback in this tool."
            _log(logger, "Keynote fallback skipped: source is not PPTX.")
        if keynote_supported:
            raise RuntimeError(
                "No available macOS PDF export engine completed successfully. "
                "Install Keynote from the Mac App Store for the lightest fallback, "
                "or install LibreOffice if Keynote output fidelity is not acceptable.\n"
                f"LibreOffice error: {last_error}\n"
                f"Keynote error: {keynote_detail}"
            )
        if input_suffix in {".xlsx", ".xlsm"}:
            raise RuntimeError(
                "No available macOS PDF export engine completed successfully. "
                "XLSX/XLSM export on macOS requires LibreOffice Calc.\n"
                f"LibreOffice error: {last_error}"
            )
        raise RuntimeError(
            "No available macOS PDF export engine completed successfully. "
            "DOCX export on macOS requires LibreOffice or Pages.\n"
            f"LibreOffice error: {last_error}\n"
            f"Pages error: {pages_detail}"
        )

    if libreoffice_failed:
        if not soffice_exists:
            if sys.platform == "win32" and com_errors:
                detail = "\n".join(f"- {item}" for item in com_errors)
                wps_dirs = _wps_runtime_dirs()
                wps_hint = (
                    "\nDetected WPS runtime directories:\n"
                    + "\n".join(f"- {item}" for item in wps_dirs)
                    if wps_dirs
                    else ""
                )
                raise RuntimeError(
                    "Office/WPS PDF export failed, and LibreOffice soffice was not found. "
                    "If WPS is installed, open WPS once or repair WPS so KWPP.Application COM is registered, "
                    "or set PPTX_TOOLS_WPP to the full wpp.exe path."
                    f"\nCOM export attempts:\n{detail}{wps_hint}"
                )
            raise FileNotFoundError(
                "LibreOffice soffice was not found. Install LibreOffice or add soffice to PATH."
            )
        raise RuntimeError(f"LibreOffice failed to export PDF: {last_error}")
