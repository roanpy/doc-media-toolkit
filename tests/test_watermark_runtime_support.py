from __future__ import annotations

import json
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx_output_watermark.dependencies import (
    dependency_statuses,
    missing_dependency_message,
)
from pptx_output_watermark.libreoffice_runner import (
    resolve_soffice_path,
    run_convert_command,
)
from pptx_output_watermark.models import ExportOptions
from pptx_output_watermark.presentation_rendering import (
    _export_excel_to_pdf,
    _export_word_to_pdf,
    check_document_com_engine_installed,
)
from pptx_output_watermark.runtime_temp import (
    OWNER_FILE_NAME,
    cleanup_stale_runtime_entries,
    register_runtime_dir,
)


class RuntimeTempCleanupTest(unittest.TestCase):
    def test_cleanup_removes_stale_dead_runtime_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            stale_dir = temp_root / "pptx_output_watermark_preview_source_stale"
            register_runtime_dir(stale_dir, purpose="test")
            owner_file = stale_dir / OWNER_FILE_NAME
            payload = json.loads(owner_file.read_text(encoding="utf-8"))
            payload["pid"] = 999999
            owner_file.write_text(json.dumps(payload), encoding="utf-8")
            old_time = time.time() - 48 * 3600
            for path in (stale_dir, owner_file):
                os.utime(path, (old_time, old_time))

            with patch(
                "pptx_output_watermark.runtime_temp.tempfile.gettempdir",
                return_value=str(temp_root),
            ):
                cleanup_stale_runtime_entries(
                    patterns=("pptx_output_watermark_preview_source_*",),
                    max_age_hours=24,
                )

            self.assertFalse(stale_dir.exists())


class MacDependencyStatusTest(unittest.TestCase):
    def test_keynote_permission_denied_is_reported(self) -> None:
        options = ExportOptions(input_path=Path("sample.pptx"), output_format="pdf")
        with (
            patch(
                "pptx_output_watermark.dependencies.resolve_soffice_path",
                return_value="/missing/soffice",
            ),
            patch(
                "pptx_output_watermark.dependencies.resolve_binary",
                return_value="/missing/ffmpeg",
            ),
            patch(
                "pptx_output_watermark.dependencies.os.path.exists", return_value=False
            ),
            patch(
                "pptx_output_watermark.dependencies.conversion_disabled",
                return_value=(False, ""),
            ),
            patch(
                "pptx_output_watermark.dependencies.keynote_automation_status",
                return_value=("permission_denied", "automation denied"),
            ),
            patch(
                "pptx_output_watermark.dependencies.keynote_conversion_disabled",
                return_value=(False, ""),
            ),
            patch(
                "pptx_output_watermark.dependencies.keynote_available",
                return_value=True,
            ),
            patch(
                "pptx_output_watermark.dependencies.sys.platform",
                "darwin",
            ),
        ):
            statuses = dependency_statuses(options)

        self.assertEqual(statuses[0].status_code, "keynote_permission_denied")
        self.assertFalse(statuses[0].available)
        self.assertIn("automation denied", statuses[0].detail)

    def test_missing_engine_message_contains_actionable_download(self) -> None:
        options = ExportOptions(input_path=Path("sample.pptx"), output_format="pdf")
        with (
            patch(
                "pptx_output_watermark.dependencies.check_document_com_engine_installed",
                return_value=(False, ""),
            ),
            patch(
                "pptx_output_watermark.dependencies.resolve_soffice_path",
                return_value="/missing/soffice",
            ),
            patch(
                "pptx_output_watermark.dependencies.resolve_binary",
                return_value="/missing/ffmpeg",
            ),
            patch(
                "pptx_output_watermark.dependencies.os.path.exists", return_value=False
            ),
            patch("pptx_output_watermark.dependencies.sys.platform", "win32"),
        ):
            message = missing_dependency_message(options)

        self.assertIn("Install LibreOffice", message or "")
        self.assertIn("https://www.libreoffice.org/", message or "")

    def test_bundled_soffice_is_preferred_to_system_install(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            bundled = (
                root
                / "libreoffice"
                / "LibreOffice.app"
                / "Contents"
                / "MacOS"
                / "soffice"
            )
            bundled.parent.mkdir(parents=True)
            bundled.touch()
            with (
                patch(
                    "pptx_output_watermark.libreoffice_runner.sys.platform", "darwin"
                ),
                patch(
                    "pptx_output_watermark.libreoffice_runner.sys._MEIPASS",
                    str(root),
                    create=True,
                ),
                patch.dict(os.environ, {"PPTX_TOOLS_SOFFICE": ""}),
            ):
                self.assertEqual(resolve_soffice_path(), str(bundled))

    def test_macos_launches_the_containing_libreoffice_app(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            app = Path(temp_dir) / "LibreOffice.app"
            soffice = app / "Contents" / "MacOS" / "soffice"
            soffice.parent.mkdir(parents=True)
            soffice.touch()
            with (
                patch(
                    "pptx_output_watermark.libreoffice_runner.sys.platform", "darwin"
                ),
                patch(
                    "pptx_output_watermark.libreoffice_runner.run_process"
                ) as run_process,
            ):
                run_convert_command(
                    str(soffice),
                    ["--headless"],
                    profile_dir=temp_dir,
                    timeout_seconds=30,
                )

            command = run_process.call_args.args[0]
            self.assertEqual(command[4:6], ["-a", str(app.resolve())])


class WordPdfRuntimeTest(unittest.TestCase):
    def test_word_export_uses_fixed_format_pdf(self) -> None:
        calls: list[tuple[object, ...]] = []

        class Document:
            def ExportAsFixedFormat(self, *args) -> None:
                calls.append(args)

        with (
            tempfile.TemporaryDirectory() as temp_dir,
            patch(
                "pptx_output_watermark.presentation_rendering._wait_for_readable_pdf",
                return_value=True,
            ),
        ):
            self.assertTrue(
                _export_word_to_pdf(Document(), Path(temp_dir) / "output.pdf")
            )

        self.assertEqual(calls[0][1], 17)

    def test_docx_checks_word_then_wps_com_on_windows(self) -> None:
        with (
            patch("pptx_output_watermark.presentation_rendering.sys.platform", "win32"),
            patch(
                "pptx_output_watermark.presentation_rendering._registered_com_progid",
                side_effect=lambda progid: progid == "KWPS.Application",
            ),
        ):
            self.assertEqual(
                check_document_com_engine_installed(".docx"),
                (True, "KWPS.Application"),
            )


class ExcelPdfRuntimeTest(unittest.TestCase):
    def test_excel_export_uses_fixed_format_pdf(self) -> None:
        calls: list[tuple[object, ...]] = []

        class Workbook:
            def ExportAsFixedFormat(self, *args, **kwargs) -> None:
                calls.append(args or tuple(kwargs.values()))

        with (
            tempfile.TemporaryDirectory() as temp_dir,
            patch(
                "pptx_output_watermark.presentation_rendering._wait_for_readable_pdf",
                return_value=True,
            ),
        ):
            self.assertTrue(
                _export_excel_to_pdf(Workbook(), Path(temp_dir) / "output.pdf")
            )

        self.assertEqual(calls[0][0], 0)

    def test_xlsx_checks_excel_com_engines_on_windows(self) -> None:
        with (
            patch("pptx_output_watermark.presentation_rendering.sys.platform", "win32"),
            patch(
                "pptx_output_watermark.presentation_rendering._registered_com_progid",
                side_effect=lambda progid: progid == "Excel.Application",
            ),
        ):
            self.assertEqual(
                check_document_com_engine_installed(".xlsx"),
                (True, "Excel.Application"),
            )


if __name__ == "__main__":
    unittest.main()
