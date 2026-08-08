from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from pptx_output_watermark.libreoffice_runner import resolve_soffice_path
from xlsx_image_compactor import compact_xlsx

FIXTURE = Path(__file__).parent / "fixtures" / "smart_target_sample.xlsx"


def copy_fixture(root: Path, name: str = "sample.xlsx") -> Path:
    target = root / name
    shutil.copy2(FIXTURE, target)
    return target


class XlsxImageCompactorTest(unittest.TestCase):
    def test_compresses_images_and_preserves_workbook_parts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = copy_fixture(root)
            output = root / "compressed.xlsx"

            result = compact_xlsx(
                source,
                0.25,
                output=output,
                validate_render=False,
                logger=lambda _: None,
            )

            self.assertLess(output.stat().st_size, source.stat().st_size)
            with ZipFile(source) as before, ZipFile(output) as after:
                self.assertEqual(set(before.namelist()), set(after.namelist()))
                for name in before.namelist():
                    if not name.startswith("xl/media/"):
                        self.assertEqual(before.read(name), after.read(name), name)
                self.assertNotEqual(
                    before.read("xl/media/image.jpg"),
                    after.read("xl/media/image.jpg"),
                )
            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            image = report["images"][0]
            self.assertEqual(image["occurrences"][0]["slide_number"], 2)
            self.assertIn("Summary", image["occurrences"][0]["owner_path"])
            self.assertEqual(
                report["presentation"]["render_validation"]["status"],
                "not_requested",
            )

    @unittest.skipUnless(
        Path(resolve_soffice_path()).is_file(), "LibreOffice is not installed"
    )
    def test_libreoffice_pdf_layout_validation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = copy_fixture(root)
            output = root / "compressed.xlsx"

            result = compact_xlsx(
                source, 0.25, output=output, validate_render=True, logger=lambda _: None
            )

            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            validation = report["presentation"]["render_validation"]
            self.assertEqual(validation["status"], "passed")
            self.assertEqual(validation["page_count"], 3)
            self.assertGreaterEqual(validation["minimum_page_ssim"], 0.985)

    def test_xlsm_macro_bytes_are_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = copy_fixture(root, "sample.xlsm")
            with ZipFile(source, "a", compression=ZIP_DEFLATED) as archive:
                archive.writestr("xl/vbaProject.bin", b"macro-bytes-must-not-change")
            output = root / "compressed.xlsm"

            compact_xlsx(
                source,
                0.25,
                output=output,
                validate_render=False,
                logger=lambda _: None,
            )

            with ZipFile(output) as archive:
                self.assertEqual(
                    archive.read("xl/vbaProject.bin"), b"macro-bytes-must-not-change"
                )

    def test_refuses_signed_invalid_and_legacy_workbooks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            signed = copy_fixture(root, "signed.xlsx")
            with ZipFile(signed, "a", compression=ZIP_DEFLATED) as archive:
                archive.writestr("_xmlsignatures/sig1.xml", b"<Signature/>")
            with self.assertRaisesRegex(RuntimeError, "Digitally signed"):
                compact_xlsx(signed, 0.25, validate_render=False, logger=lambda _: None)

            invalid = root / "invalid.xlsx"
            invalid.write_bytes(b"not-an-ooxml-package")
            with self.assertRaisesRegex(RuntimeError, "Encrypted or invalid"):
                compact_xlsx(
                    invalid, 0.01, validate_render=False, logger=lambda _: None
                )

            legacy = root / "legacy.xls"
            legacy.write_bytes(b"legacy")
            with self.assertRaisesRegex(ValueError, "XLSX/XLSM"):
                compact_xlsx(legacy, 0.01, validate_render=False, logger=lambda _: None)

    def test_source_under_target_is_skipped_without_duplicate_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = copy_fixture(Path(temp_dir))
            result = compact_xlsx(
                source,
                source.stat().st_size / 1_000_000 + 0.1,
                validate_render=False,
                logger=lambda _: None,
            )
            self.assertTrue(result["skipped"])
            self.assertEqual(result["output"], source.resolve())
            self.assertTrue(result["report_path"].is_file())

    def test_forced_mode_requires_safe_output_and_confirmation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = copy_fixture(root)
            with self.assertRaisesRegex(ValueError, "explicit confirmation"):
                compact_xlsx(
                    source,
                    0.25,
                    output=root / "forced.xlsx",
                    forced=True,
                    validate_render=False,
                    logger=lambda _: None,
                )


if __name__ == "__main__":
    unittest.main()
