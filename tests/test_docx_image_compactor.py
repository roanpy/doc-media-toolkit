from __future__ import annotations

import io
import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_STORED, ZipFile

from PIL import Image

from docx_image_compactor import compact_docx
from pptx_output_watermark.libreoffice_runner import resolve_soffice_path


RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def make_docx(path: Path, *, signed: bool = False, macro: bool = False) -> None:
    image_bytes = io.BytesIO()
    Image.new("RGB", (900, 900), (40, 100, 180)).save(
        image_bytes, format="PNG", compress_level=0
    )
    signature = (
        f'<Relationship Id="sig" Type="{RELATIONSHIPS}/digital-signature/origin" '
        'Target="_xmlsignatures/origin.sigs"/>'
        if signed
        else ""
    )
    with ZipFile(path, "w", ZIP_STORED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            f'<Relationships xmlns="{RELATIONSHIPS}">'
            f'<Relationship Id="doc" Type="{OFFICE_RELS}/officeDocument" Target="word/document.xml"/>'
            f"{signature}</Relationships>",
        )
        archive.writestr(
            "word/document.xml",
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'xmlns:r="{OFFICE_RELS}"><w:body><w:p><w:r><w:drawing/>'
            "</w:r></w:p><w:sectPr/></w:body></w:document>",
        )
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<Relationships xmlns="{RELATIONSHIPS}">'
            f'<Relationship Id="image" Type="{OFFICE_RELS}/image" '
            'Target="media/image1.png"/></Relationships>',
        )
        archive.writestr("word/media/image1.png", image_bytes.getvalue())
        archive.writestr("word/settings.xml", b"settings-must-not-change")
        if signed:
            archive.writestr("_xmlsignatures/origin.sigs", b"signed")
        if macro:
            archive.writestr("word/vbaProject.bin", b"macro-must-not-change")


class DocxImageCompactorTest(unittest.TestCase):
    def test_compresses_only_referenced_media_and_preserves_package_parts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.docm"
            output = root / "output.docm"
            make_docx(source, macro=True)

            result = compact_docx(
                source,
                0.3,
                output=output,
                validate_render=False,
                logger=lambda _: None,
            )

            self.assertLess(output.stat().st_size, source.stat().st_size)
            self.assertLessEqual(output.stat().st_size, 300_000)
            with ZipFile(source) as before, ZipFile(output) as after:
                self.assertEqual(set(before.namelist()), set(after.namelist()))
                for name in set(before.namelist()) - {"word/media/image1.png"}:
                    self.assertEqual(before.read(name), after.read(name), name)
                self.assertLess(
                    len(after.read("word/media/image1.png")),
                    len(before.read("word/media/image1.png")),
                )
            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            self.assertEqual(report["target"]["status"], "met")
            self.assertTrue(report["presentation"]["macro_parts_preserved"])
            self.assertTrue(result["report_path"].with_suffix(".md").is_file())

    def test_refuses_signed_documents(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "signed.docx"
            make_docx(source, signed=True)
            with self.assertRaisesRegex(RuntimeError, "Digitally signed"):
                compact_docx(source, 0.3, logger=lambda _: None)

    def test_refuses_invalid_packages_and_source_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            invalid = root / "invalid.docx"
            invalid.write_bytes(b"not a Word package")
            with self.assertRaisesRegex(RuntimeError, "Encrypted or invalid"):
                compact_docx(invalid, 0.001, logger=lambda _: None)

            source = root / "valid.docx"
            make_docx(source)
            with self.assertRaisesRegex(ValueError, "must not overwrite"):
                compact_docx(source, 0.3, output=source, logger=lambda _: None)

    def test_forced_mode_requires_failed_safe_output_and_confirmation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "valid.docx"
            make_docx(source)
            with self.assertRaisesRegex(ValueError, "explicit confirmation"):
                compact_docx(
                    source,
                    0.3,
                    output=root / "forced.docx",
                    forced=True,
                    logger=lambda _: None,
                )

            safe_output = root / "safe.docx"
            safe_output.write_bytes(b"small")
            with self.assertRaisesRegex(ValueError, "safe output above target"):
                compact_docx(
                    source,
                    0.3,
                    output=root / "forced.docx",
                    forced=True,
                    safe_output=safe_output,
                    confirm_forced=True,
                    logger=lambda _: None,
                )

    def test_source_under_target_is_skipped_without_duplicate_document(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "small.docx"
            make_docx(source)

            result = compact_docx(source, 10, logger=lambda _: None)

            self.assertTrue(result["skipped"])
            self.assertEqual(result["output"], source.resolve())
            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            self.assertEqual(report["target"]["status"], "source_already_meets")
            self.assertFalse(next(source.parent.glob("*_compressed_*"), None))

    @unittest.skipUnless(Path(resolve_soffice_path()).is_file(), "LibreOffice required")
    def test_real_docx_pdf_layout_validation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = Path(__file__).parent / "fixtures" / "smart_target_sample.docx"
            output = root / "output.docx"

            result = compact_docx(source, 0.2, output=output, logger=lambda _: None)

            validation = json.loads(result["report_path"].read_text(encoding="utf-8"))[
                "presentation"
            ]["render_validation"]
            self.assertEqual(validation["status"], "passed")
            self.assertGreaterEqual(validation["minimum_page_ssim"], 0.985)


if __name__ == "__main__":
    unittest.main()
