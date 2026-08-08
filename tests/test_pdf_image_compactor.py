from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageChops, ImageOps
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DecodedStreamObject,
    DictionaryObject,
    NameObject,
    NumberObject,
)
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from pdf_image_compactor import compact_pdf


def make_scanned_pdf(path: Path) -> None:
    image_path = path.with_suffix(".jpg")
    gradient = Image.linear_gradient("L").resize((1400, 1800))
    texture = Image.effect_noise((1400, 1800), 2)
    image = ImageOps.colorize(
        ImageChops.add(gradient, texture, scale=1.2),
        (25, 60, 90),
        (235, 220, 190),
    )
    image.save(image_path, format="JPEG", quality=100, subsampling=0)

    document = canvas.Canvas(str(path), pagesize=letter)
    document.setTitle("Scan with preserved structures")
    document.bookmarkPage("page-one")
    document.addOutlineEntry("Page one", "page-one")
    document.drawImage(str(image_path), 0, 0, width=612, height=792)
    text = document.beginText(36, 740)
    text.setTextRenderMode(3)
    text.textLine("searchable OCR text must stay intact")
    document.drawText(text)
    document.linkURL("https://example.com", (36, 36, 180, 54))
    document.acroForm.textfield(
        name="reviewer",
        value="Example User",
        x=400,
        y=30,
        width=100,
        height=20,
    )
    document.showPage()
    document.drawImage(str(image_path), 0, 0, width=612, height=792)
    document.showPage()
    document.setFont("Helvetica", 12)
    document.drawString(36, 740, "digital page text must stay vector text")
    document.drawImage(str(image_path), 36, 420, width=180, height=240)
    document.drawInlineImage(image, 300, 420, width=90, height=120)
    document.save()
    image_path.unlink()

    reader = PdfReader(path)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    writer.add_attachment("evidence.txt", b"attachment-must-not-change")
    with path.open("wb") as stream:
        writer.write(stream)


def mark_signed(source: Path, output: Path) -> None:
    reader = PdfReader(source)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    writer.root_object[NameObject("/Perms")] = DictionaryObject(
        {NameObject("/DocMDP"): DictionaryObject()}
    )
    with output.open("wb") as stream:
        writer.write(stream)


def make_bilevel_scan(path: Path) -> None:
    width, height = 1728, 2200
    row_bytes = width // 8
    pixels = bytearray()
    for row in range(height):
        pixels.extend(bytes([0 if row % 44 < 2 else 255]) * row_bytes)

    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    image = DecodedStreamObject()
    image.set_data(bytes(pixels))
    image.update(
        {
            NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Image"),
            NameObject("/Width"): NumberObject(width),
            NameObject("/Height"): NumberObject(height),
            NameObject("/ColorSpace"): NameObject("/DeviceGray"),
            NameObject("/BitsPerComponent"): NumberObject(1),
        }
    )
    image_reference = writer._add_object(image)
    content = DecodedStreamObject()
    content.set_data(b"q 612 0 0 792 0 0 cm /Im0 Do Q")
    page[NameObject("/Contents")] = writer._add_object(content)
    page[NameObject("/Resources")] = DictionaryObject(
        {
            NameObject("/XObject"): DictionaryObject(
                {NameObject("/Im0"): image_reference}
            )
        }
    )
    with path.open("wb") as stream:
        writer.write(stream)


class PdfImageCompactorTest(unittest.TestCase):
    def test_bilevel_scan_uses_lossless_ccitt_group4(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "bilevel.pdf"
            output = root / "compressed.pdf"
            make_bilevel_scan(source)

            result = compact_pdf(source, 0.2, output=output, logger=lambda _: None)

            self.assertLess(output.stat().st_size, source.stat().st_size)
            image_object = (
                PdfReader(output).pages[0].images[0].indirect_reference.get_object()
            )
            self.assertIn("/CCITTFaxDecode", str(image_object["/Filter"]))
            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            self.assertEqual(
                report["presentation"]["page_analysis"][0]["kind"], "scanned"
            )

    def test_compresses_scan_and_preserves_text_forms_links_and_attachments(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "scan.pdf"
            output = root / "compressed.pdf"
            make_scanned_pdf(source)
            target_mb = source.stat().st_size * 0.75 / 1_000_000

            result = compact_pdf(
                source, target_mb, output=output, logger=lambda _: None
            )

            before = PdfReader(source)
            after = PdfReader(output)
            self.assertEqual(
                before.pages[0].extract_text(), after.pages[0].extract_text()
            )
            self.assertEqual(
                before.get_fields()["reviewer"]["/V"],
                after.get_fields()["reviewer"]["/V"],
            )
            self.assertEqual(before.attachments, after.attachments)
            self.assertEqual(
                len(before.pages[0]["/Annots"]), len(after.pages[0]["/Annots"])
            )
            self.assertLess(output.stat().st_size, source.stat().st_size)
            report = json.loads(result["report_path"].read_text(encoding="utf-8"))
            self.assertEqual(
                report["presentation"]["render_validation"]["status"], "passed"
            )
            self.assertEqual(
                report["presentation"]["render_validation"]["renderers"],
                ["poppler-cairo", "pdfium"],
            )
            self.assertGreaterEqual(
                report["presentation"]["render_validation"]["minimum_page_ssim"],
                0.985,
            )
            self.assertEqual(
                [page["kind"] for page in report["presentation"]["page_analysis"]],
                ["mixed", "scanned", "digital"],
            )
            self.assertTrue(
                any(
                    item["reason"] == "Inline image preserved"
                    for item in report["presentation"]["skipped_images"]
                )
            )

    def test_refuses_signed_and_encrypted_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pdf"
            make_scanned_pdf(source)
            target_mb = source.stat().st_size * 0.75 / 1_000_000

            signed = root / "signed.pdf"
            mark_signed(source, signed)
            with self.assertRaisesRegex(RuntimeError, "Digitally signed"):
                compact_pdf(signed, target_mb, logger=lambda _: None)

            reader = PdfReader(source)
            writer = PdfWriter()
            writer.append_pages_from_reader(reader)
            writer.encrypt("secret")
            encrypted = root / "encrypted.pdf"
            with encrypted.open("wb") as stream:
                writer.write(stream)
            with self.assertRaisesRegex(RuntimeError, "Encrypted PDFs"):
                compact_pdf(encrypted, target_mb, logger=lambda _: None)

    def test_forced_mode_requires_failed_safe_output_and_confirmation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pdf"
            make_scanned_pdf(source)
            target_mb = source.stat().st_size * 0.75 / 1_000_000
            with self.assertRaisesRegex(ValueError, "explicit confirmation"):
                compact_pdf(
                    source,
                    target_mb,
                    output=root / "forced.pdf",
                    forced=True,
                    logger=lambda _: None,
                )

            safe_output = root / "safe.pdf"
            safe_output.write_bytes(b"small")
            with self.assertRaisesRegex(ValueError, "safe output above target"):
                compact_pdf(
                    source,
                    target_mb,
                    output=root / "forced.pdf",
                    forced=True,
                    safe_output=safe_output,
                    confirm_forced=True,
                    logger=lambda _: None,
                )

    def test_source_under_target_is_skipped_without_duplicate_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            make_scanned_pdf(source)
            result = compact_pdf(
                source,
                source.stat().st_size / 1_000_000 + 0.1,
                logger=lambda _: None,
            )
            self.assertTrue(result["skipped"])
            self.assertEqual(result["output"], source.resolve())
            self.assertTrue(result["report_path"].is_file())


if __name__ == "__main__":
    unittest.main()
