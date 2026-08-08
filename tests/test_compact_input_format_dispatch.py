from __future__ import annotations

import argparse
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import docx_image_compactor
import pdf_image_compactor
import xlsx_image_compactor
from pptx_video_compactor import CancelledError, compact_input_path


def make_args(
    input_path: Path,
    *,
    target_size_mb: float | None = 1.0,
    image_profile: str = "high",
    image_ssim_threshold: float = 0.99,
    output: Path | None = None,
    quality_mode: str = "safe",
) -> argparse.Namespace:
    return argparse.Namespace(
        input_pptx=input_path,
        target_size_mb=target_size_mb,
        image_profile=image_profile,
        image_ssim_threshold=image_ssim_threshold,
        output=output,
        quality_mode=quality_mode,
    )


def backend_result(source: Path, *, skipped: bool = False) -> dict:
    return {
        "input": source,
        "output": source
        if skipped
        else source.with_name(f"{source.stem}_out{source.suffix}"),
        "report_path": source.with_suffix(".report.json"),
        "skipped": skipped,
    }


class CompactInputFormatDispatchTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.base = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_routes_ooxml_and_pdf_suffixes(self) -> None:
        cases = [
            ("sample.docx", "docx_image_compactor", "compact_docx"),
            ("sample.docm", "docx_image_compactor", "compact_docx"),
            ("sample.xlsx", "xlsx_image_compactor", "compact_xlsx"),
            ("sample.xlsm", "xlsx_image_compactor", "compact_xlsx"),
            ("sample.pdf", "pdf_image_compactor", "compact_pdf"),
        ]
        for name, module_name, func_name in cases:
            with self.subTest(name=name), patch(f"{module_name}.{func_name}") as mock:
                mock.side_effect = lambda source, *_a, **_kw: backend_result(source)
                source = self.base / name
                result = compact_input_path(make_args(source), logger=lambda _: None)
                mock.assert_called_once()
                self.assertEqual(result["input_pptx"], source.resolve())
                self.assertEqual(
                    result["output_pptx"],
                    source.resolve().with_name(f"{source.stem}_out{source.suffix}"),
                )
                self.assertFalse(result["skipped"])
                self.assertEqual(result["reason"], "")

    def test_parameter_mapping(self) -> None:
        source = self.base / "sample.docx"
        with patch("docx_image_compactor.compact_docx") as mock:
            mock.side_effect = lambda source, *_a, **_kw: backend_result(source)
            compact_input_path(
                make_args(
                    source,
                    target_size_mb=2.5,
                    image_profile="balanced",
                    image_ssim_threshold=0.97,
                ),
                logger=lambda _: None,
            )
        args, kwargs = mock.call_args
        self.assertEqual(args[0], source.resolve())
        self.assertEqual(args[1], 2.5)
        self.assertEqual(kwargs["image_profile"], "balanced")
        self.assertEqual(kwargs["image_ssim_threshold"], 0.97)
        self.assertIsNone(kwargs["output"])
        self.assertFalse(kwargs["forced"])
        self.assertFalse(kwargs["confirm_forced"])
        self.assertIsNone(kwargs["safe_output"])
        self.assertNotIn("validate_render", kwargs)
        self.assertNotIn("progress_callback", kwargs)

    def test_target_size_required(self) -> None:
        with patch("docx_image_compactor.compact_docx") as mock:
            with self.assertRaises(ValueError):
                compact_input_path(
                    make_args(self.base / "sample.docx", target_size_mb=None),
                    logger=lambda _: None,
                )
            mock.assert_not_called()

    def test_image_profile_none_rejected(self) -> None:
        with patch("pdf_image_compactor.compact_pdf") as mock:
            with self.assertRaises(ValueError):
                compact_input_path(
                    make_args(self.base / "sample.pdf", image_profile="none"),
                    logger=lambda _: None,
                )
            mock.assert_not_called()

    def test_cancel_callback_short_circuits(self) -> None:
        with patch("xlsx_image_compactor.compact_xlsx") as mock:
            with self.assertRaises(CancelledError):
                compact_input_path(
                    make_args(self.base / "sample.xlsx"),
                    logger=lambda _: None,
                    cancel_callback=lambda: True,
                )
            mock.assert_not_called()

    def test_forced_mode_derives_safe_output(self) -> None:
        source = self.base / "sample.docx"
        with patch("docx_image_compactor.compact_docx") as mock:
            mock.side_effect = lambda source, *_a, **_kw: backend_result(source)
            compact_input_path(
                make_args(source, target_size_mb=1.0, quality_mode="forced"),
                logger=lambda _: None,
            )
        _, kwargs = mock.call_args
        self.assertTrue(kwargs["forced"])
        self.assertTrue(kwargs["confirm_forced"])
        self.assertEqual(
            kwargs["safe_output"],
            docx_image_compactor.default_output_path(
                source.resolve(), 1.0, forced=False
            ),
        )

    def test_skipped_result_gets_reason(self) -> None:
        with patch("pdf_image_compactor.compact_pdf") as mock:
            mock.side_effect = lambda source, *_a, **_kw: backend_result(
                source, skipped=True
            )
            result = compact_input_path(
                make_args(self.base / "sample.pdf"), logger=lambda _: None
            )
        self.assertTrue(result["skipped"])
        self.assertEqual(result["reason"], "already_meets_target")
        self.assertEqual(result["output_pptx"], result["input_pptx"])

    def test_public_default_output_path_wrappers(self) -> None:
        source = (self.base / "sample.docx").resolve()
        safe = docx_image_compactor.default_output_path(source, 1.0)
        forced = docx_image_compactor.default_output_path(source, 1.0, forced=True)
        self.assertEqual(safe.suffix, ".docx")
        self.assertIn("_forced", forced.name)
        self.assertNotIn("_forced", safe.name)
        self.assertEqual(
            xlsx_image_compactor.default_output_path(
                self.base / "sample.xlsx", 0.5
            ).suffix,
            ".xlsx",
        )
        self.assertEqual(
            pdf_image_compactor.default_output_path(
                self.base / "sample.pdf", 0.5
            ).suffix,
            ".pdf",
        )


if __name__ == "__main__":
    unittest.main()
