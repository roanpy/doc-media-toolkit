from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx_quality_audit import QualityAuditWorker
from pptx_tools.video_manager import _ssim_videos


class _FakeProcess:
    def __init__(
        self, stderr_text: str = "SSIM All:0.987654\n", returncode: int = 0
    ) -> None:
        self._stderr_text = stderr_text
        self.returncode = returncode

    def communicate(self, timeout: int | None = None) -> tuple[None, str]:
        return None, self._stderr_text


class QualityAuditFfmpegResolutionTest(unittest.TestCase):
    def test_evaluate_single_uses_resolved_ffmpeg_binary(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            original_dir = root / "original"
            compressed_dir = root / "compressed"
            original_dir.mkdir()
            compressed_dir.mkdir()

            media_rel = Path("ppt/media/sample.mp4")
            output_rel = Path("ppt/media/sample_out.mp4")
            (original_dir / media_rel).parent.mkdir(parents=True, exist_ok=True)
            (compressed_dir / output_rel).parent.mkdir(parents=True, exist_ok=True)
            (original_dir / media_rel).write_bytes(b"orig")
            (compressed_dir / output_rel).write_bytes(b"comp")

            worker = QualityAuditWorker(
                Path("input.pptx"), Path("output.pptx"), Path("report.json")
            )
            asset = {
                "media_path": str(media_rel),
                "output_media_path": str(output_rel),
                "is_video": True,
            }

            with (
                patch(
                    "pptx_quality_audit.resolve_binary",
                    return_value="/custom/bin/ffmpeg",
                ) as resolve_mock,
                patch(
                    "pptx_quality_audit.start_process", return_value=_FakeProcess()
                ) as popen_mock,
            ):
                result = worker._evaluate_single(asset, original_dir, compressed_dir)

            self.assertEqual(result.status, "success")
            self.assertAlmostEqual(result.ssim or 0.0, 0.987654, places=6)
            resolve_mock.assert_called_once_with("ffmpeg")
            called_cmd = popen_mock.call_args.args[0]
            self.assertEqual(called_cmd[0], "/custom/bin/ffmpeg")
            self.assertIn(
                "setpts=PTS-STARTPTS", called_cmd[called_cmd.index("-lavfi") + 1]
            )

    def test_library_ssim_normalizes_container_timestamps(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            candidate = root / "candidate.mp4"
            reference = root / "reference.avi"
            candidate.touch()
            reference.touch()
            with (
                patch(
                    "pptx_video_compactor.resolve_binary",
                    return_value="/custom/bin/ffmpeg",
                ),
                patch(
                    "pptx_output_watermark.process_utils.start_process",
                    return_value=_FakeProcess(),
                ) as start_process,
                patch("pptx_output_watermark.process_utils.finish_process"),
            ):
                result = _ssim_videos(candidate, reference)

            self.assertAlmostEqual(result or 0.0, 0.987654, places=6)
            command = start_process.call_args.args[0]
            self.assertIn("setpts=PTS-STARTPTS", command[command.index("-lavfi") + 1])

    def test_evaluate_single_reports_ffmpeg_decode_failure(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            original_dir = root / "original"
            compressed_dir = root / "compressed"
            original_dir.mkdir()
            compressed_dir.mkdir()

            media_rel = Path("ppt/media/sample.wmv")
            output_rel = Path("ppt/media/sample_out.wmv")
            (original_dir / media_rel).parent.mkdir(parents=True, exist_ok=True)
            (compressed_dir / output_rel).parent.mkdir(parents=True, exist_ok=True)
            (original_dir / media_rel).write_bytes(b"orig")
            (compressed_dir / output_rel).write_bytes(b"comp")

            worker = QualityAuditWorker(
                Path("input.pptx"), Path("output.pptx"), Path("report.json")
            )
            asset = {
                "media_path": str(media_rel),
                "output_media_path": str(output_rel),
                "is_video": True,
            }

            with (
                patch(
                    "pptx_quality_audit.resolve_binary",
                    return_value="/custom/bin/ffmpeg",
                ),
                patch(
                    "pptx_quality_audit.start_process",
                    return_value=_FakeProcess(
                        "Unsupported codec\nConversion failed!\n", returncode=1
                    ),
                ),
            ):
                result = worker._evaluate_single(asset, original_dir, compressed_dir)

            self.assertEqual(result.status, "error")
            self.assertIsNone(result.ssim)
            self.assertIn("无法计算视频质量评分", result.error or "")
            self.assertIn("WMV", result.error or "")
            self.assertIn("Unsupported codec", result.error or "")

    def test_pixel_identical_transparent_png_recovers_invalid_ssim(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            original_dir = root / "original"
            compressed_dir = root / "compressed"
            media_rel = Path("ppt/media/transparent.png")
            original = original_dir / media_rel
            compressed = compressed_dir / media_rel
            original.parent.mkdir(parents=True)
            compressed.parent.mkdir(parents=True)
            image = Image.new("RGBA", (8, 8), (0, 0, 0, 0))
            image.save(original, compress_level=1)
            image.save(compressed, optimize=True, compress_level=9)

            worker = QualityAuditWorker(
                Path("input.pptx"), Path("output.pptx"), Path("report.json")
            )
            asset = {
                "media_path": str(media_rel),
                "output_media_path": str(media_rel),
                "is_video": False,
            }
            with (
                patch(
                    "pptx_quality_audit.resolve_binary",
                    return_value="/custom/bin/ffmpeg",
                ),
                patch(
                    "pptx_quality_audit.start_process",
                    return_value=_FakeProcess("SSIM All:0.000000\n"),
                ),
            ):
                result = worker._evaluate_single(asset, original_dir, compressed_dir)

            self.assertEqual(result.status, "success")
            self.assertEqual(result.ssim, 1.0)

    def test_asset_paths_cannot_escape_audit_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(ValueError, "超出评估目录"):
                QualityAuditWorker._asset_target(Path(temp_dir), "../escaped.bin")

    def test_cancel_stops_registered_audit_processes(self) -> None:
        worker = QualityAuditWorker(
            Path("input.pptx"), Path("output.pptx"), Path("report.json")
        )
        process = _FakeProcess()
        worker._active_procs.add(process)

        with patch("pptx_quality_audit.kill_process") as kill_process:
            worker.cancel()

        self.assertTrue(worker.cancel_requested)
        kill_process.assert_called_once_with(process)


if __name__ == "__main__":
    unittest.main()
