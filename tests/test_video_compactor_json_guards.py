from __future__ import annotations

import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx_video_compactor import (
    VideoAsset,
    ffprobe_json,
    load_json_file,
    parse_json_payload,
    validate_encoded_asset,
)


class VideoCompactorJsonGuardsTest(unittest.TestCase):
    def test_parse_json_payload_rejects_none(self) -> None:
        with self.assertRaises(SystemExit) as ctx:
            parse_json_payload(None, source="ffprobe")

        self.assertIn("did not return JSON output", str(ctx.exception))

    def test_ffprobe_json_rejects_empty_stdout(self) -> None:
        fake_result = subprocess.CompletedProcess(
            args=["ffprobe"],
            returncode=0,
            stdout=None,
            stderr="",
        )

        with patch("pptx_video_compactor.run", return_value=fake_result):
            with self.assertRaises(SystemExit) as ctx:
                ffprobe_json(Path("sample.mp4"))

        self.assertIn("ffprobe", str(ctx.exception))

    def test_load_json_file_rejects_empty_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            json_path = Path(temp_dir) / "report.json"
            json_path.write_text("", encoding="utf-8")

            with self.assertRaises(SystemExit) as ctx:
                load_json_file(json_path, source="Report")

        self.assertIn("returned empty JSON output", str(ctx.exception))

    def test_encoded_video_rejects_unexpected_frame_loss(self) -> None:
        asset = VideoAsset(
            media_path="sample.avi",
            zip_size=100,
            duration_sec=4.8,
            width=720,
            height=1080,
            has_audio=True,
            original_fps=30.0,
            original_frame_count=144,
            target_fps=30.0,
        )
        probe = {
            "streams": [
                {
                    "codec_type": "video",
                    "width": 720,
                    "height": 1080,
                    "duration": "4.8",
                    "nb_frames": "142",
                },
                {"codec_type": "audio"},
            ],
            "format": {"duration": "4.8"},
        }
        with patch("pptx_video_compactor.ffprobe_json", return_value=probe):
            with self.assertRaisesRegex(ValueError, "frame count mismatch"):
                validate_encoded_asset(asset, Path("sample.mp4"), 720, 1080)

    def test_encoded_video_allows_one_frame_container_tolerance(self) -> None:
        asset = VideoAsset(
            media_path="sample.mp4",
            zip_size=100,
            duration_sec=4.8,
            width=720,
            height=1080,
            has_audio=False,
            original_fps=30.0,
            original_frame_count=144,
            target_fps=30.0,
        )
        probe = {
            "streams": [
                {
                    "codec_type": "video",
                    "width": 720,
                    "height": 1080,
                    "duration": "4.8",
                    "nb_frames": "143",
                }
            ],
            "format": {"duration": "4.8"},
        }
        with patch("pptx_video_compactor.ffprobe_json", return_value=probe):
            validate_encoded_asset(asset, Path("sample.mp4"), 720, 1080)


if __name__ == "__main__":
    unittest.main()
