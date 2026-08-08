from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts import release_audit, run_compression_benchmark


class LocalReleaseToolsTest(unittest.TestCase):
    def test_manifest_defaults_are_applied_and_sensitive_extras_are_dropped(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "sample.pptx"
            sample.write_bytes(b"sample")
            spec = run_compression_benchmark.SampleSpec.from_manifest_entry(
                {
                    "path": str(sample),
                    "token": "must-not-be-copied",
                    "meta_case": "photo-heavy",
                },
                default_encoder="gpu",
                default_quality_mode="forced",
            )

        self.assertEqual(spec.encoder, "gpu")
        self.assertEqual(spec.quality_mode, "forced")
        self.assertEqual(spec.extra, {"meta_case": "photo-heavy"})

    def test_manifest_rejects_relative_paths_and_invalid_thresholds(self) -> None:
        with self.assertRaisesRegex(ValueError, "absolute"):
            run_compression_benchmark.SampleSpec.from_manifest_entry(
                {"path": "relative.pptx"}
            )
        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "sample.pptx"
            sample.write_bytes(b"sample")
            with self.assertRaisesRegex(ValueError, "video_ssim_threshold"):
                run_compression_benchmark.SampleSpec.from_manifest_entry(
                    {"path": str(sample), "video_ssim_threshold": 1.1}
                )

    def test_benchmark_reads_path_report_and_redacts_absolute_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sample = root / "source.pptx"
            output = root / "bench.pptx"
            report_path = root / "bench.report.json"
            sample.write_bytes(b"source")
            output.write_bytes(b"output")
            report_path.write_text(
                json.dumps(
                    {
                        "target": {
                            "target_bytes": 10,
                            "actual_bytes": 8,
                            "delta_bytes": -2,
                            "target_ratio": 0.8,
                            "status": "met",
                        },
                        "presentation": {
                            "target_capacity_attempts": [
                                {"kind": "initial"},
                                {"kind": "correction"},
                            ]
                        },
                        "videos": [
                            {
                                "status": "encoded_gpu",
                                "quality_status": "below_threshold",
                            }
                        ],
                        "images": [],
                    }
                ),
                encoding="utf-8",
            )
            spec = run_compression_benchmark.SampleSpec(path=str(sample))
            with patch.object(
                run_compression_benchmark,
                "compact_input_path",
                return_value={
                    "output_pptx": output,
                    "report_path": report_path,
                    "skipped": False,
                },
            ):
                result = run_compression_benchmark.run_one_sample(
                    spec, root / "results", quiet=True
                )

        self.assertEqual(result["input_path"], "source.pptx")
        self.assertEqual(result["output_path"], "bench.pptx")
        self.assertEqual(result["report_path"], "bench.report.json")
        self.assertEqual(result["correction_rounds"], 1)
        self.assertTrue(result["assets"]["gpu_used"])
        self.assertEqual(result["assets"]["below_threshold_assets"], 1)

    def test_missing_sample_is_recorded_without_aborting_the_batch(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            spec = run_compression_benchmark.SampleSpec(path=str(root / "missing.pptx"))
            result = run_compression_benchmark.run_one_sample(
                spec, root / "results", quiet=True
            )

        self.assertEqual(result["input_path"], "missing.pptx")
        self.assertIn("Sample file not found", result["error"])

    def test_empty_dist_is_not_a_release_pass(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            result = release_audit.check_dist_artifacts(Path(temp_dir))
        self.assertEqual(result["status"], "fail")

    def test_git_state_marks_dirty_worktree_as_release_failure(self) -> None:
        command_results = [
            {"returncode": 0, "stdout": "abc123\n", "stderr": ""},
            {"returncode": 0, "stdout": "main\n", "stderr": ""},
            {"returncode": 0, "stdout": " M README.md\n", "stderr": ""},
        ]
        with (
            patch.object(release_audit.shutil, "which", return_value="/usr/bin/git"),
            patch.object(release_audit, "_run", side_effect=command_results),
        ):
            result = release_audit.check_git_state()
        self.assertEqual(result["status"], "fail")
        self.assertTrue(result["dirty"])
        self.assertEqual(result["commit"], "abc123")

    def test_dist_artifacts_include_sha256(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            artifact = root / "app.dmg"
            artifact.write_bytes(b"artifact")
            result = release_audit.check_dist_artifacts(root)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["files"][0]["path"], "app.dmg")
        self.assertEqual(
            result["files"][0]["sha256"],
            "c7c5c1d70c5dec4416ab6158afd0b223ef40c29b1dc1f97ed9428b94d4cadb1c",
        )


if __name__ == "__main__":
    unittest.main()
