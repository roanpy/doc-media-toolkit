from __future__ import annotations

import json
import io
import tarfile
import tempfile
import unittest
from argparse import Namespace
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
        self.assertNotIn(str(root.parent), json.dumps(result))

    def test_release_report_redacts_build_host_paths(self) -> None:
        private_path = Path.home() / ".cache" / "uv" / "build"
        args = Namespace(
            skip_pip_audit=True,
            public_binary=False,
            with_sbom=None,
            skip_sbom=False,
            dist_dir=Path(tempfile.gettempdir()) / "candidate",
            evidence=None,
        )
        with (
            patch.object(
                release_audit,
                "check_git_state",
                return_value={"name": "git provenance", "status": "pass"},
            ),
            patch.object(
                release_audit,
                "check_uv_lock",
                return_value={
                    "name": "uv lock",
                    "status": "pass",
                    "detail": str(private_path),
                },
            ),
            patch.object(
                release_audit,
                "check_dist_artifacts",
                return_value={
                    "name": "packaged artifacts",
                    "status": "pass",
                    "detail": str(args.dist_dir),
                },
            ),
        ):
            report = release_audit.run_all_checks(args)

        serialized = json.dumps(report)
        self.assertNotIn(str(Path.home()), serialized)
        self.assertNotIn(str(Path(tempfile.gettempdir())), serialized)
        self.assertIn("<home>", serialized)
        self.assertIn("<temp>", serialized)

    def test_public_binary_gate_fails_closed_without_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            result = release_audit.check_public_binary_evidence(None, Path(temp_dir))
        self.assertEqual(result["status"], "fail")
        self.assertIn("--evidence", result["detail"])

    def test_public_binary_sidecars_require_real_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            empty = root / "empty.json"
            empty.write_text("{}", encoding="utf-8")
            sidecar = {"path": empty.name, "sha256": release_audit._sha256(empty)}

            self.assertTrue(
                release_audit._validate_sbom(
                    root,
                    sidecar,
                    label="sbom",
                    artifact_sha256="0" * 64,
                )
            )
            self.assertTrue(
                release_audit._validate_native_inventory(
                    root,
                    sidecar,
                    label="native_inventory",
                    artifact_sha256="0" * 64,
                )
            )
            self.assertTrue(
                release_audit._validate_malware_report(
                    root,
                    sidecar,
                    label="malware",
                    artifact_sha256="0" * 64,
                )
            )

    def test_public_binary_audit_cannot_skip_dependency_scan(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "artifact").write_bytes(b"artifact")
            args = Namespace(
                skip_pip_audit=True,
                public_binary=True,
                with_sbom=None,
                skip_sbom=False,
                dist_dir=root,
                evidence=None,
            )
            with (
                patch.object(
                    release_audit,
                    "check_git_state",
                    return_value={"name": "git provenance", "status": "pass"},
                ),
                patch.object(
                    release_audit,
                    "check_uv_lock",
                    return_value={"name": "uv lock", "status": "pass"},
                ),
            ):
                report = release_audit.run_all_checks(args)

        pip_check = next(
            check for check in report["checks"] if check["name"] == "pip-audit"
        )
        self.assertEqual(pip_check["status"], "fail")

    def test_public_binary_gate_accepts_complete_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)

            def sidecar(name: str, content: bytes = b"verified") -> dict[str, str]:
                path = root / name
                path.write_bytes(content)
                return {"path": name, "sha256": release_audit._sha256(path)}

            def json_sidecar(
                name: str,
                value: dict[str, object],
                *,
                artifact_sha256: str | None = None,
            ) -> dict[str, str]:
                descriptor = sidecar(name, json.dumps(value).encode("utf-8"))
                if artifact_sha256 is not None:
                    descriptor["artifact_sha256"] = artifact_sha256
                return descriptor

            def source_sidecar(name: str) -> dict[str, str]:
                path = root / name
                with tarfile.open(path, "w:gz") as archive:
                    for member_name in (
                        "source-root/build_ffmpeg_runtime.sh",
                        "source-root/changes.diff",
                        "source-root/SHA256SUMS",
                        "source-root/BUILD-INFO.txt",
                        "source-root/sources/ffmpeg-8.1.2.tar.xz",
                        "source-root/sources/x264-b35605.tar.bz2",
                        "source-root/sources/zlib-1.3.2.tar.xz",
                    ):
                        info = tarfile.TarInfo(member_name)
                        info.size = 1
                        archive.addfile(info, fileobj=io.BytesIO(b"x"))
                return {"path": name, "sha256": release_audit._sha256(path)}

            artifact = sidecar("Doc-Media-Toolkit-macOS-arm64.dmg", b"dmg")
            evidence = {
                "schema": "doc-media-toolkit.public-binary-evidence.v1",
                "version": release_audit.PROJECT_VERSION,
                "artifacts": [
                    {
                        **artifact,
                        "platform": "macos",
                        "architecture": "arm64",
                        "package_type": "dmg",
                        "signature": {
                            "status": "valid",
                            "type": "developer-id",
                            "report": sidecar("codesign.txt"),
                        },
                        "notarization": {
                            "status": "valid",
                            "report": sidecar("notarization.txt"),
                        },
                        "malware_scan": {
                            "status": "clean",
                            "report": json_sidecar(
                                "malware.json",
                                {
                                    "schema": "doc-media-toolkit.malware-scan.v1",
                                    "status": "clean",
                                    "scanner": "test-scanner",
                                    "exit_code": 0,
                                    "artifact_sha256": artifact["sha256"],
                                },
                            ),
                        },
                        "sbom": json_sidecar(
                            "sbom.cdx.json",
                            {
                                "bomFormat": "CycloneDX",
                                "specVersion": "1.5",
                                "components": [{"name": "test-component"}],
                            },
                            artifact_sha256=artifact["sha256"],
                        ),
                        "native_inventory": json_sidecar(
                            "native-inventory.json",
                            {
                                "schema": "doc-media-toolkit.native-inventory.v1",
                                "entries": [
                                    {
                                        "path": "Contents/Frameworks/test.dylib",
                                        "sha256": "0" * 64,
                                    }
                                ],
                            },
                            artifact_sha256=artifact["sha256"],
                        ),
                        "ffmpeg": {
                            "bundled": True,
                            "version": "8.1.2",
                            "configuration": (
                                "--enable-gpl --enable-version3 --enable-libx264"
                            ),
                            "license": "GPL-3.0-or-later",
                            "corresponding_source": source_sidecar(
                                "ffmpeg-corresponding-source.tar.gz"
                            ),
                        },
                    }
                ],
            }
            evidence_path = root / "public-binary-evidence.json"
            evidence_path.write_text(json.dumps(evidence), encoding="utf-8")
            result = release_audit.check_public_binary_evidence(evidence_path, root)

        self.assertEqual(result["status"], "pass")

    def test_public_binary_gate_rejects_windows_onefile_and_missing_source(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            artifact = root / "Doc-Media-Toolkit-windows-x64.exe"
            artifact.write_bytes(b"exe")
            evidence = {
                "schema": "doc-media-toolkit.public-binary-evidence.v1",
                "version": release_audit.PROJECT_VERSION,
                "artifacts": [
                    {
                        "path": artifact.name,
                        "sha256": release_audit._sha256(artifact),
                        "platform": "windows",
                        "package_type": "onefile",
                        "ffmpeg": {
                            "bundled": True,
                            "version": "8.1.2",
                            "configuration": "--enable-gpl",
                            "license": "GPL-3.0-or-later",
                        },
                    }
                ],
            }
            evidence_path = root / "evidence.json"
            evidence_path.write_text(json.dumps(evidence), encoding="utf-8")
            result = release_audit.check_public_binary_evidence(evidence_path, root)

        self.assertEqual(result["status"], "fail")
        self.assertTrue(any("onefile" in finding for finding in result["findings"]))
        self.assertTrue(
            any("corresponding_source" in finding for finding in result["findings"])
        )

    def test_corresponding_source_rejects_empty_required_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = Path(temp_dir) / "source.tar.gz"
            with tarfile.open(archive_path, "w:gz") as archive:
                for member_name in (
                    "source-root/build_ffmpeg_runtime.sh",
                    "source-root/changes.diff",
                    "source-root/SHA256SUMS",
                    "source-root/BUILD-INFO.txt",
                    "source-root/sources/ffmpeg-8.1.2.tar.xz",
                    "source-root/sources/x264-b35605.tar.bz2",
                    "source-root/sources/zlib-1.3.2.tar.xz",
                ):
                    info = tarfile.TarInfo(member_name)
                    info.size = 0 if member_name.endswith("BUILD-INFO.txt") else 1
                    archive.addfile(info, fileobj=io.BytesIO(b"x" * info.size))

            findings = release_audit._validate_corresponding_source(
                archive_path, label="source"
            )

        self.assertTrue(any("empty required files" in finding for finding in findings))


if __name__ == "__main__":
    unittest.main()
