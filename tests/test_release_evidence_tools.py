from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts import generate_native_inventory, scan_release_artifact


class ReleaseEvidenceToolsTest(unittest.TestCase):
    def test_native_inventory_records_relative_hashes_without_machine_paths(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Doc Media Toolkit.app"
            binary = root / "Contents" / "MacOS" / "Doc Media Toolkit"
            library = root / "Contents" / "Frameworks" / "libexample.dylib"
            binary.parent.mkdir(parents=True)
            library.parent.mkdir(parents=True)
            binary.write_bytes(b"executable")
            library.write_bytes(b"library")
            binary.chmod(0o755)
            artifact = root.parent / "Doc-Media-Toolkit-macOS-arm64.dmg"
            artifact.write_bytes(b"package")

            report = generate_native_inventory.build_inventory(
                root,
                platform_name="macos",
                architecture="arm64",
                artifact=artifact,
            )

        self.assertEqual(report["schema"], "doc-media-toolkit.native-inventory.v1")
        self.assertEqual(report["platform"], "macos")
        self.assertEqual(report["architecture"], "arm64")
        paths = {entry["path"] for entry in report["entries"]}
        self.assertEqual(
            paths,
            {
                "Contents/MacOS/Doc Media Toolkit",
                "Contents/Frameworks/libexample.dylib",
            },
        )
        self.assertTrue(all(len(entry["sha256"]) == 64 for entry in report["entries"]))
        self.assertEqual(report["artifact"], artifact.name)
        self.assertEqual(len(report["artifact_sha256"]), 64)
        self.assertNotIn(temp_dir, str(report))

    def test_malware_scan_is_fail_closed_without_scanner(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            artifact = Path(temp_dir) / "artifact.zip"
            artifact.write_bytes(b"artifact")
            with patch.object(
                scan_release_artifact, "find_scanner", return_value=(None, None)
            ):
                report = scan_release_artifact.scan_artifact(artifact)

        self.assertEqual(report["schema"], "doc-media-toolkit.malware-scan.v1")
        self.assertEqual(report["status"], "unavailable")
        self.assertIsNone(report["exit_code"])
        self.assertIsNone(report["scanner"])

    def test_malware_scan_redacts_artifact_path_and_accepts_clean_exit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            artifact = Path(temp_dir) / "artifact.zip"
            artifact.write_bytes(b"artifact")
            with (
                patch.object(
                    scan_release_artifact,
                    "find_scanner",
                    return_value=("/usr/local/bin/clamscan", "clamscan"),
                ),
                patch.object(
                    scan_release_artifact,
                    "_run",
                    return_value=(0, f"clean {artifact}", ""),
                ),
            ):
                report = scan_release_artifact.scan_artifact(artifact)

        self.assertEqual(report["status"], "clean")
        self.assertEqual(report["exit_code"], 0)
        self.assertNotIn(temp_dir, report["stdout_tail"])
        self.assertIn("<artifact>", report["stdout_tail"])


if __name__ == "__main__":
    unittest.main()
