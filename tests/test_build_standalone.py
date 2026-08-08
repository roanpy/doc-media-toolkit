from __future__ import annotations

import tempfile
import unittest
import sys
from pathlib import Path
from unittest.mock import Mock, patch

from scripts.build_standalone import (
    CLI_HIDDEN_IMPORTS,
    add_bundle_license_files,
    add_optional_ffmpeg_binaries,
    experimental_dist_root,
    notarize_macos_dmg,
    runtime_distribution_names,
    _python_license_file,
)


class StandaloneBuildTests(unittest.TestCase):
    def test_python_license_lookup_supports_nested_runtime_docs(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            nested = root / "share" / "doc" / "python3.12"
            nested.mkdir(parents=True)
            license_file = nested / "LICENSE.txt"
            license_file.write_text("PSF", encoding="utf-8")
            with (
                patch.object(sys, "executable", str(root / "bin" / "python")),
                patch.object(sys, "base_prefix", str(root)),
            ):
                self.assertEqual(_python_license_file(), license_file.resolve())

    def test_runtime_license_inventory_follows_transitive_dependencies(self) -> None:
        names = {name.lower() for name in runtime_distribution_names()}
        self.assertIn("pyside6", names)
        self.assertIn("lxml", names)
        self.assertIn("charset-normalizer", names)
        self.assertNotIn("pyinstaller", names)
        self.assertNotIn("ruff", names)

    def test_bundle_includes_project_runtime_and_dependency_notices(self) -> None:
        project_root = Path(__file__).resolve().parents[1]
        args: list[str] = []
        add_bundle_license_files(args, ":", project_root)
        values = [str(value) for value in args]

        self.assertTrue(any("THIRD_PARTY_NOTICES.md" in value for value in values))
        self.assertTrue(any("LGPL-3.0-only.txt" in value for value in values))
        self.assertTrue(any("pypdfium2" in value for value in values))
        self.assertTrue(any("PyInstaller".lower() in value for value in values))

    def test_cli_bundle_includes_dynamic_subcommands(self) -> None:
        self.assertIn("pptx_tools.video_manager", CLI_HIDDEN_IMPORTS)
        self.assertIn("pptx_tools.image_manager", CLI_HIDDEN_IMPORTS)

    def test_cli_bundle_includes_lazy_document_backends(self) -> None:
        # _compact_document_backend loads these via importlib; PyInstaller
        # cannot discover them statically.
        self.assertIn("docx_image_compactor", CLI_HIDDEN_IMPORTS)
        self.assertIn("pdf_image_compactor", CLI_HIDDEN_IMPORTS)
        self.assertIn("xlsx_image_compactor", CLI_HIDDEN_IMPORTS)
        self.assertIn("pikepdf", CLI_HIDDEN_IMPORTS)

    def test_ffmpeg_bundle_includes_matching_license_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "ffmpeg-8.1"
            binary_dir = root / "bin"
            binary_dir.mkdir(parents=True)
            ffmpeg = binary_dir / "ffmpeg"
            ffprobe = binary_dir / "ffprobe"
            ffmpeg.write_bytes(b"ffmpeg")
            ffprobe.write_bytes(b"ffprobe")
            (root / "LICENSE.md").write_text("license", encoding="utf-8")
            (root / "COPYING.GPLv3").write_text("gpl", encoding="utf-8")
            args: list[str] = []

            with patch(
                "scripts.build_standalone.resolve_binary",
                side_effect=lambda name: {"ffmpeg": ffmpeg, "ffprobe": ffprobe}[name],
            ):
                add_optional_ffmpeg_binaries(
                    args,
                    ":",
                    required=True,
                    max_binary_mb=10,
                )

        self.assertEqual(args.count("--add-binary"), 2)
        self.assertEqual(args.count("--add-data"), 2)

    def test_ffmpeg_bundle_refuses_missing_license(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            ffmpeg = root / "ffmpeg"
            ffprobe = root / "ffprobe"
            ffmpeg.write_bytes(b"ffmpeg")
            ffprobe.write_bytes(b"ffprobe")
            with (
                patch(
                    "scripts.build_standalone.resolve_binary",
                    side_effect=lambda name: {
                        "ffmpeg": ffmpeg,
                        "ffprobe": ffprobe,
                    }[name],
                ),
                self.assertRaisesRegex(SystemExit, "LICENSE/COPYING"),
            ):
                add_optional_ffmpeg_binaries(
                    [],
                    ":",
                    required=True,
                    max_binary_mb=10,
                )

    def test_ffmpeg_bundle_requires_a_license_for_each_distribution(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            ffmpeg = root / "first" / "bin" / "ffmpeg"
            ffprobe = root / "second" / "bin" / "ffprobe"
            ffmpeg.parent.mkdir(parents=True)
            ffprobe.parent.mkdir(parents=True)
            ffmpeg.write_bytes(b"ffmpeg")
            ffprobe.write_bytes(b"ffprobe")
            (ffmpeg.parents[1] / "LICENSE").write_text("license", encoding="utf-8")
            with (
                patch(
                    "scripts.build_standalone.resolve_binary",
                    side_effect=lambda name: {
                        "ffmpeg": ffmpeg,
                        "ffprobe": ffprobe,
                    }[name],
                ),
                self.assertRaisesRegex(SystemExit, "LICENSE/COPYING"),
            ):
                add_optional_ffmpeg_binaries(
                    [],
                    ":",
                    required=True,
                    max_binary_mb=10,
                )

    def test_notarization_is_optional_and_uses_keychain_profile(self) -> None:
        dmg = Path("/tmp/Doc-Media-Toolkit.dmg")
        with patch("scripts.build_standalone.subprocess.run") as run:
            notarize_macos_dmg(dmg, "")
            self.assertEqual(run.call_count, 0)
            notarize_macos_dmg(dmg, "release-profile")
        self.assertEqual(run.call_count, 2)
        self.assertIn("release-profile", run.call_args_list[0].args[0])
        self.assertIn("staple", run.call_args_list[1].args[0])

    def test_experimental_build_uses_branch_and_commit_directory(self) -> None:
        completed = [
            Mock(stdout="agent/smart-target-compression-core\n"),
            Mock(stdout="abc1234\n"),
        ]
        with patch("scripts.build_standalone.subprocess.run", side_effect=completed):
            path = experimental_dist_root(Path("/project"))
        self.assertEqual(
            path,
            Path(
                "/project/dist/experimental/agent_smart-target-compression-core/abc1234"
            ),
        )


if __name__ == "__main__":
    unittest.main()
