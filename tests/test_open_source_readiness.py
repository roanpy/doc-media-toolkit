from __future__ import annotations

import os
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx_output_watermark.gui import detect_language as detect_watermark_language
from pptx_tools import __version__
from pptx_tools.gui import detect_language as detect_shell_language
from pptx_tools.gui import help_topics
from pptx_video_compactor_gui import detect_language as detect_compactor_language


ROOT = Path(__file__).resolve().parents[1]


class OpenSourceReadinessTest(unittest.TestCase):
    def test_public_entry_points_and_policy_files_are_present(self) -> None:
        required = (
            "LICENSE",
            "README.md",
            "README.zh-CN.md",
            "CONTRIBUTING.md",
            "CODE_OF_CONDUCT.md",
            "SECURITY.md",
            "THIRD_PARTY_NOTICES.md",
            "docs/LICENSING.md",
            "docs/DEPENDENCIES.md",
            "docs/INSTALL.md",
            "docs/INSTALL.zh-CN.md",
            "docs/releases/v0.2.1.md",
            "docs/releases/v0.2.0.md",
            "docs/releases/v0.2.0-candidate-audit.md",
            "scripts/build_ffmpeg_runtime.sh",
            "licenses/GPL-3.0-only.txt",
            "licenses/LGPL-3.0-only.txt",
            ".github/pull_request_template.md",
            ".github/ISSUE_TEMPLATE/bug_report.yml",
        )
        self.assertFalse([path for path in required if not (ROOT / path).is_file()])

        metadata = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
        self.assertIn('license = "MIT"', metadata)
        self.assertNotIn("LicenseRef-Proprietary", metadata)
        self.assertIn("github.com/roanpy/doc-media-toolkit", metadata)
        self.assertNotIn("github.com/roanpy/pptx-tools", metadata)
        self.assertIn('{ file = "src/pptx_tools/__init__.py" }', metadata)

    def test_help_discloses_language_and_open_source_scope(self) -> None:
        self.assertIn("开源、语言与隐私", help_topics("zh"))
        self.assertIn("Open Source, Language, and Privacy", help_topics("en"))

    def test_first_launch_defaults_to_english_and_keeps_chinese_override(self) -> None:
        detectors = (
            (detect_shell_language, "PPTX_TOOLS_LANG"),
            (detect_watermark_language, "PPTX_OUTPUT_WATERMARK_LANG"),
            (detect_compactor_language, "PPTX_VIDEO_COMPACTOR_LANG"),
        )
        for detector, variable in detectors:
            with self.subTest(detector=detector.__module__):
                with patch.dict(os.environ, {}, clear=True):
                    self.assertEqual(detector(), "en")
                with patch.dict(os.environ, {variable: "zh-CN"}, clear=True):
                    self.assertEqual(detector(), "zh")

    def test_public_readmes_show_current_version(self) -> None:
        self.assertEqual(__version__, "0.2.1")
        for name in ("README.md", "README.zh-CN.md"):
            self.assertIn(__version__, (ROOT / name).read_text(encoding="utf-8"))

    def test_release_workflow_builds_candidates_without_publish_permission(
        self,
    ) -> None:
        workflow = (ROOT / ".github/workflows/release.yml").read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", workflow)
        self.assertIn("contents: read", workflow)
        self.assertNotIn("contents: write", workflow)
        self.assertNotIn("release_tag", workflow)
        self.assertNotIn("gh release", workflow)
        self.assertNotIn("--windows-onefile", workflow)
        self.assertIn("scripts/build_ffmpeg_runtime.sh", workflow)
        self.assertNotIn("GyanD/codexffmpeg", workflow)

    def test_bundled_ffmpeg_build_is_pinned_and_ships_corresponding_source(
        self,
    ) -> None:
        script = (ROOT / "scripts/build_ffmpeg_runtime.sh").read_text(encoding="utf-8")
        self.assertIn('FFMPEG_VERSION="8.1.2"', script)
        self.assertIn('FFMPEG_SHA256="464beb5e', script)
        self.assertIn('X264_COMMIT="b35605ace3dd', script)
        self.assertIn('ZLIB_VERSION="1.3.2"', script)
        self.assertIn("--enable-libx264", script)
        self.assertIn("--enable-videotoolbox", script)
        self.assertIn("--enable-mediafoundation", script)
        self.assertIn("corresponding-source", script)


if __name__ == "__main__":
    unittest.main()
