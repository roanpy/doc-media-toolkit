from __future__ import annotations

import os
import re
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx_output_watermark.gui import detect_language as detect_watermark_language
from pptx_tools import __version__
from pptx_tools.gui import detect_language as detect_shell_language
from pptx_tools.gui import help_topics
from pptx_video_compactor_gui import detect_language as detect_compactor_language
from scripts.check_public_safety import (
    check_file_content,
    check_sensitive_path,
    check_symbolic_link,
    private_denylist_patterns,
    public_files,
)


ROOT = Path(__file__).resolve().parents[1]


class OpenSourceReadinessTest(unittest.TestCase):
    def test_macos_preferred_language_parser(self) -> None:
        completed = subprocess.CompletedProcess(
            args=[],
            returncode=0,
            stdout='(\n    "zh-Hans-US",\n    "en-US"\n)\n',
            stderr="",
        )
        with (
            patch("pptx_tools.language.sys.platform", "darwin"),
            patch("pptx_tools.language.subprocess.run", return_value=completed),
        ):
            from pptx_tools.language import _macos_preferred_language

            self.assertEqual(_macos_preferred_language(), "zh-Hans-US")

    def test_public_entry_points_and_policy_files_are_present(self) -> None:
        required = (
            "LICENSE",
            "README.md",
            "README.zh-CN.md",
            "CONTRIBUTING.md",
            "CODE_OF_CONDUCT.md",
            "SECURITY.md",
            "THIRD_PARTY_NOTICES.md",
            ".gitattributes",
            "MANIFEST.in",
            "docs/LICENSING.md",
            "docs/DEPENDENCIES.md",
            "docs/INSTALL.md",
            "docs/INSTALL.zh-CN.md",
            "docs/releases/v0.2.4.md",
            "docs/releases/v0.2.3.md",
            "docs/releases/v0.2.2.md",
            "docs/releases/v0.2.1.md",
            "docs/releases/v0.2.0.md",
            "docs/releases/v0.2.0-candidate-audit.md",
            "scripts/build_ffmpeg_runtime.sh",
            "scripts/generate_native_inventory.py",
            "scripts/scan_release_artifact.py",
            "scripts/release_audit.py",
            "licenses/GPL-3.0-only.txt",
            "licenses/LGPL-3.0-only.txt",
            ".github/pull_request_template.md",
            ".github/ISSUE_TEMPLATE/bug_report.yml",
        )
        self.assertFalse([path for path in required if not (ROOT / path).is_file()])

        attributes = (ROOT / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("* text=auto", attributes)
        self.assertIn("*.sh text eol=lf", attributes)

        metadata = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
        self.assertIn('license = "MIT"', metadata)
        self.assertIn('name = "doc-media-toolkit"', metadata)
        self.assertNotIn('name = "pptx-tools"', metadata)
        self.assertIn('pptx-tools = "pptx_tools.cli:main"', metadata)
        self.assertIn('pptx-tools-gui = "pptx_tools.gui:main"', metadata)
        self.assertNotIn("LicenseRef-Proprietary", metadata)
        self.assertIn("github.com/roanpy/doc-media-toolkit", metadata)
        self.assertNotIn("github.com/roanpy/pptx-tools", metadata)
        self.assertIn('{ file = "src/pptx_tools/__init__.py" }', metadata)
        self.assertEqual(metadata.count('"setuptools>=77.0.3"'), 2)

    def test_help_discloses_language_and_open_source_scope(self) -> None:
        self.assertIn("开源、语言与隐私", help_topics("zh"))
        self.assertIn("Open Source, Language, and Privacy", help_topics("en"))

    def test_language_uses_system_locale_and_keeps_explicit_overrides(self) -> None:
        detectors = (
            (detect_shell_language, "PPTX_TOOLS_LANG"),
            (detect_watermark_language, "PPTX_OUTPUT_WATERMARK_LANG"),
            (detect_compactor_language, "PPTX_VIDEO_COMPACTOR_LANG"),
        )
        with (
            patch("pptx_tools.language._macos_preferred_language", return_value=""),
            patch("pptx_tools.language.QLocale.system") as system_locale,
        ):
            system_locale.return_value.name.return_value = "zh_CN"
            for detector, variable in detectors:
                with self.subTest(detector=detector.__module__, locale="zh"):
                    with patch.dict(os.environ, {}, clear=True):
                        self.assertEqual(detector(), "zh")
            system_locale.return_value.name.return_value = "en_US"
            for detector, variable in detectors:
                with self.subTest(detector=detector.__module__, locale="en"):
                    with patch.dict(os.environ, {}, clear=True):
                        self.assertEqual(detector(), "en")
                    with patch.dict(os.environ, {variable: "zh-CN"}, clear=True):
                        self.assertEqual(detector(), "zh")
                    with patch.dict(os.environ, {variable: "en-US"}, clear=True):
                        self.assertEqual(detector(), "en")

    def test_macos_native_language_wins_over_posix_qt_locale(self) -> None:
        detectors = (
            (detect_shell_language, "PPTX_TOOLS_LANG"),
            (detect_watermark_language, "PPTX_OUTPUT_WATERMARK_LANG"),
            (detect_compactor_language, "PPTX_VIDEO_COMPACTOR_LANG"),
        )
        with (
            patch(
                "pptx_tools.language._macos_preferred_language",
                return_value="zh-Hans-US",
            ),
            patch("pptx_tools.language.QLocale.system") as system_locale,
        ):
            system_locale.return_value.name.return_value = "C"
            for detector, variable in detectors:
                with self.subTest(detector=detector.__module__):
                    with patch.dict(
                        os.environ,
                        {"LANG": "C.UTF-8", "LC_ALL": "C.UTF-8"},
                        clear=True,
                    ):
                        self.assertEqual(detector(), "zh")
                    with patch.dict(
                        os.environ,
                        {
                            "LANG": "C.UTF-8",
                            "LC_ALL": "C.UTF-8",
                            variable: "en-US",
                        },
                        clear=True,
                    ):
                        self.assertEqual(detector(), "en")

    def test_public_readmes_show_current_version(self) -> None:
        self.assertEqual(__version__, "0.2.4")
        for name in ("README.md", "README.zh-CN.md"):
            self.assertIn(__version__, (ROOT / name).read_text(encoding="utf-8"))

    def test_documented_setup_commands_exist(self) -> None:
        documented = (
            ROOT / "README.md",
            ROOT / "README.zh-CN.md",
            ROOT / "CONTRIBUTING.md",
            ROOT / "docs" / "INSTALL.md",
            ROOT / "docs" / "INSTALL.zh-CN.md",
            ROOT / "docs" / "USER_GUIDE.zh-CN.md",
            ROOT / "AGENTS.md",
        )
        for path in documented:
            with self.subTest(path=path.name):
                self.assertNotIn("setup_env.sh", path.read_text(encoding="utf-8"))
                self.assertNotIn(
                    'python -m pip install -e ".[dev,build]"',
                    path.read_text(encoding="utf-8"),
                )

    def test_private_public_safety_denylist_is_literal_and_not_echoed(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / ".public-safety-denylist.local").write_text(
                "# private phrases\nProject [Alpha]\n", encoding="utf-8"
            )
            sample = root / "sample.txt"
            sample.write_text("Internal project [alpha]", encoding="utf-8")

            findings = check_file_content(root, sample, private_denylist_patterns(root))

            self.assertEqual(findings, ["sample.txt: matched sensitive/local pattern"])
            self.assertNotIn("Project", findings[0])

    def test_public_safety_checks_extensionless_text_names_and_links(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            extensionless = root / "credentials"
            extensionless.write_text(
                "/".join(("", "Users", "example", "private")), encoding="utf-8"
            )
            sensitive = root / ".env.production"
            sensitive.write_text("placeholder", encoding="utf-8")
            local_agent_config = root / ".codex" / "settings.json"
            invalid_text = root / "invalid.txt"
            invalid_text.write_bytes(b"\xff")
            linked = root / "linked"

            self.assertEqual(
                check_file_content(root, extensionless),
                ["credentials: matched sensitive/local pattern"],
            )
            self.assertEqual(
                check_sensitive_path(root, sensitive),
                [".env.production: sensitive file name or directory"],
            )
            self.assertEqual(
                check_sensitive_path(root, local_agent_config),
                [".codex/settings.json: sensitive file name or directory"],
            )
            self.assertEqual(
                check_file_content(root, invalid_text),
                ["invalid.txt: declared text file is not valid UTF-8"],
            )
            with patch.object(Path, "is_symlink", return_value=True):
                self.assertEqual(
                    check_symbolic_link(root, linked),
                    ["linked: symbolic links are not allowed"],
                )

    def test_public_safety_parses_nul_delimited_git_paths(self) -> None:
        completed = subprocess.CompletedProcess(
            args=[], returncode=0, stdout=b"normal.txt\0line\nbreak.txt\0", stderr=b""
        )
        with (
            tempfile.TemporaryDirectory() as temp_dir,
            patch(
                "scripts.check_public_safety.subprocess.run", return_value=completed
            ) as run,
        ):
            root = Path(temp_dir)
            paths = public_files(root)

        self.assertEqual(
            [path.relative_to(root).as_posix() for path in paths],
            ["normal.txt", "line\nbreak.txt"],
        )
        self.assertIn("-z", run.call_args.args[0])

    def test_public_safety_fallback_keeps_broken_links(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            linked = root / "broken-link"
            linked.symlink_to(root / "missing")
            with patch(
                "scripts.check_public_safety.subprocess.run",
                side_effect=FileNotFoundError,
            ):
                paths = public_files(root)

        self.assertEqual(paths, [linked])

    def test_source_distribution_includes_document_fixtures(self) -> None:
        manifest = (ROOT / "MANIFEST.in").read_text(encoding="utf-8")
        self.assertIn("tests/fixtures/*.docx", manifest)
        self.assertIn("tests/fixtures/*.xlsx", manifest)

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
        self.assertIn("uv export --locked --all-extras", workflow)
        self.assertIn("pip install uv==0.12.3", workflow)
        self.assertIn("pip install pip-audit==2.9.0", workflow)
        self.assertNotIn("pip install --upgrade pip", workflow)
        self.assertIn("pip install --require-hashes", workflow)
        self.assertIn("- macos", workflow)
        self.assertIn("scripts/build_ffmpeg_runtime.sh", workflow)
        self.assertNotIn("GyanD/codexffmpeg", workflow)

        ci = (ROOT / ".github/workflows/ci.yml").read_text(encoding="utf-8")
        self.assertIn("uv export --locked --extra dev", ci)
        self.assertIn("pip install uv==0.12.3", ci)
        self.assertIn("pip install --require-hashes", ci)

    def test_github_actions_are_pinned_to_full_commit_shas(self) -> None:
        for name in ("ci.yml", "release.yml"):
            workflow = (ROOT / ".github" / "workflows" / name).read_text(
                encoding="utf-8"
            )
            refs = re.findall(r"\buses:\s+[^@\s]+@([^\s]+)", workflow)
            self.assertTrue(refs, name)
            self.assertTrue(all(re.fullmatch(r"[0-9a-f]{40}", ref) for ref in refs))

    def test_bundled_ffmpeg_build_is_pinned_and_ships_corresponding_source(
        self,
    ) -> None:
        script = (ROOT / "scripts/build_ffmpeg_runtime.sh").read_text(encoding="utf-8")
        self.assertIn('FFMPEG_VERSION="8.1.2"', script)
        self.assertIn('FFMPEG_SHA256="464beb5e', script)
        self.assertIn('X264_COMMIT="b35605ace3dd', script)
        self.assertIn('ZLIB_VERSION="1.3.2"', script)
        self.assertIn("command -v pkg-config || command -v pkgconf", script)
        self.assertIn('--pkg-config="$pkg_config"', script)
        self.assertIn("download_verified_any", script)
        self.assertIn("github.com/madler/zlib/releases/download", script)
        self.assertIn("--enable-libx264", script)
        self.assertIn("--enable-videotoolbox", script)
        self.assertIn("--enable-mediafoundation", script)
        self.assertIn("--enable-d3d11va", script)
        self.assertIn("corresponding-source", script)
        ffprobe_checksums = [
            line
            for line in script.splitlines()
            if "sha256_file" in line and "bin/ffprobe$binary_suffix" in line
        ]
        self.assertEqual(len(ffprobe_checksums), 1)
        self.assertIn("SHA256SUMS-FFMPEG.txt", script)

    def test_release_evidence_tools_are_fail_closed_and_documented(self) -> None:
        inventory = (ROOT / "scripts/generate_native_inventory.py").read_text(
            encoding="utf-8"
        )
        malware = (ROOT / "scripts/scan_release_artifact.py").read_text(
            encoding="utf-8"
        )
        release = (ROOT / "docs/RELEASE.md").read_text(encoding="utf-8")
        self.assertIn('"doc-media-toolkit.native-inventory.v1"', inventory)
        self.assertIn('"doc-media-toolkit.malware-scan.v1"', malware)
        self.assertIn("No supported malware scanner found", malware)
        self.assertIn("generate_native_inventory.py", release)
        self.assertIn("scan_release_artifact.py", release)


if __name__ == "__main__":
    unittest.main()
