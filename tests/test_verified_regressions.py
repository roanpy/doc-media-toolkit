from __future__ import annotations

import os
import plistlib
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch
from zipfile import ZipFile

from PySide6.QtCore import QMimeData, QPoint, QPointF, QSettings, Qt, QUrl
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QPixmap
from PySide6.QtMultimedia import QMediaPlayer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import (
    QApplication,
    QAbstractButton,
    QAbstractScrollArea,
    QAbstractSlider,
    QAbstractSpinBox,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStyle,
    QTreeWidgetItem,
    QWidget,
)
from PIL import Image

from pptx_output_watermark.gui import (
    MainWindow,
    PreviewArtifacts,
    PreviewSource,
    StyledDialog as WatermarkStyledDialog,
)
from pptx_output_watermark.models import WatermarkOptions
from pptx_output_watermark.process_utils import run_process, terminate_active_processes
from pptx_output_watermark.pptx_video_support import (
    VideoTranscodeProfile,
    _ffmpeg_progress_seconds,
    _resolve_zip_target,
    watermark_video_file,
)
from pptx_video_compactor import (
    CancelledError,
    ImageAsset,
    assign_image_plan,
    build_output_pptx,
    compact_input_path,
    resolve_zip_target,
    run as compactor_run,
)
from pptx_video_compactor_gui import (
    DOCUMENT_INPUT_EXTENSIONS,
    SUPPORTED_COMPACTOR_INPUT_EXTENSIONS,
    CompressionWorker,
    MainWindow as CompressionMainWindow,
    STRINGS as COMPRESSION_STRINGS,
    StyledDialog as CompressionStyledDialog,
    build_namespace,
)
from pptx_tools.gui import (
    AIConnectionWorker,
    HELP_SECTIONS,
    HelpDialog,
    MainWindow as ToolboxMainWindow,
    SettingsDialog,
    ToolSwitch,
    persistent_library_setting,
)
from pptx_tools.ai_client import AIConfig, OpenAICompatibleClient
from pptx_tools.image_manager import ImageProject
from pptx_tools.image_manager_gui import (
    ImportPreviewDialog,
    MainWindow as ImageLibraryMainWindow,
    SimilarImageReviewDialog,
)
from pptx_tools.video_manager import VideoProject
from pptx_tools.manager_i18n import current_language, set_language
from pptx_tools.video_manager_gui import (
    CleanupDialog,
    MainWindow as VideoLibraryMainWindow,
    PendingCleanupDialog,
    PptxUpgradeReviewDialog,
    REVIEW_TAGS_ROLE,
    ResponsiveVideoThumbnail,
    VideoMatchDialog,
    _exec_centered_message,
    _format_duration,
    _format_position,
)
from pptx_tools.ui_theme import (
    SHARED_DIALOG_QSS,
    SHARED_MAIN_QSS,
    DelayedTooltipStyle,
    configure_ui_font,
)
from scripts.build_standalone import (
    BUNDLE_IDENTIFIER,
    add_optional_libreoffice_runtime,
    finalize_macos_bundle_metadata,
    normalize_args,
    project_version,
    require_release_python,
    resolve_libreoffice_root,
)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
_SETTINGS_DIR = tempfile.TemporaryDirectory(prefix="pptx-tools-settings-")
QSettings.setDefaultFormat(QSettings.Format.IniFormat)
QSettings.setPath(
    QSettings.Format.IniFormat,
    QSettings.Scope.UserScope,
    _SETTINGS_DIR.name,
)


class VideoWatermarkGuardTest(unittest.TestCase):
    def test_ffmpeg_progress_supports_current_and_legacy_timestamps(self) -> None:
        self.assertEqual(_ffmpeg_progress_seconds("out_time", "00:01:30.5"), 90.5)
        self.assertEqual(_ffmpeg_progress_seconds("out_time_us", "1500000"), 1.5)
        self.assertEqual(_ffmpeg_progress_seconds("out_time_ms", "1500000"), 1.5)

    def test_missing_dimensions_abort_before_creating_overlay(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "output.mp4"
            with (
                patch(
                    "pptx_output_watermark.pptx_video_support._probe_transcode_profile",
                    return_value=VideoTranscodeProfile(),
                ),
                patch(
                    "pptx_output_watermark.pptx_video_support.write_watermark_overlay_image"
                ) as write_overlay,
                patch(
                    "pptx_output_watermark.pptx_video_support._encode_watermarked_video"
                ) as encode_video,
            ):
                with self.assertRaisesRegex(RuntimeError, "video dimensions"):
                    watermark_video_file(
                        Path("input.mp4"), output_path, WatermarkOptions(enabled=True)
                    )

            write_overlay.assert_not_called()
            encode_video.assert_not_called()


class PackageTargetResolutionTest(unittest.TestCase):
    def test_absolute_targets_are_normalized_to_zip_paths(self) -> None:
        expected = "ppt/media/video1.mp4"
        source = "ppt/slides/slide1.xml"
        target = "/ppt/media/video1.mp4"

        self.assertEqual(_resolve_zip_target(source, target), expected)
        self.assertEqual(resolve_zip_target(source, target), expected)

    def test_backslash_targets_are_normalized_before_path_resolution(self) -> None:
        self.assertEqual(
            resolve_zip_target("word/document.xml", r"media\..\..\outside.txt"),
            "outside.txt",
        )


class AtomicPptxOutputTest(unittest.TestCase):
    def test_build_output_replaces_media_and_can_replace_input_atomically(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            deck = Path(temp_dir) / "same.pptx"
            replacement = Path(temp_dir) / "replacement.png"
            replacement.write_bytes(b"compressed-image")
            with ZipFile(deck, "w") as archive:
                archive.writestr("[Content_Types].xml", b"<Types />")
                archive.writestr("ppt/presentation.xml", b"presentation")
                archive.writestr("ppt/media/image1.png", b"original-image")

            image = ImageAsset(
                media_path="ppt/media/image1.png",
                zip_size=len(b"original-image"),
                output_path=str(replacement),
            )
            build_output_pptx(deck, deck, {}, {image.media_path: image})

            with ZipFile(deck) as archive:
                self.assertEqual(archive.read("ppt/presentation.xml"), b"presentation")
                self.assertEqual(archive.read(image.media_path), b"compressed-image")
            self.assertEqual(
                build_namespace(deck, None, "high", "none", deck).output, deck
            )


class TransparentPngTest(unittest.TestCase):
    def test_compression_preserves_alpha_channel(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "transparent.png"
            image = Image.new("RGBA", (64, 64), (20, 80, 160, 255))
            for x in range(32):
                for y in range(64):
                    image.putpixel((x, y), (20, 80, 160, 0))
            image.save(source, compress_level=0)

            result = compact_input_path(
                build_namespace(source, None, "none", "balanced"),
                logger=lambda message: None,
            )

            with Image.open(result["output_pptx"]) as compressed:
                self.assertEqual(compressed.format, "PNG")
                self.assertEqual(
                    compressed.convert("RGBA").getchannel("A").getextrema(), (0, 255)
                )

    def test_lossless_profile_only_recompresses_png(self) -> None:
        assets = {
            "ppt/media/image1.png": ImageAsset(
                media_path="ppt/media/image1.png", zip_size=1000
            ),
            "ppt/media/image2.jpeg": ImageAsset(
                media_path="ppt/media/image2.jpeg", zip_size=1000
            ),
        }

        assign_image_plan(assets, "lossless")

        self.assertEqual(assets["ppt/media/image1.png"].status, "planned")
        self.assertEqual(assets["ppt/media/image1.png"].quality, 100)
        self.assertEqual(assets["ppt/media/image2.jpeg"].status, "copy_requested")
        self.assertEqual(assets["ppt/media/image2.jpeg"].target_bytes, 1000)


class PreviewSourceCleanupTest(unittest.TestCase):
    def test_active_worker_source_is_not_deleted(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "source.pptx"
            input_path.touch()
            source_root = root / "preview-source"
            source_root.mkdir()
            key = (str(input_path.resolve()), 1, 1)
            source = PreviewSource(key, source_root, [], 1)
            window = SimpleNamespace(
                preview_source_cache={key: source},
                preview_source_cleanup_pending=set(),
                preview_thread=None,
                preview_worker=SimpleNamespace(source=source),
                preview_dirty=False,
            )

            MainWindow.drop_preview_source(window, input_path)

            self.assertNotIn(key, window.preview_source_cache)
            self.assertTrue(source_root.exists())

            MainWindow.finish_preview_thread(window)

            self.assertFalse(source_root.exists())


class DmgArgumentValidationTest(unittest.TestCase):
    def test_release_build_requires_python_312(self) -> None:
        require_release_python((3, 12))
        with self.assertRaisesRegex(SystemExit, "require Python 3.12"):
            require_release_python((3, 14))

    def test_invalid_dmg_modes_fail_during_argument_normalization(self) -> None:
        project_root = Path("/project")
        cases = (
            ("windows", False, False, "only supported on macOS"),
            ("macos", True, False, "GUI onedir build"),
            ("macos", False, True, "GUI onedir build"),
        )
        for platform, cli, onefile, message in cases:
            with self.subTest(platform=platform, cli=cli, onefile=onefile):
                args = SimpleNamespace(
                    windows_onefile=False,
                    gui=not cli,
                    cli=cli,
                    onefile=onefile,
                    target_platform="auto",
                    icon=None,
                    name=None,
                    dmg=True,
                )
                with (
                    patch(
                        "scripts.build_standalone.host_platform", return_value=platform
                    ),
                    self.assertRaisesRegex(SystemExit, message),
                ):
                    normalize_args(args, project_root)

    def test_macos_bundle_metadata_uses_project_version_and_valid_identifier(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            package = root / "src" / "pptx_tools"
            package.mkdir(parents=True)
            (package / "__init__.py").write_text(
                '__version__ = "1.2.3"\n', encoding="utf-8"
            )
            app_bundle = root / "Doc Media Toolkit.app"
            contents = app_bundle / "Contents"
            contents.mkdir(parents=True)
            plist_path = contents / "Info.plist"
            with plist_path.open("wb") as plist_file:
                plistlib.dump(
                    {
                        "CFBundleIdentifier": "Doc Media Toolkit",
                        "CFBundleShortVersionString": "0.0.0",
                    },
                    plist_file,
                )

            with patch("scripts.build_standalone.subprocess.run") as run:
                finalize_macos_bundle_metadata(
                    app_bundle,
                    version=project_version(root),
                )

            with plist_path.open("rb") as plist_file:
                metadata = plistlib.load(plist_file)
            self.assertEqual(metadata["CFBundleIdentifier"], BUNDLE_IDENTIFIER)
            self.assertEqual(metadata["CFBundleShortVersionString"], "1.2.3")
            self.assertEqual(metadata["CFBundleVersion"], "1.2.3")
            self.assertEqual(metadata["LSMinimumSystemVersion"], "13.0")
            self.assertEqual(run.call_count, 2)
            self.assertIn("--deep", run.call_args_list[0].args[0])
            self.assertIn("--verify", run.call_args_list[1].args[0])

    def test_complete_libreoffice_runtime_can_be_added_as_build_data(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            app = Path(temp_dir) / "LibreOffice.app"
            (app / "Contents" / "MacOS").mkdir(parents=True)
            (app / "Contents" / "Resources").mkdir(parents=True)
            (app / "Contents" / "MacOS" / "soffice").touch()
            (app / "Contents" / "Resources" / "LICENSE").write_text(
                "license", encoding="utf-8"
            )
            args: list[str] = []
            with patch("scripts.build_standalone.sys.platform", "darwin"):
                self.assertEqual(resolve_libreoffice_root(app), app.resolve())
                bundled = add_optional_libreoffice_runtime(
                    args,
                    ":",
                    explicit_root=app,
                    required=True,
                )

            self.assertEqual(bundled, app.resolve())
            self.assertEqual(
                args,
                [
                    "--add-data",
                    f"{app.resolve()}:libreoffice/LibreOffice.app",
                ],
            )

    def test_bundled_libreoffice_rejects_onefile_build(self) -> None:
        args = SimpleNamespace(
            windows_onefile=False,
            gui=True,
            cli=False,
            onefile=True,
            target_platform="auto",
            icon=None,
            name=None,
            dmg=False,
            bundle_libreoffice=True,
            require_libreoffice_bundle=False,
        )
        with self.assertRaisesRegex(SystemExit, "requires an onedir build"):
            normalize_args(args, Path("/project"))


class AutomaticAuditWorkflowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_checked_option_audits_then_optimizes_failed_pptx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            output = root / "source_compact.pptx"
            source.touch()
            output.touch()
            window = CompressionMainWindow()
            self.assertEqual(window.video_threshold_spinbox.width(), 80)
            self.assertEqual(window.image_threshold_spinbox.width(), 80)
            self.assertEqual(
                [window.archive_mode_select.itemData(index) for index in range(4)],
                ["off", "1080p", "mp4", "original"],
            )
            window.input_paths = [source]
            window.file_statuses = {source: "running"}
            window.auto_optimize_checkbox.setChecked(True)
            window.refresh_file_list()
            first_row = window.file_list.itemWidget(window.file_list.item(0))
            type_icon = first_row.findChild(QLabel, "fileTypeIcon")
            self.assertIsNotNone(type_icon)
            self.assertFalse(type_icon.pixmap().isNull())

            with patch.object(window, "run_audit") as run_audit:
                window.on_finished([(source, output, 1, False, None)], [])
            run_audit.assert_called_once_with({source})

            old_source = root / "old.pptx"
            old_source.touch()
            window.input_paths.append(old_source)
            window.file_statuses[old_source] = "done"
            window.refresh_file_list()
            with patch.object(window, "_audit_next"):
                window.run_audit({source})
            self.assertEqual(window.audit_queue, [source])

            window.audit_queue = []
            window.failed_audits = {source: [object()]}
            with patch.object(window, "on_optimize_clicked") as optimize:
                window._audit_next()
            optimize.assert_called_once_with()
            window.close()

    def test_video_library_archive_is_opt_in_by_default(self) -> None:
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        keys = (
            "compression/archive_videos",
            "compression/archive_source_quality",
            "compression/overwrite_original",
            "video_manager/last_project",
        )
        previous = {key: settings.value(key) for key in keys}
        for key in keys:
            settings.remove(key)
        window = None
        try:
            with patch.dict(
                os.environ, {"PPTX_VIDEO_COMPACTOR_LANG": "en"}, clear=False
            ):
                window = CompressionMainWindow()
            self.assertEqual(window.archive_mode_select.currentData(), "off")
            self.assertEqual(window.image_archive_mode_select.currentData(), "off")
            self.assertIn("optional", window.archive_mode_label.text().lower())
            self.assertFalse(window.image_library_button.isEnabled())
            self.assertTrue(window.archive_settings_panel.isHidden())
            self.assertIn("deduplicated", window.archive_summary_label.text().lower())
            self.assertIn("border-radius: 10px", SHARED_MAIN_QSS)
            self.assertIn("font-size: 12px", SHARED_DIALOG_QSS)
            self.assertFalse(window.overwrite_checkbox.isEnabled())
            self.assertFalse(window.archive_library_button.isEnabled())
            window.archive_mode_select.setCurrentIndex(1)
            self.assertTrue(window.overwrite_checkbox.isEnabled())
            self.assertTrue(window.archive_library_button.isEnabled())
            self.assertEqual(
                window.archive_library_button.text(),
                window.text["archive_library_unset"],
            )
        finally:
            for key, value in previous.items():
                if value is None:
                    settings.remove(key)
                else:
                    settings.setValue(key, value)
            if window is not None:
                window.close()

    def test_compression_shows_and_can_change_archive_library(self) -> None:
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        previous = settings.value("video_manager/last_project")
        window = None
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                first = VideoProject.create(Path(temp_dir) / "first", "第一视频库")
                second_root = Path(temp_dir) / "second"
                settings.setValue("video_manager/last_project", str(first.root))
                window = CompressionMainWindow()
                window.archive_mode_select.setCurrentIndex(1)
                self.assertIn("first", window.archive_library_button.text())
                self.assertIn(str(first.root), window.archive_library_button.toolTip())

                with patch.object(
                    QFileDialog,
                    "getExistingDirectory",
                    return_value=str(second_root),
                ):
                    selected = window.choose_archive_library()
                second_root = second_root.resolve()
                self.assertEqual(selected, second_root)
                self.assertTrue((second_root / "video-project.json").is_file())
                self.assertEqual(
                    settings.value("video_manager/last_project"), str(second_root)
                )
                self.assertIn("second", window.archive_library_button.text())

                settings.setValue("video_manager/last_project", str(first.root))
                window.on_activated()
                self.assertIn("first", window.archive_library_button.text())
        finally:
            if previous is None:
                settings.remove("video_manager/last_project")
            else:
                settings.setValue("video_manager/last_project", previous)
            if window is not None:
                window.close()

    def test_video_and_image_audit_thresholds_are_independent_and_persisted(
        self,
    ) -> None:
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        keys = (
            "compression/video_ssim_threshold",
            "compression/video_ssim_threshold_manual",
            "compression/image_ssim_threshold",
        )
        previous = {key: settings.value(key) for key in keys}
        for key in keys:
            settings.remove(key)
        first = None
        second = None
        try:
            first = CompressionMainWindow()
            self.assertEqual(first.video_threshold_spinbox.value(), 0.95)
            self.assertEqual(first.image_threshold_spinbox.value(), 0.99)
            self.assertEqual(first.image_profile_select.currentData(), "lossless")
            first.profile_select.setCurrentIndex(3)
            self.assertEqual(first.video_threshold_spinbox.value(), 0.90)
            first.video_threshold_spinbox.setValue(0.91)
            first.profile_select.setCurrentIndex(2)
            self.assertEqual(first.video_threshold_spinbox.value(), 0.91)
            first.image_threshold_spinbox.setValue(0.98)
            first.close()
            first = None

            second = CompressionMainWindow()
            self.assertEqual(second.video_threshold_spinbox.value(), 0.91)
            self.assertEqual(second.image_threshold_spinbox.value(), 0.98)
        finally:
            for key, value in previous.items():
                if value is None:
                    settings.remove(key)
                else:
                    settings.setValue(key, value)
            if first is not None:
                first.close()
            if second is not None:
                second.close()

    def test_compression_progress_does_not_regress_on_retry(self) -> None:
        worker = CompressionWorker(
            [], None, "high", "lossless", "zh", COMPRESSION_STRINGS["zh"]
        )
        worker.total_bytes = 100
        values: list[int] = []
        worker.progress.connect(lambda value, _label: values.append(value))

        worker._progress(100, 0.9, 1.0, "first pass")
        worker._progress(100, 0.2, 1.0, "retry")

        self.assertEqual(values, [90, 90])


class DesktopLifecycleTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.addCleanup(set_language, current_language())
        set_language("zh")
        self.settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        self.last_project = self.settings.value("video_manager/last_project")
        self.library_sort = self.settings.value("video_library/sort")
        self.library_sort_descending = self.settings.value(
            "video_library/sort_descending"
        )
        self.settings.remove("video_manager/last_project")
        self.settings.remove("video_library/sort")
        self.settings.remove("video_library/sort_descending")

    def test_temporary_library_is_not_restored_as_default(self) -> None:
        key = "image_manager/last_project"
        previous = self.settings.value(key)
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir) / "library"
                root.mkdir()
                (root / "image-project.json").write_text("{}", encoding="utf-8")
                self.settings.setValue(key, str(root))

                self.assertEqual(
                    persistent_library_setting(self.settings, key),
                    "",
                )
                self.assertIsNone(self.settings.value(key))
        finally:
            if previous is None:
                self.settings.remove(key)
            else:
                self.settings.setValue(key, previous)

    def test_restored_library_offers_batch_status_refresh(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = VideoProject.create(Path(temp_dir) / "library")
            window = VideoLibraryMainWindow()
            with (
                patch.dict(os.environ, {"QT_QPA_PLATFORM": "cocoa"}),
                patch.object(window, "_offer_refresh_modified_variants") as offer,
            ):
                window.open_project(project.root, report_errors=False)
            offer.assert_called_once_with()
            window.close()

    def test_unavailable_persistent_library_path_is_not_forgotten(self) -> None:
        key = "video_manager/last_project"
        root = Path.home() / f".pptx-tools-unmounted-library-{os.getpid()}"
        self.settings.setValue(key, str(root))

        self.assertEqual(
            persistent_library_setting(
                self.settings,
                key,
            ),
            str(root.resolve()),
        )
        self.assertEqual(self.settings.value(key, "", str), str(root))

    def test_library_pages_recover_when_saved_storage_becomes_available(
        self,
    ) -> None:
        image_key = "image_manager/last_project"
        previous_image = self.settings.value(image_key)
        self.settings.remove(image_key)
        video_window = VideoLibraryMainWindow()
        image_window = ImageLibraryMainWindow()
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir)
                video = VideoProject.create(root / "video")
                image = ImageProject.create(root / "image")
                self.settings.setValue("video_manager/last_project", str(video.root))
                self.settings.setValue("image_manager/last_project", str(image.root))

                video_window.on_activated()
                image_window.on_activated()

                self.assertEqual(video_window.project.root, video.root)
                self.assertEqual(image_window.project.root, image.root)
        finally:
            video_window.close()
            image_window.close()
            if previous_image is None:
                self.settings.remove(image_key)
            else:
                self.settings.setValue(image_key, previous_image)

    def test_library_ai_actions_explain_missing_configuration(self) -> None:
        keys = (
            "doc_media_ai_base_url",
            "doc_media_ai_model",
        )
        previous = {key: self.app.property(key) for key in keys}
        for key in keys:
            self.app.setProperty(key, "")
        video_window = VideoLibraryMainWindow()
        image_window = ImageLibraryMainWindow()
        try:
            video_window._update_action_states()
            image_window._sync_actions()
            self.assertFalse(video_window.ai_button.isEnabled())
            self.assertFalse(video_window.detail_ai_button.isEnabled())
            self.assertIn("顶栏齿轮", video_window.ai_button.toolTip())
            self.assertIn("顶栏齿轮", video_window.detail_ai_button.toolTip())
            self.assertFalse(image_window.ai_button.isEnabled())
            self.assertIn("顶栏齿轮", image_window.ai_button.toolTip())
            self.assertEqual(image_window.new_project_button.height(), 32)
            self.assertEqual(image_window.open_project_button.height(), 32)
        finally:
            video_window.close()
            image_window.close()
            for key, value in previous.items():
                self.app.setProperty(key, value)

    def test_ai_text_validation_does_not_probe_vision_when_disabled(self) -> None:
        worker = AIConnectionWorker(
            AIConfig(
                "https://example.test/v1",
                "text-model",
                vision_enabled=False,
            )
        )
        results: list[object] = []
        worker.finished.connect(results.append)
        with (
            patch.object(OpenAICompatibleClient, "test_connection", return_value="ok"),
            patch.object(
                OpenAICompatibleClient, "probe_vision_support"
            ) as vision_probe,
        ):
            worker.run()

        vision_probe.assert_not_called()
        self.assertEqual(results[0]["vision"], None)

    def tearDown(self) -> None:
        if self.last_project is not None:
            self.settings.setValue("video_manager/last_project", self.last_project)
        if self.library_sort is None:
            self.settings.remove("video_library/sort")
        else:
            self.settings.setValue("video_library/sort", self.library_sort)
        if self.library_sort_descending is None:
            self.settings.remove("video_library/sort_descending")
        else:
            self.settings.setValue(
                "video_library/sort_descending", self.library_sort_descending
            )

    def test_image_library_stages_previews_and_imports_files(self) -> None:
        previous = self.settings.value("image_manager/last_project")
        self.settings.remove("image_manager/last_project")
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir)
                source = root / "product.png"
                Image.new("RGB", (120, 80), (40, 90, 160)).save(source)
                project = ImageProject.create(root / "library")
                window = ImageLibraryMainWindow()
                window._set_project(project)

                self.assertEqual(
                    window.choose_import_button.height(),
                    window.workflow_menu_button.height(),
                )
                self.assertTrue(window.workflow_settings.isHidden())
                window._toggle_workflow_settings()
                self.assertFalse(window.workflow_settings.isHidden())
                self.assertGreaterEqual(window.workflow_settings.minimumHeight(), 38)
                window._toggle_workflow_settings()
                window.add_pending_paths([source, source])
                self.assertEqual(window.pending_paths, [source.resolve()])
                self.assertTrue(window.preview_import_button.isEnabled())
                self.assertIn("product.png", window.import_drop_label.text())

                preview = ImportPreviewDialog(window, window.pending_paths)
                self.assertEqual(len(preview.records), 1)
                self.assertFalse(preview.records[0]["pixmap"].isNull())
                preview.close()

                window.import_category_input.setText("产品图")
                window.import_pending()
                deadline = time.monotonic() + 3
                while window.worker_thread is not None and time.monotonic() < deadline:
                    QApplication.processEvents()
                    time.sleep(0.01)
                self.assertIsNone(window.worker_thread)
                self.assertFalse(window.pending_paths)
                self.assertEqual(project.assets()[0]["category"], "产品图")
                window.library_filter_input.setText("不存在的分类")
                self.assertTrue(window.tree.isHidden())
                self.assertFalse(window.library_empty.isHidden())
                window.library_filter_input.clear()
                self.assertFalse(window.tree.isHidden())
                self.assertEqual(
                    window.more_actions_button.height(),
                    window.choose_import_button.height(),
                )
                self.assertGreaterEqual(
                    window.library_toolbar_layout.indexOf(window.edit_button), 0
                )
                self.assertGreaterEqual(
                    window.library_toolbar_layout.indexOf(window.similar_button), 0
                )
                self.assertGreaterEqual(
                    window.library_toolbar_layout.indexOf(window.more_actions_button),
                    0,
                )
                self.assertEqual(
                    window.library_toolbar_layout.indexOf(window.ai_button), -1
                )
                self.assertEqual(
                    window.library_toolbar_layout.indexOf(window.open_location_button),
                    -1,
                )
                self.assertTrue(window.ai_button.isHidden())
                self.assertTrue(window.open_location_button.isHidden())
                self.assertTrue(window.pending_cleanup_button.isHidden())
                self.assertEqual(window.similar_button.text(), "查找相似图")
                self.assertIn("不会自动合并", window.similar_button.toolTip())
                self.assertTrue(window.more_actions_menu.toolTipsVisible())
                window._sync_more_actions()
                self.assertTrue(
                    all(
                        action.toolTip()
                        for action, _button in window.more_action_targets
                    )
                )
                self.assertEqual(
                    [
                        action.text()
                        for action in window.more_actions_menu.actions()
                        if action.text()
                    ],
                    [
                        "资源",
                        "AI 整理建议",
                        "打开位置",
                        "移除选中",
                        "库维护",
                        "库体检",
                        "清理未引用文件",
                        "待清理 (0)",
                        "重置已忽略候选",
                    ],
                )
                window.update_responsive_layout(1000)
                self.assertFalse(window.health_filter_combo.isHidden())
                self.assertTrue(
                    all(
                        button.isHidden()
                        for button in window.health_filter_buttons.values()
                    )
                )
                window.update_responsive_layout(1280)
                self.assertTrue(window.health_filter_combo.isHidden())
                self.assertTrue(
                    all(
                        not button.isHidden()
                        for button in window.health_filter_buttons.values()
                    )
                )
                window.close()
        finally:
            if previous is None:
                self.settings.remove("image_manager/last_project")
            else:
                self.settings.setValue("image_manager/last_project", previous)

    def test_image_library_health_filters_and_review_preview(self) -> None:
        previous = self.settings.value("image_manager/last_project")
        previous_width = self.settings.value("image_manager/min_width")
        previous_height = self.settings.value("image_manager/min_height")
        self.settings.remove("image_manager/last_project")
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir)
                first = root / "first.png"
                copy = root / "first-copy.png"
                similar = root / "similar.png"
                small = root / "small.png"
                Image.new("RGB", (120, 80), (100, 100, 100)).save(first)
                copy.write_bytes(first.read_bytes())
                Image.new("RGB", (120, 80), (101, 101, 101)).save(similar)
                Image.new("RGB", (8, 8), (20, 30, 40)).save(small)
                project = ImageProject.create(root / "library")
                project.import_paths([first, copy, similar, small])
                project.assets()[-1]["origins"] = []
                project.save()
                window = ImageLibraryMainWindow()
                window._set_project(project)
                window.min_width_spin.setValue(16)
                window.min_height_spin.setValue(16)

                self.assertEqual(
                    [button.text() for button in window.health_filter_buttons.values()],
                    [
                        "全部 3",
                        "重复来源 1",
                        "相似 2",
                        "过小 1",
                        "无来源 1",
                    ],
                )
                self.assertEqual(window.health_filter_combo.itemText(0), "全部 3")
                for key, expected in (
                    ("duplicate_origins", 1),
                    ("similar", 2),
                    ("undersized", 1),
                    ("no_origin", 1),
                ):
                    window._filter_by_health(key)
                    self.assertEqual(window.tree.topLevelItemCount(), expected)
                    self.assertEqual(window.health_filter_combo.currentData(), key)
                window._filter_by_health("all")
                window.tree.setCurrentItem(window.tree.topLevelItem(0))
                self.assertEqual(window.detail_origins.topLevelItemCount(), 2)
                self.assertEqual(window.pending_cleanup_button.text(), "待清理 0")
                window._sync_more_actions()
                self.assertTrue(window.health_action.isEnabled())
                with patch(
                    "pptx_tools.image_manager_gui.QMessageBox.information"
                ) as information:
                    window.show_library_health()
                self.assertIn("未发现异常", information.call_args.args[2])

                left, right = project.assets()[:2]
                dialog = SimilarImageReviewDialog(window, project, left, right, 100.0)
                previews = dialog.findChildren(QLabel, "imagePreview")
                self.assertEqual(len(previews), 2)
                self.assertTrue(
                    all(preview.pixmap() is not None for preview in previews)
                )
                dialog._choose("ignore")
                self.assertEqual(dialog.decision, "ignore")
                window.close()
        finally:
            for key, value in (
                ("image_manager/last_project", previous),
                ("image_manager/min_width", previous_width),
                ("image_manager/min_height", previous_height),
            ):
                if value is None:
                    self.settings.remove(key)
                else:
                    self.settings.setValue(key, value)

    def test_switching_tabs_preserves_window_and_settings(self) -> None:
        window = ToolboxMainWindow()
        self.assertEqual(window.tabs.count(), 4)
        self.assertIsNotNone(window.embedded_tools[2].window)
        watermark_window = window.embedded_tools[0].window
        self.assertIsNotNone(watermark_window)
        watermark_window.watermark_text_input.setText("keep this setting")

        window.tabs.setCurrentIndex(1)
        window.tabs.setCurrentIndex(0)

        self.assertIs(window.embedded_tools[0].window, watermark_window)
        self.assertEqual(
            watermark_window.watermark_text_input.text(), "keep this setting"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "images.docx"
            source.touch()
            watermark_window.set_files([source])
            window.tabs.setCurrentIndex(3)
            self.assertEqual(
                window.embedded_tools[3].window.pending_paths, [source.resolve()]
            )
        window.close()

    def test_shared_shell_uses_canonical_header_and_text_tabs(self) -> None:
        with patch.dict(os.environ, {"PPTX_TOOLS_LANG": "en"}, clear=False):
            window = ToolboxMainWindow()
        self.assertEqual(window.minimumWidth(), 880)
        self.assertEqual(window.header_card.height(), 58)
        self.assertTrue(window.header_eyebrow.isHidden())
        self.assertEqual(window.switcher.height(), ToolSwitch.HEIGHT)
        self.assertEqual(window.switcher.accessibleName(), "Tool switch")
        self.assertIn(
            window.switcher.labels[0],
            window.switcher.accessibleDescription(),
        )
        QTest.keyClick(window.switcher, Qt.Key.Key_End)
        self.assertEqual(
            window.switcher.current_index(),
            len(window.switcher.labels) - 1,
        )
        self.assertIn(
            window.switcher.labels[-1],
            window.switcher.accessibleDescription(),
        )
        QTest.keyClick(window.switcher, Qt.Key.Key_1)
        self.assertEqual(window.switcher.current_index(), 0)
        self.assertEqual(
            window.help_button.accessibleName(),
            window.text["help_button"],
        )
        self.assertEqual(window.settings_button.size(), window.help_button.size())
        self.assertEqual(window.settings_button.width(), 40)
        self.assertTrue(window.settings_button.text() == "")
        self.assertTrue(window.help_button.text() == "")
        self.assertFalse(window.settings_button.icon().isNull())
        self.assertFalse(window.help_button.icon().isNull())
        self.assertIn("font-size: 18px", window.styleSheet())
        self.assertNotIn("border-bottom", window.styleSheet())

        help_dialog = HelpDialog("zh", 0, window)
        self.assertEqual(help_dialog.minimumWidth(), 760)
        self.assertEqual(help_dialog.minimumHeight(), 540)
        self.assertEqual(help_dialog.topic_tree.topLevelItemCount(), 3)
        self.assertEqual(
            help_dialog.topic_tree.currentItem().text(0),
            "文档及媒体水印导出",
        )
        self.assertIn("文档及媒体水印导出", help_dialog.body.toPlainText())
        self.assertIn("font-size: 22px", help_dialog.styleSheet())
        self.assertIn("font-size: 13px", help_dialog.styleSheet())
        self.assertIn(
            "QTreeWidget#helpTopicTree::item:selected",
            help_dialog.styleSheet(),
        )

        help_dialog.search_input.setText("API Key")
        QApplication.processEvents()
        self.assertEqual(help_dialog.topic_tree.currentItem().text(0), "AI 辅助")
        self.assertIn("API Key", help_dialog.body.toPlainText())
        help_dialog.search_input.clear()
        QApplication.processEvents()
        self.assertFalse(help_dialog.topic_tree.topLevelItem(0).isHidden())
        self.assertFalse(help_dialog.topic_tree.topLevelItem(1).isHidden())
        self.assertFalse(help_dialog.topic_tree.topLevelItem(2).isHidden())
        help_dialog.close()
        window.close()

    def test_main_surfaces_use_canonical_pixel_type_scale(self) -> None:
        windows = [
            ToolboxMainWindow(),
            MainWindow(),
            CompressionMainWindow(),
            VideoLibraryMainWindow(),
        ]
        for index, window in enumerate(windows):
            stylesheet = window.styleSheet()
            self.assertIn("font-size:", stylesheet)
            if index:
                self.assertIn("font-size: 11px", stylesheet)
            if index:
                self.assertIn("QPushButton:focus", stylesheet)
            window.close()

    def test_processed_image_archive_uses_output_without_blocking_compression(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.png"
            output = root / "source-compressed.png"
            Image.new("RGB", (8, 8), (255, 0, 0)).save(source)
            Image.new("RGB", (8, 8), (0, 0, 255)).save(output)
            library = ImageProject.create(root / "images")
            completed = []
            worker = CompressionWorker(
                [source],
                None,
                "high",
                "high",
                "zh",
                COMPRESSION_STRINGS["zh"],
                image_library_root=library.root,
                image_archive_mode="processed",
                image_archive_category="压缩结果",
            )
            worker.finished.connect(
                lambda results, failures, cancelled, stopped: completed.append(
                    (results, failures, cancelled, stopped)
                )
            )

            with patch(
                "pptx_video_compactor_gui.compact_input_path",
                return_value={"output_pptx": output},
            ):
                worker.run()

            reopened = ImageProject.open(library.root)
            self.assertEqual(len(reopened.assets()), 1)
            self.assertEqual(reopened.assets()[0]["category"], "压缩结果")
            self.assertEqual(completed[0][1], [])

            failed_archive = []
            worker = CompressionWorker(
                [source],
                None,
                "high",
                "high",
                "zh",
                COMPRESSION_STRINGS["zh"],
                image_library_root=root / "missing-library",
                image_archive_mode="processed",
            )
            worker.finished.connect(
                lambda results, failures, cancelled, stopped: failed_archive.append(
                    (results, failures)
                )
            )
            with patch(
                "pptx_video_compactor_gui.compact_input_path",
                return_value={"output_pptx": output},
            ):
                worker.run()
            self.assertEqual(len(failed_archive[0][0]), 1)
            self.assertEqual(failed_archive[0][1], [])

    def test_original_images_are_archived_before_overwriting_pptx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            deck = root / "deck.pptx"
            image = root / "image.png"
            Image.new("RGB", (8, 8), (20, 40, 60)).save(image)
            relationship = (
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/image" Target="../media/image1.png"/>'
                "</Relationships>"
            )
            with ZipFile(deck, "w") as archive:
                archive.writestr("ppt/slides/slide1.xml", "<root/>")
                archive.writestr("ppt/slides/_rels/slide1.xml.rels", relationship)
                archive.writestr("ppt/media/image1.png", image.read_bytes())
            library = ImageProject.create(root / "images")
            worker = CompressionWorker(
                [deck],
                None,
                "high",
                "none",
                "zh",
                COMPRESSION_STRINGS["zh"],
                image_library_root=library.root,
                image_archive_mode="original",
                overwrite_original=True,
            )

            def compact_after_archive(*_args, **_kwargs):
                self.assertEqual(len(ImageProject.open(library.root).assets()), 1)
                return {"output_pptx": deck}

            with patch(
                "pptx_video_compactor_gui.compact_input_path",
                side_effect=compact_after_archive,
            ) as compact:
                worker.run()
            compact.assert_called_once()

    def test_watermark_page_uses_three_panes_and_expanded_log(self) -> None:
        with patch("pptx_output_watermark.gui.detect_language", return_value="zh"):
            window = MainWindow()
        self.assertTrue(
            all(
                button.toolTip()
                for button in window.findChildren(QPushButton)
                if button.text().strip()
            )
        )
        self.assertEqual(window.minimumWidth(), 880)
        self.assertEqual(
            window.center_pane.parentWidget(), window.left_pane.parentWidget()
        )
        self.assertEqual(window.left_pane.minimumWidth(), 320)
        self.assertEqual(window.left_pane.maximumWidth(), 320)
        self.assertGreaterEqual(window.center_pane.minimumWidth(), 520)
        self.assertEqual(window.right_pane.minimumWidth(), 410)
        self.assertEqual(window.right_pane.maximumWidth(), 460)
        window.update_responsive_layout(960)
        self.assertEqual(window.left_pane.width(), 260)
        self.assertEqual(window.center_pane.minimumWidth(), 0)
        self.assertEqual(window.right_pane.minimumWidth(), 310)
        self.assertEqual(window.right_pane.maximumWidth(), 330)
        for control in (
            window.output_format_select,
            window.output_mode_select,
            window.dpi_select,
            window.color_select,
        ):
            self.assertGreaterEqual(control.minimumWidth(), control.sizeHint().width())
        self.assertIn("QPushButton:focus", window.styleSheet())
        self.assertTrue(window.advanced_panel.isHidden())
        self.assertFalse(window.event_log.isVisible())
        self.assertTrue(window.log_drawer.isHidden())
        self.assertEqual(window.log_hover_timer.interval(), 1000)
        self.assertEqual(window.color_select.currentText(), "灰蓝")
        self.assertFalse(window.color_select.itemIcon(0).isNull())
        self.assertEqual(
            window.top_controls.itemAtPosition(1, 0).widget(),
            window.output_quality_label_widget.parentWidget(),
        )
        self.assertEqual(
            window.top_controls.itemAtPosition(1, 1).widget(),
            window.keep_videos_group,
        )
        format_group = window.output_format_select.parentWidget()
        self.assertLessEqual(
            window.output_format_select.geometry().right(),
            format_group.contentsRect().right(),
        )
        window.append_log("[INFO] 日志浮层验证")
        window.resize(960, 652)
        window.show()
        QApplication.processEvents()
        window.show_log_drawer(auto_hide=False)
        self.assertFalse(window.log_drawer.isHidden())
        self.assertLessEqual(
            window.log_drawer.height(),
            int(window.content_widget.height() * 0.35) + 1,
        )
        window.hide_log_drawer()
        self.assertTrue(window.log_drawer.isHidden())
        window.is_running = True
        window.show_log_drawer(auto_hide=False)
        window.hide_log_drawer_if_idle()
        self.assertFalse(window.log_drawer.isHidden())
        window.is_running = False
        window.hide_log_drawer()
        window.on_file_failed(Path("failed.pptx"), "验证失败")
        self.assertFalse(window.log_drawer.isHidden())
        self.assertFalse(window.log_drawer_timer.isActive())
        window.on_finished([], [("failed.pptx", "验证失败")], False)
        self.assertIn(window.text["status_failed"], window.log_shelf.text())

        window.advanced_button.click()
        self.assertFalse(window.advanced_panel.isHidden())
        self.assertEqual(window.advanced_button.text(), "高级设置 · 收起高级设置")

        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = Path(temp_dir) / "sample.pptx"
            docx_path = Path(temp_dir) / "sample.docx"
            pptx_path.touch()
            docx_path.touch()
            window.set_files([pptx_path, docx_path])
            pptx_path = pptx_path.resolve()
            docx_path = docx_path.resolve()
            self.assertEqual(window.checked_paths, {pptx_path, docx_path})
            self.assertEqual(window.selection_summary_label.text(), "已选 2/2 个文件")
            window.output_format_select.setCurrentIndex(
                window.output_format_select.findData("pptx")
            )
            window.output_mode_select.setCurrentIndex(
                window.output_mode_select.findData("image")
            )
            window.update_image_quality_control()
            self.assertFalse(window.keep_videos_group.isHidden())
            self.assertTrue(window.keep_videos_checkbox.isEnabled())
            self.assertEqual(window.keep_videos_checkbox.text(), "不保留")
            self.assertGreaterEqual(
                window.keep_videos_group.rect().right()
                - window.keep_videos_checkbox.geometry().right(),
                8,
            )
            window.keep_videos_checkbox.click()
            self.assertEqual(window.keep_videos_checkbox.text(), "加水印并回填")
            self.assertTrue(
                window.current_settings(
                    window.input_paths
                ).preserve_videos_in_image_pptx
            )
            first_row = window.file_list.itemWidget(window.file_list.item(0))
            first_toggle = first_row.findChild(QPushButton, "fileTypeToggle")
            self.assertIsNotNone(first_toggle)
            self.assertFalse(first_toggle.icon().isNull())
            self.assertEqual(first_row.property("included"), "true")
            first_toggle.setChecked(False)
            self.assertEqual(first_row.property("included"), "false")
            self.assertEqual(window.pending_paths(), [docx_path])
            self.assertEqual(window.selection_summary_label.text(), "已选 1/2 个文件")
            first_toggle.setChecked(True)
            window.output_format_select.setCurrentIndex(
                window.output_format_select.findData("pptx")
            )
            window.update_idle_status_label()
            self.assertEqual(window.queue_count_label.text(), "共 2 个文件")
            self.assertIn("DOCX / PDF", window.queue_count_label.toolTip())

            preview_paths = []
            for index in range(3):
                preview_path = Path(temp_dir) / f"preview-{index}.png"
                Image.new("RGB", (160, 90), (index * 40, 80, 120)).save(preview_path)
                preview_paths.append(preview_path)
            window.current_preview = PreviewArtifacts(
                temp_root=None,
                source_key=("preview", 0, 0),
                original_paths=preview_paths,
                preview_paths=preview_paths,
                total_pages=8,
            )
            window.update_preview_display()
            self.assertEqual(window.preview_thumbnail_list.count(), 3)
            self.assertFalse(window.preview_image_label_secondary.isHidden())
            self.assertEqual(window.preview_page_label.text(), "1-2 / 8")
            window.update_preview_geometry()
            self.assertLessEqual(
                (window.preview_image_label.height() * 2)
                + window.preview_pages_layout.spacing(),
                window.preview_scroll_area.viewport().height() + 2,
            )
            self.assertFalse(window.preview_thumbnail_list.isHidden())
            self.assertFalse(window.preview_prev_button.isHidden())
            self.assertFalse(window.preview_next_button.isHidden())
            self.assertEqual(window.preview_prev_button.text(), "上一组")
            self.assertEqual(
                window.preview_prev_button.accessibleName(),
                window.text["preview_prev"],
            )
            window.preview_thumbnail_list.setCurrentRow(2)
            self.assertEqual(window.current_preview_page, 2)

            mixed_paths = [
                Path(temp_dir) / "mixed-landscape.png",
                Path(temp_dir) / "mixed-portrait.png",
            ]
            Image.new("RGB", (160, 90), (20, 80, 120)).save(mixed_paths[0])
            Image.new("RGB", (90, 160), (40, 80, 120)).save(mixed_paths[1])
            window.current_preview = PreviewArtifacts(
                temp_root=None,
                source_key=("mixed-preview", 0, 0),
                original_paths=mixed_paths,
                preview_paths=mixed_paths,
                total_pages=2,
            )
            window.current_preview_page = 0
            window.update_preview_display()
            self.assertEqual(window.preview_group_size(), 1)
            self.assertTrue(window.preview_image_label_secondary.isHidden())

            portrait_paths = []
            for index in range(2):
                portrait_path = Path(temp_dir) / f"portrait-{index}.png"
                Image.new("RGB", (90, 160), (index * 40, 80, 120)).save(portrait_path)
                portrait_paths.append(portrait_path)
            window.current_preview = PreviewArtifacts(
                temp_root=None,
                source_key=("portrait-preview", 0, 0),
                original_paths=portrait_paths,
                preview_paths=portrait_paths,
                total_pages=2,
            )
            window.current_preview_page = 0
            window.update_preview_display()
            self.assertTrue(window.preview_image_label_secondary.isHidden())
            self.assertEqual(
                window.preview_page_label.text(), "第 1/2 页预览 · 共 2 页"
            )

            window.change_preview_page(1)
            self.assertEqual(window.current_preview_page, 1)

            window.current_preview = PreviewArtifacts(
                temp_root=None,
                source_key=("single-preview", 0, 0),
                original_paths=preview_paths[:1],
                preview_paths=preview_paths[:1],
                total_pages=1,
            )
            window.update_preview_display()
            self.assertFalse(window.preview_prev_button.isHidden())
            self.assertFalse(window.preview_next_button.isHidden())
            self.assertFalse(window.preview_prev_button.isEnabled())
            self.assertFalse(window.preview_next_button.isEnabled())
            self.assertFalse(window.preview_thumbnail_list.isHidden())
            self.assertTrue(window.preset_hint_label.wordWrap())
            self.assertGreaterEqual(window.preset_hint_label.minimumHeight(), 48)
        window.clear_preview(window.text["preview_waiting"])
        self.assertFalse(window.preview_page_label.isVisible())
        self.assertFalse(window.preview_thumbnail_list.isHidden())
        self.assertEqual(window.preview_thumbnail_list.count(), 1)
        window.input_paths.clear()
        window.clear_preview(window.text["preview_waiting"])
        self.assertFalse(window.preview_image_label_secondary.isHidden())
        window.resize(960, 652)
        window.show()
        QApplication.processEvents()
        compact_preview_width = window.preview_image_label.width()
        window.resize(1512, 949)
        QApplication.processEvents()
        self.assertGreater(
            window.preview_image_label.width(),
            compact_preview_width,
        )
        window.close()

    def test_watermark_english_mode_localizes_preview_and_file_controls(self) -> None:
        with patch("pptx_output_watermark.gui.detect_language", return_value="en"):
            window = MainWindow()
        self.assertEqual(
            window.preview_thumbnail_list.item(0).text(),
            "Select a file to show page thumbnails",
        )
        self.assertTrue(
            any(
                button.text() == "Collapse"
                for button in window.log_drawer.findChildren(QPushButton)
            )
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "sample.pptx"
            source.touch()
            window.set_files([source])
            row = window.file_list.itemWidget(window.file_list.item(0))
            toggle = row.findChild(QPushButton, "fileTypeToggle")
            self.assertIn("Click to include or exclude", toggle.toolTip())
        window.clear_preview(window.text["preview_waiting"])
        self.assertEqual(
            window.preview_thumbnail_list.item(0).text(),
            "Select a file to show page thumbnails",
        )
        window.close()

    def test_compression_page_prioritizes_settings_over_queue_width(self) -> None:
        with patch("pptx_video_compactor_gui.detect_language", return_value="zh"):
            window = CompressionMainWindow()
        self.assertEqual(window.log_hover_timer.interval(), 1000)
        self.assertEqual(window.minimumWidth(), 880)
        self.assertEqual(window.left_pane.minimumWidth(), 320)
        self.assertEqual(window.left_pane.maximumWidth(), 320)
        window.update_responsive_layout(960)
        for control in (
            window.video_threshold_spinbox,
            window.image_threshold_spinbox,
        ):
            self.assertGreaterEqual(
                control.height(), control.minimumSizeHint().height()
            )
        self.assertEqual(window.left_pane.width(), 260)
        self.assertGreaterEqual(window.settings_controls_widget.height(), 30)
        self.assertEqual(
            window.right_pane.parentWidget(), window.left_pane.parentWidget()
        )
        self.assertEqual(window.video_threshold_spinbox.width(), 80)
        self.assertEqual(window.image_threshold_spinbox.width(), 80)
        window.append_log("[INFO] 日志浮层验证")
        window.resize(960, 652)
        window.show()
        QApplication.processEvents()
        window.show_log_drawer(auto_hide=False)
        self.assertFalse(window.log_drawer.isHidden())
        shelf_origin = window.log_shelf.mapTo(
            window.content_widget, window.log_shelf.rect().topLeft()
        )
        self.assertEqual(window.log_drawer.x(), shelf_origin.x())
        self.assertEqual(window.log_drawer.width(), window.log_shelf.width())
        self.assertLessEqual(
            window.log_drawer.height(),
            int(window.content_widget.height() * 0.35) + 1,
        )
        window.hide_log_drawer()
        self.assertTrue(window.log_drawer.isHidden())
        window.on_file_failed(Path("failed.pptx"), "验证失败")
        self.assertFalse(window.log_drawer.isHidden())
        self.assertFalse(window.log_drawer_timer.isActive())

        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = Path(temp_dir) / "long-presentation-name.pptx"
            pptx_path.write_bytes(b"pptx")
            window.set_files([pptx_path])
            self.assertEqual(window.total_estimate_label.text(), "共 1 个")
            self.assertIn("预计输出", window.total_estimate_label.toolTip())
            row = window.file_list.itemWidget(window.file_list.item(0))
            self.assertIsNotNone(row)
            metadata = row.findChild(QLabel, "fileMeta")
            self.assertIsNotNone(metadata)
            self.assertIn("PPTX", metadata.text())
            self.assertIn("估算", metadata.text())
            self.assertEqual(window.results_tree.topLevelItemCount(), 1)
            output_path = Path(temp_dir) / "compressed.pptx"
            output_path.write_bytes(b"x")
            window.output_paths[pptx_path.resolve()] = output_path
            window.file_statuses[pptx_path.resolve()] = "done"
            window.refresh_results_tree()
            result = window.results_tree.topLevelItem(0)
            self.assertEqual(result.text(0), pptx_path.name)
            self.assertEqual(result.text(5), window.text["done_marker"])
            self.assertEqual(result.text(6), str(output_path))
        window.close()

    def test_compression_assessment_row_uses_measured_width(self) -> None:
        with patch("pptx_video_compactor_gui.detect_language", return_value="zh"):
            window = CompressionMainWindow()
        window.show()
        try:
            self.assertEqual(
                [
                    window.standard_encoder_select.itemText(index)
                    for index in range(window.standard_encoder_select.count())
                ],
                ["自动硬件", "仅 CPU", "优先 GPU"],
            )
            self.assertEqual(window.assessment_controls_layout.spacing(), 16)
            self.assertEqual(window.assessment_actions_layout.spacing(), 16)
            for width, expected_actions_visible in (
                (880, True),
                (1000, True),
                (1100, True),
                (1180, True),
                (1280, False),
                (1440, False),
            ):
                window.resize(width, 620)
                QApplication.processEvents()
                self.assertEqual(
                    window.assessment_row_actions_widget.isVisibleTo(window),
                    expected_actions_visible,
                    f"unexpected assessment layout at {width}px",
                )
                host = (
                    window.assessment_row_actions_widget
                    if expected_actions_visible
                    else window.assessment_row_primary_widget
                )
                for button in (window.audit_button, window.optimize_button):
                    self.assertIs(button.parentWidget(), host)
                    origin = button.mapTo(host, QPoint(0, 0))
                    self.assertGreaterEqual(origin.x(), 0)
                    self.assertLessEqual(origin.x() + button.width(), host.width())

                centers = []
                for control in window.assessment_primary_controls:
                    origin = control.mapTo(
                        window.assessment_row_primary_widget, QPoint(0, 0)
                    )
                    centers.append(origin.y() + control.height() / 2)
                    self.assertGreaterEqual(
                        control.height(), control.minimumSizeHint().height()
                    )
                    self.assertLessEqual(
                        origin.x() + control.width(),
                        window.assessment_row_primary_widget.width(),
                    )
                self.assertLessEqual(max(centers) - min(centers), 1)
                row_bottom = window.assessment_row_primary_widget.mapTo(
                    window.right_pane,
                    QPoint(0, window.assessment_row_primary_widget.height()),
                ).y()
                results_top = window.results_tree.mapTo(
                    window.right_pane, QPoint(0, 0)
                ).y()
                self.assertGreaterEqual(results_top - row_bottom, 8)
                if expected_actions_visible:
                    self.assertTrue(
                        window.assessment_controls_layout.alignment()
                        & Qt.AlignmentFlag.AlignLeft
                    )
                else:
                    self.assertTrue(
                        window.assessment_controls_layout.alignment()
                        & Qt.AlignmentFlag.AlignRight
                    )

            last_control = window.optimize_button
            last_origin = last_control.mapTo(
                window.assessment_row_primary_widget, QPoint(0, 0)
            )
            self.assertEqual(
                last_origin.x() + last_control.width(),
                window.assessment_row_primary_widget.width(),
            )

            window.resize(1100, 620)
            QApplication.processEvents()
            self.assertGreater(
                window._assessment_row_required_width(),
                window.assessment_row_primary_widget.width(),
            )

            window.resize(1280, 620)
            window.forced_button.show()
            window.update_responsive_layout(window.content_widget.width())
            QApplication.processEvents()
            self.assertTrue(window.assessment_row_actions_widget.isVisibleTo(window))
            self.assertIs(
                window.forced_button.parentWidget(),
                window.assessment_row_actions_widget,
            )

            window.resize(1440, 620)
            QApplication.processEvents()
            self.assertFalse(window.assessment_row_actions_widget.isVisibleTo(window))
            self.assertIs(
                window.forced_button.parentWidget(),
                window.assessment_row_primary_widget,
            )
        finally:
            window.close()

    def test_workspace_controls_stay_in_bounds_across_size_matrix(self) -> None:
        window = ToolboxMainWindow()
        control_types = (
            QAbstractButton,
            QAbstractSlider,
            QAbstractSpinBox,
            QComboBox,
            QLineEdit,
        )
        try:
            for width in (880, 1000, 1100, 1180, 1280, 1440):
                for height in (560, 620, 700):
                    window.resize(width, height)
                    window.show()
                    QApplication.processEvents()
                    for tab_index in range(window.tabs.count()):
                        window.tabs.setCurrentIndex(tab_index)
                        QApplication.processEvents()
                        page = window.tabs.currentWidget()
                        for control in page.findChildren(QWidget):
                            if not isinstance(control, control_types):
                                continue
                            if not control.isVisibleTo(page) or control.isWindow():
                                continue
                            parent = control.parentWidget()
                            if (
                                parent is None
                                or isinstance(parent, QAbstractScrollArea)
                                or isinstance(
                                    parent.parentWidget(), QAbstractScrollArea
                                )
                            ):
                                continue
                            origin = control.mapTo(page, QPoint(0, 0))
                            self.assertGreaterEqual(
                                origin.x(),
                                0,
                                f"{window.tabs.tabText(tab_index)} {width}x{height}",
                            )
                            self.assertGreaterEqual(origin.y(), 0)
                            self.assertLessEqual(
                                origin.x() + control.width(),
                                page.width(),
                                f"{window.tabs.tabText(tab_index)} {width}x{height}",
                            )
                            self.assertLessEqual(
                                origin.y() + control.height(), page.height()
                            )
                            self.assertGreaterEqual(
                                control.height(),
                                control.minimumSizeHint().height(),
                                f"{window.tabs.tabText(tab_index)} {width}x{height}",
                            )
        finally:
            window.close()

    def test_shell_embedded_pages_keep_top_border_visible(self) -> None:
        window = ToolboxMainWindow()
        try:
            window.resize(1440, 900)
            window.show()
            QApplication.processEvents()
            for tab_index in range(window.tabs.count()):
                page = window.tabs.widget(tab_index)
                self.assertIsNotNone(page)
                margins = page.layout().contentsMargins()  # type: ignore[union-attr]
                self.assertEqual(margins.top(), 6, window.tabs.tabText(tab_index))
                content_frame = next(
                    (
                        page.findChild(QWidget, object_name)
                        for object_name in (
                            "projectBar",
                            "rightCard",
                            "leftCard",
                        )
                        if page.findChild(QWidget, object_name) is not None
                    ),
                    None,
                )
                self.assertIsNotNone(content_frame, window.tabs.tabText(tab_index))
                self.assertGreaterEqual(
                    content_frame.mapTo(page, QPoint(0, 0)).y(),  # type: ignore[union-attr]
                    6,
                    window.tabs.tabText(tab_index),
                )
        finally:
            window.close()

    def test_video_library_uses_transient_pptx_workflow(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "selected.pptx"
            source.touch()
            window = VideoLibraryMainWindow()
            self.assertEqual(window.minimumWidth(), 880)
            self.assertEqual(window.help_button.size().width(), 30)
            self.assertEqual(window.help_button.size().height(), 30)
            self.assertEqual(window.help_button.objectName(), "helpIconButton")
            self.assertTrue(window.video_tree.isHidden())
            self.assertTrue(window.log_output.isHidden())
            self.assertTrue(window.detail_drawer.isHidden())
            window.project = None
            window.refresh_views()
            self.assertEqual(window.windowTitle(), "PPTX 视频资产库")
            self.assertIn("font-size: 11px", window.styleSheet())
            self.assertIn("min-height: 28px", window.styleSheet())
            self.assertEqual(
                window.content_splitter.orientation(), Qt.Orientation.Vertical
            )
            self.assertTrue(window.log_panel_button.isCheckable())
            self.assertEqual(window.archive_button.text(), "归档 PPTX 视频")
            self.assertEqual(window.choose_input_button.text(), "选择 PPTX（可多选）")
            self.assertNotEqual(
                window.choose_input_button.objectName(), "primaryAction"
            )
            self.assertIn("可拖入或多选 PPTX", window.input_summary.text())
            self.assertEqual(window.external_import_button.text(), "导入外部视频并匹配")
            self.assertEqual(window.upgrade_button.text(), "高清回填 PPTX（另存）")
            self.assertEqual(window.health_button.text(), "库体检")
            self.assertIn("不会修改视频库", window.health_button.toolTip())
            self.assertEqual(window.preview_button.text(), "播放")
            self.assertIsNone(window.external_import_button.parentWidget())
            self.assertEqual(
                window.import_button.parentWidget(), window.library_action_row
            )
            self.assertEqual(
                window.activate_button.parentWidget(), window.library_action_row
            )
            self.assertEqual(
                window.review_button.parentWidget(), window.library_action_row
            )
            self.assertEqual(
                [
                    window.library_action_row.layout().itemAt(index).widget().text()
                    for index in range(window.library_action_row.layout().count())
                ],
                [
                    "核实版本",
                    "添加版本",
                    "设为高清源",
                    "整理视频库",
                    "待清理 (0)",
                    "更多操作",
                ],
            )
            self.assertEqual(window.attention_filter_combo.count(), 5)
            self.assertTrue(window.attention_filter_combo.isHidden())
            window.update_responsive_layout(960)
            self.assertFalse(window.attention_filter_combo.isHidden())
            self.assertTrue(window.library_stat_buttons["all"].isHidden())
            window.update_responsive_layout(1440)
            self.assertTrue(window.attention_filter_combo.isHidden())
            self.assertFalse(window.library_stat_buttons["all"].isHidden())
            self.assertEqual(window.library_stat_buttons["all"].text(), "全部 0")
            self.assertTrue(window.library_stat_buttons["all"].isChecked())
            self.assertEqual(window.library_version_label.text(), "0 版本")
            self.assertEqual(window.library_stats_label.text(), "显示 0/0")
            self.assertEqual(window.content_splitter.widget(0).minimumHeight(), 76)
            self.assertTrue(window.workflow_settings.isHidden())
            window._toggle_workflow_settings()
            self.assertFalse(window.workflow_settings.isHidden())
            self.assertGreaterEqual(window.workflow_settings.minimumHeight(), 38)
            window._toggle_workflow_settings()
            self.assertTrue(window.workflow_settings.isHidden())
            self.assertEqual(
                window.library_filter_input.placeholderText(),
                "筛选名称、路径或哈希",
            )
            self.assertEqual(window.library_filter_input.maximumWidth(), 320)
            self.assertEqual(window.video_tree.headerItem().text(1), "分辨率")
            self.assertEqual(window.video_tree.headerItem().text(2), "时长")
            window._sort_by_header(2)
            self.assertEqual(window.library_sort_mode, "duration")
            self.assertTrue(window.library_sort_descending)
            self.assertTrue(
                self.settings.value("video_library/sort_descending", False, bool)
            )
            window._sort_by_header(2)
            self.assertFalse(window.library_sort_descending)
            self.assertFalse(
                self.settings.value("video_library/sort_descending", True, bool)
            )
            window._sort_by_header(6)
            self.assertEqual(window.library_sort_mode, "path")
            self.assertTrue(window.library_empty.isVisibleTo(window))
            self.assertTrue(window.video_tree.isHidden())
            self.assertFalse(window.archive_button.isEnabled())
            self.assertEqual(
                window.category_input.sizePolicy().horizontalPolicy(),
                QSizePolicy.Policy.Expanding,
            )
            self.assertEqual(
                window.library_empty.parentWidget()
                .layout()
                .indexOf(window.library_empty),
                1,
            )
            self.assertTrue(window.library_actions.isHidden())
            self.assertEqual(
                window.library_actions.sizePolicy().verticalPolicy(),
                QSizePolicy.Policy.Fixed,
            )
            self.assertTrue(window.status_frame.isHidden())
            self.assertEqual(window.log_shelf.height(), 26)
            self.assertTrue(window.log_output.isHidden())
            window.append_log("test")
            self.assertFalse(window.log_output.isHidden())
            self.assertIn("test", window.log_shelf.text())
            self.assertFalse(hasattr(window, "deck_tree"))
            window.set_files([source, Path(temp_dir) / "missing.pptx"])
            self.assertEqual(window.input_paths, [source.resolve()])
            self.assertTrue(
                any(
                    source.name in widget.text()
                    for widget in window.workflow_chip_widgets
                    if isinstance(widget, QPushButton)
                )
            )
            second = Path(temp_dir) / "second.pptx"
            second.touch()
            with patch.object(
                QFileDialog,
                "getOpenFileNames",
                return_value=([str(source), str(second)], "PowerPoint (*.pptx)"),
            ):
                window.choose_input_pptx()
            self.assertEqual(window.input_paths, [source.resolve(), second.resolve()])
            self.assertEqual(
                sum(
                    widget.objectName() == "inputChip"
                    for widget in window.workflow_chip_widgets
                ),
                2,
            )
            self.assertIn(str(second.resolve()), window.input_summary.toolTip())
            mime_data = QMimeData()
            mime_data.setUrls([QUrl.fromLocalFile(str(source))])
            drag_event = QDragEnterEvent(
                QPoint(10, 10),
                Qt.DropAction.CopyAction,
                mime_data,
                Qt.MouseButton.NoButton,
                Qt.KeyboardModifier.NoModifier,
            )
            QApplication.sendEvent(window.input_summary, drag_event)
            self.assertTrue(drag_event.isAccepted())
            drop_event = QDropEvent(
                QPointF(10, 10),
                Qt.DropAction.CopyAction,
                mime_data,
                Qt.MouseButton.NoButton,
                Qt.KeyboardModifier.NoModifier,
            )
            QApplication.sendEvent(window.input_summary, drop_event)
            self.assertEqual(window.input_paths, [source.resolve()])
            video = Path(temp_dir) / "external.mp4"
            video.touch()
            video_mime = QMimeData()
            video_mime.setUrls([QUrl.fromLocalFile(str(video))])
            video_drag = QDragEnterEvent(
                QPoint(10, 10),
                Qt.DropAction.CopyAction,
                video_mime,
                Qt.MouseButton.NoButton,
                Qt.KeyboardModifier.NoModifier,
            )
            video_drop = QDropEvent(
                QPointF(10, 10),
                Qt.DropAction.CopyAction,
                video_mime,
                Qt.MouseButton.NoButton,
                Qt.KeyboardModifier.NoModifier,
            )
            with patch.object(window, "import_external_paths") as import_paths:
                QApplication.sendEvent(window.video_tree.viewport(), video_drag)
                QApplication.sendEvent(window.video_tree.viewport(), video_drop)
            self.assertTrue(video_drag.isAccepted())
            self.assertTrue(video_drop.isAccepted())
            import_paths.assert_called_once()
            window.clear_input_paths()
            self.assertEqual(window.input_paths, [])
            self.assertIn("可拖入或多选", window.input_summary.text())
            project_root = Path(temp_dir) / "topic-library"
            VideoProject.create(project_root)
            window.open_project(project_root)
            self.assertIn(project_root.name, window.project_label.text())
            self.assertIn(str(project_root.resolve()), window.project_label.toolTip())
            self.assertEqual(window.open_project_button.text(), "切换 / 打开视频库")
            self.assertEqual(window.project_menu_button.text(), "更多")
            window.close()

    def test_ai_video_merge_uses_the_suggested_primary_family(self) -> None:
        window = VideoLibraryMainWindow()
        primary = {"id": "primary-family", "name": "主视频"}
        duplicate = {"id": "duplicate-family", "name": "重复视频"}
        variants = {
            "primary-variant": (primary, {"id": "primary-variant"}),
            "duplicate-variant": (duplicate, {"id": "duplicate-variant"}),
        }
        merged: list[tuple[str, str, bool]] = []
        project = SimpleNamespace(
            find_variant=lambda variant_id: variants[variant_id],
            family_merge_impact=lambda _source, _target: {
                "variant_count": 1,
                "reference_count": 2,
                "known_hash_count": 1,
            },
            merge_families=lambda source, target, confirmed_same_content: merged.append(
                (source, target, confirmed_same_content)
            ),
        )
        window.project = project
        with (
            patch.object(
                QMessageBox,
                "question",
                return_value=QMessageBox.StandardButton.Yes,
            ),
            patch.object(window, "refresh_views"),
        ):
            count = window._review_ai_video_merge_groups(
                [
                    {
                        "item_ids": ["duplicate-variant", "primary-variant"],
                        "primary_id": "primary-variant",
                        "confidence": 0.9,
                        "reason": "画面与规格一致",
                    }
                ]
            )

        self.assertEqual(count, 1)
        self.assertEqual(
            merged,
            [("duplicate-family", "primary-family", True)],
        )
        window.project = None
        window.close()

    def test_video_match_dialog_selects_candidate_without_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "candidate.mp4"
            source.touch()
            item = {
                "source": str(source),
                "sha256": "a" * 64,
                "metadata": {"width": 640, "height": 360, "duration_sec": 2},
                "candidates": [
                    {
                        "family_id": "family",
                        "family_name": "候选视频.wmv",
                        "source_path": str(source),
                        "source_sha256": "b" * 64,
                        "width": 1920,
                        "height": 1080,
                        "duration_sec": 2,
                        "score": 96,
                        "strict_match": False,
                        "confidence": {
                            "frame_total_distance": 8,
                            "audio_consistent": True,
                        },
                    }
                ],
            }
            with patch(
                "pptx_tools.video_manager_gui.create_video_thumbnail",
                return_value=False,
            ):
                parent = VideoLibraryMainWindow()
                dialog = VideoMatchDialog(
                    parent,
                    item,
                    allow_new_family=False,
                    allow_remember=True,
                )
            self.assertEqual(dialog.selected_family_id, "family")
            self.assertTrue(dialog.remember)
            self.assertLessEqual(dialog.minimumWidth(), 960)
            self.assertEqual(dialog.tree.topLevelItem(0).text(0), "候选视频")
            self.assertEqual(dialog.source_preview.minimumHeight(), 120)
            self.assertTrue(
                any(
                    label.objectName() == "dialogTitle"
                    and label.text() == "确认视频匹配"
                    for label in dialog.findChildren(QLabel)
                )
            )
            self.assertEqual(dialog.link_button.objectName(), "primaryAction")
            headings = [
                label.text()
                for label in dialog.findChildren(QLabel)
                if label.objectName() == "previewHeading"
            ]
            self.assertIn("待核对视频  ·  10% / 50% / 90% 取帧", headings)
            dialog.close()
            parent.close()

    def test_video_thumbnail_follows_preview_width(self) -> None:
        label = ResponsiveVideoThumbnail()
        label.resize(720, 160)
        label.show()
        QApplication.processEvents()
        source = QPixmap(900, 180)
        source.fill(QColor("#334155"))
        label.set_source_pixmap(source)
        self.assertGreater(label.pixmap().width(), 420)
        self.assertLessEqual(label.pixmap().width(), label.contentsRect().width())

        label.resize(360, 160)
        QApplication.processEvents()
        self.assertLessEqual(label.pixmap().width(), label.contentsRect().width())
        label.close()

    def test_pptx_upgrade_review_distinguishes_keep_once_and_remember(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            current = root / "current.mp4"
            target = root / "target.mp4"
            current.touch()
            target.touch()
            item = {
                "media_path": "ppt/media/media1.mp4",
                "source": str(current),
                "sha256": "a" * 64,
                "metadata": {"width": 640, "height": 360, "duration_sec": 2},
                "match_kind": "content",
                "family_id": "family",
                "family_name": "候选视频",
                "target_source": str(target),
                "already_high_quality": False,
                "occurrences": [{"slide_path": "ppt/slides/slide1.xml"}],
            }
            families = [
                {
                    "id": "family",
                    "name": "候选视频",
                    "source_path": str(target),
                    "source_sha256": "b" * 64,
                }
            ]
            with patch(
                "pptx_tools.video_manager_gui.create_video_thumbnail",
                return_value=False,
            ):
                parent = VideoLibraryMainWindow()
                dialog = PptxUpgradeReviewDialog(
                    parent, root / "demo.pptx", [item], families
                )

            overrides, remembered, kept = dialog.decisions()
            self.assertEqual(dialog.tree.minimumWidth(), 560)
            self.assertGreaterEqual(dialog.minimumWidth(), 1080)
            self.assertTrue(
                any(
                    label.objectName() == "dialogTitle"
                    and label.text() == "确认 PPTX 高清回填"
                    for label in dialog.findChildren(QLabel)
                )
            )
            self.assertEqual(overrides, {"ppt/media/media1.mp4": "family"})
            self.assertEqual(remembered, {"ppt/media/media1.mp4"})
            self.assertEqual(kept, set())

            dialog.keep_radio.setChecked(True)
            self.assertEqual(dialog.decisions(), ({}, set(), {"ppt/media/media1.mp4"}))

            dialog.remember_radio.setChecked(True)
            self.assertEqual(
                dialog.decisions(),
                (
                    {"ppt/media/media1.mp4": "family"},
                    {"ppt/media/media1.mp4"},
                    set(),
                ),
            )
            dialog.close()
            parent.close()

    def test_video_library_formats_duration_for_scanning(self) -> None:
        self.assertEqual(_format_duration(0), "0:00")
        self.assertEqual(_format_duration(59.6), "1:00")
        self.assertEqual(_format_duration(717.397), "11:57")
        self.assertEqual(_format_duration(3661), "1:01:01")

    def test_video_library_preview_plays_in_detail_drawer(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.mp4"
            source.touch()
            family = {"id": "family"}
            variant = {"id": "variant"}
            window = VideoLibraryMainWindow()
            player = MagicMock()
            player.source.return_value = QUrl()
            player.playbackState.return_value = QMediaPlayer.PlaybackState.StoppedState
            window.detail_media_player = player
            window.project = SimpleNamespace(
                source_variant=lambda _family: variant,
                require_variant_path=lambda _variant: source,
            )
            with patch.object(
                window, "selected_family_variant", return_value=(family, None)
            ):
                window.preview_selected()
            self.assertEqual(
                Path(player.setSource.call_args.args[0].toLocalFile()), source
            )
            player.play.assert_called_once_with()
            self.assertIs(
                window.detail_media_stack.currentWidget(),
                window.detail_video,
            )
            window.close()

    def test_video_library_parent_move_uses_the_whole_family(self) -> None:
        window = VideoLibraryMainWindow()
        family = {
            "id": "family",
            "category": "旧分类",
            "variants": [{"id": "a"}, {"id": "b"}],
        }
        project = MagicMock()
        window.project = project
        with (
            patch.object(
                window, "selected_family_variant", return_value=(family, None)
            ),
            patch(
                "pptx_tools.video_manager_gui.QInputDialog.getText",
                return_value=("新分类/项目", True),
            ),
            patch.object(window, "run_operation") as run_operation,
        ):
            window.move_selected_variant()

        self.assertEqual(run_operation.call_args.args[0], "正在移动视频族")
        operation = run_operation.call_args.args[1]
        operation(MagicMock(), lambda: False)
        project.move_family.assert_called_once_with("family", "新分类/项目")
        window.close()

    def test_video_library_filter_matches_family_variant_path_and_hash(self) -> None:
        window = VideoLibraryMainWindow()
        family = QTreeWidgetItem(
            ["示例验证平台", "1920×1080", "120.7 秒", "", "", "", ""]
        )
        variant = QTreeWidgetItem(
            [
                "高清源 · 验证平台录屏",
                "1920×922",
                "120.7 秒",
                "",
                "95c9d42a",
                "正常",
                "media/示例项目/验证平台录屏.mp4",
            ]
        )
        family.addChild(variant)
        family.setData(0, REVIEW_TAGS_ROLE, "review,unlinked")
        other = QTreeWidgetItem(["航空制造演示", "1280×720", "60 秒", "", "", "", ""])
        other.setData(0, REVIEW_TAGS_ROLE, "")
        window.video_tree.addTopLevelItem(family)
        window.video_tree.addTopLevelItem(other)

        for query in ("示例", "1920×922", "95c9d42a", "验证平台录屏.mp4"):
            window.library_filter_input.setText(query)
            self.assertFalse(family.isHidden(), query)
            self.assertFalse(variant.isHidden(), query)
            self.assertTrue(other.isHidden(), query)

        window.library_filter_input.clear()
        self.assertFalse(family.isHidden())
        self.assertFalse(other.isHidden())
        window.attention_filter_combo.setCurrentIndex(
            window.attention_filter_combo.findData("review")
        )
        self.assertFalse(family.isHidden())
        self.assertTrue(other.isHidden())
        window.library_filter_input.setText("没有这个视频")
        self.assertTrue(window.video_tree.isHidden())
        self.assertFalse(window.library_empty.isHidden())
        self.assertIn("没有符合当前条件", window.library_empty.text())
        self.assertIn("显示 0/2", window.library_stats_label.text())
        window.library_filter_input.clear()
        window.attention_filter_combo.setCurrentIndex(
            window.attention_filter_combo.findData("all")
        )
        self.assertFalse(window.video_tree.isHidden())
        window.close()

    def test_video_library_stats_surface_unlinked_families(self) -> None:
        variant = {
            "id": "variant",
            "label": "source",
            "sha256": "a" * 64,
            "path": "media/source.mp4",
            "size_bytes": 1024,
            "width": 1920,
            "height": 1080,
            "duration_sec": 10.0,
            "bitrate_kbps": 1000,
        }
        damaged_variant = {
            **variant,
            "id": "damaged",
            "label": "damaged",
            "sha256": "b" * 64,
            "path": "media/damaged.mp4",
            "probe_error": "moov atom not found",
        }
        family = {
            "id": "family",
            "name": "未关联视频",
            "source_variant_id": "variant",
            "variants": [variant, damaged_variant],
            "known_hashes": [variant["sha256"], damaged_variant["sha256"]],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            window = VideoLibraryMainWindow()
            window.project = SimpleNamespace(
                root=Path(temp_dir),
                pending_cleanup=lambda: [],
                families=lambda: [family],
                decks=lambda: [],
                family=lambda _family_id: family,
                find_variant=lambda variant_id: (
                    family,
                    next(
                        item for item in family["variants"] if item["id"] == variant_id
                    ),
                ),
                source_variant=lambda _family: variant,
                status=lambda _variant: "available",
                _variant_quality_key=lambda _variant: (1920 * 1080, 1000, 1024),
            )
            window.refresh_views()

            self.assertEqual(window.library_stat_buttons["all"].text(), "全部 1")
            self.assertEqual(window.library_stat_buttons["review"].text(), "待核对 1")
            self.assertEqual(window.library_stat_buttons["unlinked"].text(), "无关联 1")
            self.assertEqual(window.library_stat_buttons["multi"].text(), "多版本 1")
            self.assertEqual(
                window.library_stat_buttons["abnormal"].text(), "文件异常 1"
            )
            self.assertIn("独立素材", window.library_stat_buttons["unlinked"].toolTip())
            self.assertIn(
                "不会修改视频库", window.library_stat_buttons["review"].toolTip()
            )
            self.assertEqual(window.library_version_label.text(), "2 版本")
            self.assertEqual(window.library_stats_label.text(), "显示 1/1")
            self.assertIn("不会修改视频库", window.attention_filter_combo.toolTip())
            self.assertEqual(window.review_button.text(), "核实版本")
            self.assertIn("当前视频族", window.review_button.toolTip())
            self.assertIn("路径关联", window.relink_button.toolTip())
            self.assertTrue(window.more_actions_menu.toolTipsVisible())
            window._sync_more_actions()
            self.assertTrue(
                all(action.toolTip() for action, _button in window.more_action_targets)
            )
            self.assertEqual(window.video_tree.headerItem().text(5), "关联状态")
            self.assertEqual(
                [
                    action.text()
                    for action in window.more_actions_menu.actions()
                    if action.text()
                ],
                [
                    "核对与归并",
                    "AI 整理建议",
                    "归并视频",
                    "版本管理",
                    "设为高清源",
                    "文件与恢复",
                    "重命名",
                    "移动文件",
                    "查找丢失",
                    "刷新状态",
                    "异常处理",
                    "隔离异常",
                    "库维护",
                    "整理视频库",
                    "待清理 (0)",
                    "导出",
                    "导入哈希目录",
                    "导出哈希目录",
                    "导出关联记录",
                ],
            )
            family_item = window.video_tree.topLevelItem(0)
            self.assertEqual(
                family_item.data(0, REVIEW_TAGS_ROLE),
                "abnormal,multi,review,unlinked",
            )
            self.assertEqual(family_item.child(1).text(5), "媒体不可读")
            with patch("pptx_tools.video_manager_gui._set_video_thumbnail"):
                window.video_tree.setCurrentItem(family_item)
            self.assertFalse(window.detail_drawer.isHidden())
            self.assertEqual(window.detail_title.text(), "未关联视频")
            self.assertEqual(
                window.detail_references.text(), "关联　2 个版本 · 0 个 PPTX 引用"
            )
            window.show()
            QApplication.processEvents()
            window._position_detail_drawer()
            self.assertGreaterEqual(window.detail_preview.height(), 160)
            self.assertEqual(
                window.detail_media_stack.height(), window.detail_preview.height()
            )
            window.detail_preview.set_source_pixmap(QPixmap(16, 9))
            with patch(
                "pptx_tools.video_manager_gui.QDialog.exec", return_value=0
            ) as show_preview:
                with patch("pptx_tools.video_manager_gui._set_video_thumbnail"):
                    window.detail_preview.clicked.emit()
            show_preview.assert_called_once()
            window.library_stat_buttons["abnormal"].click()
            self.assertEqual(window.attention_filter_combo.currentData(), "abnormal")
            self.assertTrue(window.library_stat_buttons["abnormal"].isChecked())
            window._sort_by_header(5)
            self.assertEqual(window.library_sort_mode, "review")
            window.close()

    def test_tooltips_wait_one_second_before_opening(self) -> None:
        # Use QProxyStyle(None) instead of QApplication.style() to avoid a
        # PySide6 6.11 crash: after app.setStyle(QProxyStyle_subclass(...)),
        # calling QApplication.style() triggers infinite recursion in
        # PySide::getWrapperForQObject → SIGSEGV. QProxyStyle(None) delegates
        # to the base style, which is all the assertion needs.
        style = DelayedTooltipStyle(None)
        self.assertEqual(
            style.styleHint(QStyle.StyleHint.SH_ToolTip_WakeUpDelay),
            1000,
        )

    def test_dialogs_are_centered_on_first_show(self) -> None:
        configure_ui_font(QApplication.instance())
        parent = QWidget()
        parent.setGeometry(100, 80, 960, 620)
        parent.show()
        QApplication.processEvents()
        dialog = SettingsDialog("zh", "", parent)
        dialog.show()
        QApplication.processEvents()
        self.assertIn("QDialog QPushButton", dialog.styleSheet())
        parent_center = parent.frameGeometry().center()
        dialog_center = dialog.frameGeometry().center()
        self.assertLessEqual(abs(parent_center.x() - dialog_center.x()), 2)
        self.assertLessEqual(abs(parent_center.y() - dialog_center.y()), 2)
        dialog.close()
        parent.close()

    def test_watermark_job_preflight_blocks_when_runtime_is_missing(self) -> None:
        window = MainWindow()
        settings = MagicMock()
        settings.export_options.return_value = MagicMock()
        missing = SimpleNamespace(
            required=True,
            available=False,
            status_code="engine_missing",
            name="LibreOffice",
            detail="No PDF export engine is available.",
            action_label="Install LibreOffice",
            action_url="https://www.libreoffice.org/download/",
        )
        dialog = MagicMock()
        dialog.addButton.return_value = object()
        dialog.clickedButton.return_value = None
        with (
            patch(
                "pptx_output_watermark.gui.dependency_statuses",
                return_value=[missing],
            ),
            patch(
                "pptx_output_watermark.gui.QMessageBox",
                return_value=dialog,
            ),
        ):
            self.assertFalse(
                window.confirm_runtime_dependencies(settings, [Path("sample.pptx")])
            )
        dialog.exec.assert_called_once()
        window.close()

    def test_selected_review_only_opens_groups_for_current_family(self) -> None:
        window = VideoLibraryMainWindow()
        groups = [
            {"family_ids": ["selected"], "title": "selected"},
            {"family_ids": ["other"], "title": "other"},
            {"family_ids": ["selected", "duplicate"], "title": "cross"},
        ]
        with patch.object(window, "_cleanup_scan_finished") as finished:
            window._selected_review_scan_finished(groups, 0.95, "selected")
        selected_groups, threshold = finished.call_args.args
        self.assertEqual(
            [group["title"] for group in selected_groups], ["selected", "cross"]
        )
        self.assertEqual(threshold, 0.95)
        window.close()

    def test_video_library_name_sort_keeps_best_same_name_source_first(self) -> None:
        low = {
            "id": "low",
            "name": "同名视频",
            "variants": [{"width": 640, "height": 360, "size_bytes": 10}],
        }
        high = {
            "id": "high",
            "name": "同名视频",
            "variants": [{"width": 1920, "height": 1080, "size_bytes": 20}],
        }
        window = VideoLibraryMainWindow()
        window.project = SimpleNamespace(
            families=lambda: [low, high],
            source_variant=lambda family: family["variants"][0],
        )
        window.library_sort_mode = "name"
        window.library_sort_descending = True

        self.assertEqual(
            [family["id"] for family in window._sorted_families()],
            ["high", "low"],
        )
        window.close()

    def test_video_library_help_describes_current_matching_and_storage(self) -> None:
        chinese = dict(HELP_SECTIONS["zh"])["PPTX 视频资产库"]
        english = dict(HELP_SECTIONS["en"])["PPTX Video Library"]

        self.assertIn("内容指纹", chinese)
        self.assertIn("待核对", chinese)
        self.assertIn("库体检", chinese)
        self.assertIn("_cleanup/", chinese)
        self.assertIn("QSettings", chinese)
        self.assertEqual(english.count("<h2>PPTX Video Library</h2>"), 1)

    def test_new_video_library_errors_are_reported(self) -> None:
        window = VideoLibraryMainWindow()
        with (
            patch.object(QFileDialog, "getExistingDirectory", return_value="/denied"),
            patch(
                "pptx_tools.video_manager_gui.VideoProject.create",
                side_effect=PermissionError("permission denied"),
            ),
            patch.object(window, "show_error") as show_error,
        ):
            window.choose_new_project()
        show_error.assert_called_once_with("permission denied")
        window.close()

    def test_cleanup_dialog_builds_expected_decisions(self) -> None:
        def candidate(variant_id, family_id, name, *, allowed=True, keep=False):
            return {
                "family_id": family_id,
                "family_name": name,
                "variant_id": variant_id,
                "label": "source",
                "profile": "original",
                "path": "media/x.mp4",
                "exists": True,
                "sha256": "ab" * 32,
                "width": 1920,
                "height": 1080,
                "duration_sec": 2.0,
                "bitrate_kbps": 1200,
                "size_bytes": 1024,
                "video_codec": "h264",
                "audio_codec": "aac",
                "has_audio": True,
                "ssim_to_best": 1.0 if keep else 0.97,
                "confidence": {"level": "high", "matched": True},
                "auto_allowed": allowed,
                "block_reasons": [] if allowed else ["音轨不一致"],
            }

        groups = [
            {
                "kind": "within_family",
                "family_ids": ["fam-1"],
                "title": "demo（2 个版本）",
                "best_variant_id": "v-best",
                "safe_to_apply": False,
                "candidates": [
                    candidate("v-best", "fam-1", "demo", keep=True),
                    candidate("v-small", "fam-1", "demo"),
                    candidate("v-bad", "fam-1", "demo", allowed=False),
                ],
                "recommendation": {
                    "keep_variant_id": "v-best",
                    "strategy": "keep_best",
                    "original_variant_id": "v-best",
                    "unify_available": True,
                    "alternatives": {},
                },
            },
            {
                "kind": "cross_family",
                "family_ids": ["fam-a", "fam-b"],
                "title": "跨族重复：a ⇄ b",
                "best_variant_id": "va",
                "candidates": [
                    candidate("va", "fam-a", "a", keep=True),
                    candidate("vb", "fam-b", "b"),
                ],
                "recommendation": {
                    "keep_variant_id": "va",
                    "strategy": "keep_best",
                    "original_variant_id": "va",
                    "unify_available": True,
                    "alternatives": {},
                },
            },
        ]
        window = VideoLibraryMainWindow()
        dialog = CleanupDialog(window, groups, 0.95)
        self.assertEqual(dialog.tree.columnCount(), 7)
        self.assertGreaterEqual(dialog.tree.columnWidth(0), 72)
        self.assertTrue(
            any(
                label.objectName() == "dialogTitle" and label.text() == "整理视频库"
                for label in dialog.findChildren(QLabel)
            )
        )
        self.assertEqual(dialog.tree.topLevelItem(0).text(0), "demo（2 个版本）")
        self.assertEqual(
            dialog.group_radios[0]["keep"].text(),
            "保留勾选版本，其他移入待清理",
        )
        self.assertEqual(dialog.group_radios[0]["skip"].text(), "暂不处理")
        self.assertTrue(dialog.group_radios[0]["skip"].isChecked())
        self.assertTrue(dialog.group_radios[0]["candidates"]["v-bad"].isEnabled())
        self.assertIsInstance(dialog.group_radios[0]["force"], QCheckBox)
        self.assertEqual(
            dialog.group_radios[0]["force"].text(),
            "人工确认：连锁定版本也移入待清理",
        )
        blocked_row = dialog.tree.topLevelItem(0).child(2)
        dialog._select_candidate_row(blocked_row, 1)
        self.assertTrue(dialog.group_radios[0]["candidates"]["v-bad"].isChecked())
        self.assertTrue(dialog.group_radios[0]["keep"].isChecked())
        dialog.group_radios[0]["candidates"]["v-small"].setChecked(True)
        dialog.group_radios[1]["candidates"]["vb"].setChecked(True)
        self.assertTrue(dialog.group_radios[0]["candidates"]["v-small"].isChecked())
        self.assertTrue(dialog.group_radios[1]["candidates"]["vb"].isChecked())
        dialog.group_radios[0]["candidates"]["v-best"].setChecked(True)
        dialog.group_radios[1]["candidates"]["va"].setChecked(True)
        decisions = dialog.decisions()

        within = next(item for item in decisions if item["kind"] == "within_family")
        self.assertEqual(within["family_id"], "fam-1")
        self.assertEqual(within["keep_variant_id"], "v-best")
        # the blocked candidate (音轨不一致) is never scheduled for removal
        self.assertEqual(within["remove_variant_ids"], ["v-small"])

        cross = next(item for item in decisions if item["kind"] == "cross_family")
        self.assertEqual(cross["merge_into_family_id"], "fam-a")
        self.assertEqual(set(cross["merge_family_ids"]), {"fam-a", "fam-b"})
        self.assertEqual(cross["remove_variant_ids"], ["vb"])
        dialog.close()
        window.close()

    def test_empty_pending_cleanup_is_clear_and_not_actionable(self) -> None:
        window = VideoLibraryMainWindow()
        window.project = SimpleNamespace(
            pending_cleanup=lambda: [],
            cleanup_pending_issues=lambda: [],
        )
        dialog = PendingCleanupDialog(window, window)
        self.assertEqual(dialog.summary.text(), "当前没有待清理文件。")
        self.assertFalse(dialog.restore_button.isEnabled())
        self.assertFalse(dialog.empty_button.isEnabled())
        dialog.close()
        window.close()

    def test_cleanup_dialog_locks_keep_action_when_only_one_safe_version(self) -> None:
        def candidate(variant_id: str, *, allowed: bool, reasons: list[str]) -> dict:
            return {
                "family_id": "fam-1",
                "family_name": "demo",
                "variant_id": variant_id,
                "label": variant_id,
                "profile": "original",
                "path": f"media/{variant_id}.mp4",
                "exists": True,
                "sha256": "ab" * 32,
                "width": 1280,
                "height": 720,
                "duration_sec": 2.0,
                "bitrate_kbps": 1200,
                "size_bytes": 1024,
                "video_codec": "h264",
                "audio_codec": "aac",
                "has_audio": True,
                "ssim_to_best": 1.0,
                "confidence": {"level": "high", "matched": True},
                "auto_allowed": allowed,
                "can_keep": True,
                "block_reasons": reasons,
            }

        groups = [
            {
                "kind": "within_family",
                "family_ids": ["fam-1"],
                "title": "demo（2 个版本）",
                "best_variant_id": "v-best",
                "safe_to_apply": False,
                "candidates": [
                    candidate("v-best", allowed=True, reasons=[]),
                    candidate("v-locked", allowed=False, reasons=["时长不一致"]),
                ],
                "recommendation": {
                    "keep_variant_id": "v-best",
                    "strategy": "keep_best",
                    "original_variant_id": "v-best",
                    "unify_available": False,
                    "alternatives": {},
                },
            }
        ]
        window = VideoLibraryMainWindow()
        dialog = CleanupDialog(window, groups, 0.95)
        self.assertTrue(dialog.group_radios[0]["keep"].isEnabled())
        self.assertTrue(dialog.group_radios[0]["skip"].isChecked())
        self.assertTrue(dialog.group_radios[0]["force"].isEnabled())
        self.assertEqual(dialog.decisions(), [])
        dialog.group_radios[0]["keep"].setChecked(True)
        dialog.group_radios[0]["force"].setChecked(True)
        self.assertTrue(dialog.group_radios[0]["keep"].isChecked())
        self.assertFalse(dialog.group_radios[0]["skip"].isChecked())
        dialog.group_radios[0]["unify"].setChecked(True)
        self.assertFalse(dialog.group_radios[0]["force"].isChecked())
        self.assertTrue(dialog.group_radios[0]["unify"].isChecked())
        dialog.group_radios[0]["force"].setChecked(True)
        forced = dialog.decisions()
        self.assertEqual(forced[0]["force_remove_variant_ids"], ["v-locked"])
        dialog.close()
        window.close()

    def test_cleanup_noop_explains_locked_candidates(self) -> None:
        groups = [
            {
                "kind": "within_family",
                "family_ids": ["fam-1"],
                "title": "demo（2 个版本）",
                "candidates": [
                    {"auto_allowed": True, "can_keep": True, "block_reasons": []},
                    {
                        "auto_allowed": False,
                        "can_keep": True,
                        "label": "v-locked",
                        "block_reasons": ["时长不一致"],
                    },
                ],
            }
        ]
        window = VideoLibraryMainWindow()
        accepted = object()
        with (
            patch("pptx_tools.video_manager_gui.CleanupDialog") as dialog_class,
            patch("pptx_tools.video_manager_gui.QMessageBox.information") as info,
        ):
            dialog_class.DialogCode.Accepted = accepted
            dialog_class.return_value.exec.return_value = accepted
            dialog_class.return_value.decisions.return_value = []
            window._cleanup_scan_finished(groups, 0.95)
        info.assert_called_once()
        self.assertIn("时长不一致", info.call_args.args[2])
        window.close()

    def test_shared_message_dialogs_use_the_canonical_type_scale(self) -> None:
        parent = VideoLibraryMainWindow()
        watermark = WatermarkStyledDialog(parent, "提示", "正文说明")
        compression = CompressionStyledDialog(
            parent,
            "确认",
            "正文说明",
            [("继续", "continue", True), ("取消", "cancel", False)],
        )
        for dialog in (watermark, compression):
            self.assertEqual(dialog.minimumWidth(), 440)
            self.assertIn("font-size: 16px", dialog.styleSheet())
            self.assertIn("font-size: 11px", dialog.styleSheet())
            self.assertTrue(
                all(
                    button.height() == 40 for button in dialog.findChildren(QPushButton)
                )
            )
            dialog.close()
        parent.close()

    def test_active_child_process_is_terminated_cooperatively(self) -> None:
        errors = []

        def run_child() -> None:
            try:
                run_process(
                    [sys.executable, "-c", "import time; time.sleep(30)"],
                    check=True,
                )
            except subprocess.CalledProcessError as exc:
                errors.append(exc)

        thread = threading.Thread(target=run_child)
        thread.start()
        time.sleep(0.2)
        terminate_active_processes(grace_seconds=0.1)
        thread.join(timeout=3)

        self.assertFalse(thread.is_alive())
        self.assertEqual(len(errors), 1)

    def test_process_termination_can_be_limited_to_one_worker_thread(self) -> None:
        errors: list[str] = []
        ready = threading.Barrier(3)
        thread_ids: dict[str, int] = {}

        def run_child(name: str) -> None:
            thread_ids[name] = threading.get_ident()
            ready.wait()
            try:
                run_process(
                    [sys.executable, "-c", "import time; time.sleep(30)"],
                    check=True,
                )
            except subprocess.CalledProcessError:
                errors.append(name)

        left = threading.Thread(target=run_child, args=("left",))
        right = threading.Thread(target=run_child, args=("right",))
        left.start()
        right.start()
        ready.wait()
        time.sleep(0.2)

        terminate_active_processes(
            grace_seconds=0.1,
            owner_thread_id=thread_ids["left"],
        )
        left.join(timeout=3)
        self.assertFalse(left.is_alive())
        self.assertTrue(right.is_alive())

        terminate_active_processes(grace_seconds=0.1)
        right.join(timeout=3)
        self.assertFalse(right.is_alive())
        self.assertEqual(sorted(errors), ["left", "right"])

    def test_compactor_run_process_is_registered_and_terminated(self) -> None:
        errors = []

        def run_child() -> None:
            try:
                compactor_run(
                    [sys.executable, "-c", "import time; time.sleep(30)"],
                    cancel_callback=lambda: False,
                )
            except subprocess.CalledProcessError as exc:
                errors.append(exc)

        thread = threading.Thread(target=run_child)
        thread.start()
        time.sleep(0.3)
        terminate_active_processes(grace_seconds=0.1)
        thread.join(timeout=3)

        self.assertFalse(thread.is_alive())
        self.assertEqual(len(errors), 1)

    @unittest.skipIf(os.name == "nt", "POSIX process-group regression")
    def test_compactor_cancel_terminates_child_process_group(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            ready = root / "ready"
            survived = root / "survived"
            cancelled = threading.Event()
            child_code = (
                "import pathlib,sys,time; time.sleep(1); "
                "pathlib.Path(sys.argv[1]).write_text('alive'); time.sleep(30)"
            )
            parent_code = (
                "import pathlib,subprocess,sys,time; "
                "subprocess.Popen([sys.executable,'-c',sys.argv[3],sys.argv[2]]); "
                "pathlib.Path(sys.argv[1]).write_text('ready'); time.sleep(30)"
            )

            def run_child() -> None:
                try:
                    compactor_run(
                        [
                            sys.executable,
                            "-c",
                            parent_code,
                            str(ready),
                            str(survived),
                            child_code,
                        ],
                        cancel_callback=cancelled.is_set,
                    )
                except CancelledError:
                    pass

            thread = threading.Thread(target=run_child)
            thread.start()
            deadline = time.monotonic() + 3
            while not ready.exists() and time.monotonic() < deadline:
                time.sleep(0.05)
            self.assertTrue(ready.exists())
            cancelled.set()
            thread.join(timeout=3)
            time.sleep(1.1)

            self.assertFalse(thread.is_alive())
            self.assertFalse(survived.exists())


class _FakeSettings:
    store: dict[str, object] = {}

    def __init__(self, *_args: object) -> None:
        pass

    def value(self, key: str, default: object = None) -> object:
        return self.store.get(key, default)

    def setValue(self, key: str, val: object) -> None:
        self.store[key] = val


class BackfillTierDialogTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        _FakeSettings.store.clear()
        self._settings_patch = patch(
            "pptx_tools.video_manager_gui.QSettings", _FakeSettings
        )
        self._thumb_patch = patch(
            "pptx_tools.video_manager_gui._set_video_thumbnail", lambda *a, **k: None
        )
        self._settings_patch.start()
        self._thumb_patch.start()

    def tearDown(self) -> None:
        self._thumb_patch.stop()
        self._settings_patch.stop()

    def _dialog(self):
        from pptx_tools.video_manager_gui import PptxUpgradeReviewDialog

        items = [
            {
                "media_path": "ppt/media/media1.mp4",
                "metadata": {"width": 320, "height": 240, "duration_sec": 2.0},
                "match_kind": "exact",
                "family_id": "fam1",
                "family_name": "族A",
                "target_source": "/tmp/a.mp4",
                "occurrences": [],
                "source": "/tmp/embedded.mp4",
                "sha256": "x",
            }
        ]
        families = [
            {
                "id": "fam1",
                "name": "族A",
                "source_path": "/tmp/a.mp4",
                "source_sha256": "y",
                "resolution": "1920×1080",
                "width": 1920,
                "height": 1080,
                "bitrate_kbps": 8000,
                "video_codec": "h264",
                "audio_codec": "aac",
                "suffix": ".mp4",
            }
        ]
        return PptxUpgradeReviewDialog(None, Path("/tmp/deck.pptx"), items, families)

    def test_default_tier_is_best_and_accessor(self) -> None:
        dialog = self._dialog()
        self.assertEqual(dialog.quality_tier(), "best")

    def test_settings_roundtrip_and_invalid_fallback(self) -> None:
        _FakeSettings.store["backfill/quality_tier"] = "balanced"
        self.assertEqual(self._dialog().quality_tier(), "balanced")
        _FakeSettings.store["backfill/quality_tier"] = "bogus"
        self.assertEqual(self._dialog().quality_tier(), "best")

    def test_tier_switch_persists_and_updates_plan_column(self) -> None:
        dialog = self._dialog()
        row = dialog.tree.topLevelItem(0)
        self.assertIn("原样嵌入", row.text(4))  # best：1080p h264 达标
        dialog.tier_combo.setCurrentIndex(dialog.tier_combo.findData("balanced"))
        self.assertIn("转码至 ≤720p", row.text(4))
        self.assertEqual(_FakeSettings.store["backfill/quality_tier"], "balanced")


class CenteredMessageHelperTest(unittest.TestCase):
    """The static QMessageBox helpers center the first dialog against not-yet-
    finalized parent geometry on macOS; the explicit helper must center against
    the real window frame and still return the clicked button."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_helper_centers_on_parent_window_frame(self) -> None:
        parent = QWidget()
        parent.setGeometry(120, 90, 960, 620)
        parent.show()
        QApplication.processEvents()
        captured: list[tuple[QMessageBox, tuple[int, int]]] = []

        def fake_exec(self):  # noqa: ANN001
            # capture the position the helper just moved this box to
            geo = self.geometry()
            captured.append((self, (geo.x(), geo.y())))
            return QMessageBox.StandardButton.Ok

        with patch.object(QMessageBox, "exec", fake_exec):
            result = _exec_centered_message(
                parent,
                QMessageBox.Icon.Information,
                "标题",
                "正文",
            )
        self.assertEqual(result, QMessageBox.StandardButton.Ok)
        self.assertTrue(captured)
        box, (moved_x, moved_y) = captured[0]
        center = parent.frameGeometry().center()
        size = box.sizeHint()
        expected_x = max(0, center.x() - size.width() // 2)
        expected_y = max(0, center.y() - size.height() // 2)
        self.assertLessEqual(abs(moved_x - expected_x), 2)
        self.assertLessEqual(abs(moved_y - expected_y), 2)
        box.close()
        parent.close()

    def test_format_position_handles_hours_minutes_and_zero(self) -> None:
        self.assertEqual(_format_position(0), "0:00")
        self.assertEqual(_format_position(65_000), "1:05")
        self.assertEqual(_format_position(3_661_000), "1:01:01")


class DetailSeekBarTest(unittest.TestCase):
    """The detail drawer video player must support seeking via a slider
    wired to durationChanged/positionChanged and sliderReleased."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def _window(self) -> VideoLibraryMainWindow:
        window = VideoLibraryMainWindow()
        window.show()
        QApplication.processEvents()
        return window

    def test_seek_bar_hidden_until_playback_and_resets_on_close(self) -> None:
        window = self._window()
        try:
            self.assertTrue(window.detail_seek_bar.isHidden())
            self.assertEqual(window.detail_seek_slider.maximum(), 0)
            window._detail_duration_changed(10_000)
            self.assertEqual(window.detail_seek_slider.maximum(), 10_000)
            self.assertEqual(window.detail_duration_label.text(), "0:10")
            self.assertTrue(window.detail_seek_slider.isEnabled())
            window._detail_position_changed(5_000)
            self.assertEqual(window.detail_seek_slider.value(), 5_000)
            self.assertEqual(window.detail_position_label.text(), "0:05")
            window._close_detail_drawer()
            self.assertTrue(window.detail_seek_bar.isHidden())
            self.assertEqual(window.detail_seek_slider.maximum(), 0)
            self.assertEqual(window.detail_position_label.text(), "0:00")
        finally:
            window.close()

    def test_slider_released_seeks_player_and_moved_updates_label(self) -> None:
        window = self._window()
        try:
            window._detail_duration_changed(20_000)
            with patch.object(window.detail_media_player, "setPosition") as seek:
                window.detail_seek_slider.setSliderPosition(12_000)
                window._detail_seek_moved(12_000)
                self.assertEqual(window.detail_position_label.text(), "0:12")
                self.assertFalse(seek.called)
                window._detail_seek_released()
                seek.assert_called_once_with(12_000)
            # positionChanged while dragging must not fight the user's drag
            window.detail_seek_slider.setSliderPosition(8_000)
            window.detail_seek_slider.setSliderDown(True)
            window._detail_position_changed(15_000)
            self.assertEqual(window.detail_seek_slider.value(), 8_000)
        finally:
            window.close()


class WatermarkRemoveSelectionSemanticsTest(unittest.TestCase):
    """移除选中 must target highlighted rows (falling back to checked set)
    so a file whose type-toggle was unchecked can still be removed."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def _window(self):
        window = MainWindow()
        window.show()
        QApplication.processEvents()
        return window

    def test_remove_uses_highlighted_rows_when_checked_set_empty(self) -> None:
        window = self._window()
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                a = Path(temp_dir) / "a.docx"
                b = Path(temp_dir) / "b.docx"
                a.touch()
                b.touch()
                window.set_files([a, b])
                a = a.resolve()
                b = b.resolve()
                # uncheck both via type-toggle; rows remain highlighted
                window.set_path_checked(a, False)
                window.set_path_checked(b, False)
                self.assertEqual(window.checked_paths, set())
                # remove must still work because a row is highlighted
                self.assertTrue(window.remove_button.isEnabled())
                window.remove_selected_files()
                self.assertEqual(len(window.input_paths), 1)
                self.assertNotIn(a, window.input_paths)
        finally:
            window.close()

    def test_remove_falls_back_to_checked_when_no_highlight(self) -> None:
        window = self._window()
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                a = Path(temp_dir) / "a.docx"
                b = Path(temp_dir) / "b.docx"
                a.touch()
                b.touch()
                window.set_files([a, b])
                a = a.resolve()
                b = b.resolve()
                window.set_path_checked(a, False)
                self.assertEqual(window.checked_paths, {b})
                # no current row -> falls back to checked set {b}
                window.file_list.setCurrentItem(None)
                window.remove_selected_files()
                self.assertEqual(window.input_paths, [a])
        finally:
            window.close()


class DocumentFormatCompressionGuiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def _new_window(self) -> CompressionMainWindow:
        window = CompressionMainWindow()
        window.archive_mode_select.setCurrentIndex(0)
        window.image_archive_mode_select.setCurrentIndex(0)
        window.overwrite_checkbox.setChecked(False)
        window.auto_optimize_checkbox.setChecked(False)
        return window

    def test_whitelist_accepts_document_formats(self) -> None:
        for suffix in (".docx", ".docm", ".pdf", ".xlsx", ".xlsm"):
            self.assertIn(suffix, DOCUMENT_INPUT_EXTENSIONS)
            self.assertIn(suffix, SUPPORTED_COMPACTOR_INPUT_EXTENSIONS)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            files = []
            for name in ("a.docx", "b.pdf", "c.xlsx"):
                path = root / name
                path.touch()
                files.append(path.resolve())
            window = self._new_window()
            try:
                window.set_files(files)
                self.assertEqual(set(window.input_paths), set(files))
            finally:
                window.close()

    def test_start_job_requires_target_for_documents(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            doc = (Path(temp_dir) / "a.docx").resolve()
            doc.touch()
            window = self._new_window()
            try:
                window.set_files([doc])
                window.target_input.clear()
                with patch.object(window, "show_dialog") as dialog:
                    window.start_job()
                dialog.assert_called_once_with(
                    window.text["document_target_required_title"],
                    window.text["document_target_required_body"],
                )
                self.assertFalse(window.is_running)
                self.assertIsNone(window.worker)
            finally:
                window.close()

    def test_start_job_rejects_none_image_profile_for_documents(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            doc = (Path(temp_dir) / "a.pdf").resolve()
            doc.touch()
            window = self._new_window()
            try:
                window.set_files([doc])
                window.target_input.setText("1")
                none_index = window.image_profile_select.findData("none")
                self.assertGreaterEqual(none_index, 0)
                window.image_profile_select.setCurrentIndex(none_index)
                with patch.object(window, "show_dialog") as dialog:
                    window.start_job()
                dialog.assert_called_once_with(
                    window.text["document_image_profile_none_title"],
                    window.text["document_image_profile_none_body"],
                )
                self.assertFalse(window.is_running)
                self.assertIsNone(window.worker)
            finally:
                window.close()

    def test_start_job_accepts_document_with_target(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            doc = (Path(temp_dir) / "a.xlsx").resolve()
            doc.touch()
            window = self._new_window()
            try:
                window.set_files([doc])
                window.target_input.setText("1")
                with (
                    patch.object(window, "show_dialog") as dialog,
                    patch("pptx_video_compactor_gui.CompressionWorker") as worker_class,
                ):
                    window.start_job()
                dialog.assert_not_called()
                self.assertTrue(window.is_running)
                args, _ = worker_class.call_args
                self.assertEqual(args[0], [doc])
                self.assertEqual(args[1], 1.0)
                window.worker_thread.quit()
                window.worker_thread.wait(5000)
                window.is_running = False
                window.worker = None
                window.worker_thread = None
            finally:
                window.close()

    def test_worker_consumes_normalized_document_result(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            doc = root / "a.docx"
            output = root / "a_compact.docx"
            report = root / "a_compact.report.json"
            doc.touch()
            output.write_bytes(b"compressed")
            report.write_text("{}", encoding="utf-8")
            completed = []
            worker = CompressionWorker(
                [doc],
                1.0,
                "high",
                "high",
                "zh",
                COMPRESSION_STRINGS["zh"],
            )
            worker.finished.connect(
                lambda results, failures, cancelled, stopped: completed.append(
                    (results, failures)
                )
            )
            normalized = {
                "input_pptx": doc,
                "output_pptx": output,
                "report_path": report,
                "skipped": False,
                "reason": "",
            }
            with patch(
                "pptx_video_compactor_gui.compact_input_path",
                return_value=normalized,
            ) as dispatch:
                worker.run()
            dispatch.assert_called_once()
            self.assertEqual(completed[0][1], [])
            self.assertEqual(len(completed[0][0]), 1)
            source, result_output, size, skipped, reason = completed[0][0][0]
            self.assertEqual(source, doc)
            self.assertEqual(result_output, output)
            self.assertEqual(size, len(b"compressed"))
            self.assertFalse(skipped)
            self.assertEqual(reason, "")


if __name__ == "__main__":
    unittest.main()
