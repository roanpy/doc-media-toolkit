from __future__ import annotations

import json
import logging
import shutil
import tempfile
from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import (
    QEvent,
    QObject,
    QSettings,
    QThread,
    QTimer,
    Qt,
    QUrl,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QBrush,
    QColor,
    QDesktopServices,
    QDragEnterEvent,
    QDragMoveEvent,
    QDropEvent,
    QPixmap,
)
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QSlider,
    QSplitter,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pptx_output_watermark.process_utils import terminate_active_processes
from pptx_tools.manager_i18n import tr
from pptx_tools.app_logging import configure_app_logging, log_directory
from pptx_tools.media_manager_ui import MEDIA_MANAGER_STYLESHEET, OperationWorker
from pptx_tools.ui_theme import (
    configure_ui_font,
    format_user_file_size,
    install_control_help,
)
from pptx_tools.video_library_health import (
    audit_video_project,
    prune_missing_output_records,
)
from pptx_tools.video_manager import (
    BACKFILL_QUALITY_TIERS,
    DEFAULT_BACKFILL_TIER,
    VIDEO_SUFFIXES,
    VideoProject,
    _normalized_video_name,
    create_video_thumbnail,
    normalize_library_category,
    plan_backfill_action,
)


LOGGER = logging.getLogger("pptx_tools.video_manager_gui")
REVIEW_TAGS_ROLE = int(Qt.ItemDataRole.UserRole) + 1


def _format_mb(size_bytes: int) -> str:
    return format_user_file_size(size_bytes)


def _exec_centered_message(
    parent: QWidget,
    icon: QMessageBox.Icon,
    title: str,
    text: str,
    buttons: QMessageBox.StandardButton = QMessageBox.StandardButton.Ok,
    default: QMessageBox.StandardButton = QMessageBox.StandardButton.NoButton,
) -> QMessageBox.StandardButton:
    """Show a QMessageBox explicitly centered on the parent window.

    Qt's static ``QMessageBox.information/warning/question`` helpers center the
    first box against not-yet-finalized parent geometry on macOS, causing the
    first dialog of a session to appear in the top-left corner. Constructing the
    box explicitly and centering it against ``parent.window().frameGeometry()``
    after ``ensurePolished()`` avoids that.
    """
    box = QMessageBox(parent)
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setText(text)
    box.setStandardButtons(buttons)
    if default != QMessageBox.StandardButton.NoButton:
        box.setDefaultButton(default)
    box.ensurePolished()
    window = parent.window() if parent is not None else None
    if window is not None and window.isVisible():
        center = window.frameGeometry().center()
        size = box.sizeHint()
        box.move(
            max(0, center.x() - size.width() // 2),
            max(0, center.y() - size.height() // 2),
        )
    return box.exec()


def _format_duration(seconds: float) -> str:
    total = max(0, round(seconds))
    hours, remainder = divmod(total, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    return f"{minutes}:{seconds:02d}"


def _format_position(milliseconds: int) -> str:
    total = max(0, round(milliseconds / 1000))
    hours, remainder = divmod(total, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    return f"{minutes}:{seconds:02d}"


def _format_ssim(value: float | None) -> str:
    return "—" if value is None else f"{value:.3f}"


def _display_video_name(value: str) -> str:
    name = value.strip()
    while Path(name).suffix.lower() in VIDEO_SUFFIXES:
        name = Path(name).stem.strip()
    return name


class ResponsiveVideoThumbnail(QLabel):
    clicked = Signal()

    def __init__(self, text: str = "") -> None:
        super().__init__(text)
        self._source_pixmap = QPixmap()

    def set_source_pixmap(self, pixmap: QPixmap, *, single_frame: bool = False) -> None:
        if single_frame:
            frame_width = max(1, pixmap.width() // 3)
            pixmap = pixmap.copy(frame_width, 0, frame_width, pixmap.height())
        self._source_pixmap = pixmap
        self.setText("")
        self._fit_pixmap()

    def clear_source(self, text: str) -> None:
        self._source_pixmap = QPixmap()
        self.setPixmap(QPixmap())
        self.setText(text)

    def resizeEvent(self, event: QEvent) -> None:
        super().resizeEvent(event)
        self._fit_pixmap()

    def mouseReleaseEvent(self, event: QEvent) -> None:
        super().mouseReleaseEvent(event)
        if not self._source_pixmap.isNull():
            self.clicked.emit()

    def _fit_pixmap(self) -> None:
        if self._source_pixmap.isNull():
            return
        size = self.contentsRect().size()
        if size.width() <= 0 or size.height() <= 0:
            return
        self.setPixmap(
            self._source_pixmap.scaled(
                size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )


def _set_video_thumbnail(
    label: QLabel, source: Path, digest: str, *, single_frame: bool = False
) -> None:
    cache = Path(tempfile.gettempdir()) / "pptx-tools-video-thumbnails"
    target = cache / f"{digest[:32]}.jpg"
    pixmap = QPixmap()
    if target.is_file() or create_video_thumbnail(source, target):
        pixmap = QPixmap(str(target))
    if not pixmap.isNull():
        if isinstance(label, ResponsiveVideoThumbnail):
            label.set_source_pixmap(pixmap, single_frame=single_frame)
        else:
            label.setPixmap(
                pixmap.scaled(
                    label.contentsRect().size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        return
    message = tr("无法生成封面，可点击“播放”核对完整视频")
    if isinstance(label, ResponsiveVideoThumbnail):
        label.clear_source(message)
    else:
        label.setPixmap(QPixmap())
        label.setText(message)


class VideoMatchDialog(QDialog):
    """Review one unresolved video against ranked library candidates."""

    def __init__(
        self,
        parent: QWidget,
        item: dict[str, Any],
        *,
        allow_new_family: bool,
        allow_remember: bool,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr("确认视频匹配"))
        self.setMinimumSize(760, 500)
        self.resize(920, 600)
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)
        self.item = item
        self.selected_family_id: str | None = None
        self.create_new_family = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        source = Path(item["source"])
        metadata = item.get("metadata") or {}
        title = QLabel(tr("确认视频匹配"))
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        subtitle = QLabel(
            f"{tr('待处理：')}{source.name}{tr('  ·  ')}"
            + f"{int(metadata.get('width') or 0)}×{int(metadata.get('height') or 0)}{tr('  ·  ')}"
            + f"{_format_duration(float(metadata.get('duration_sec') or 0))}"
        )
        subtitle.setObjectName("dialogSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        previews = QHBoxLayout()
        previews.setSpacing(14)
        source_panel = QVBoxLayout()
        source_panel.setSpacing(6)
        source_heading = QLabel(tr("待核对视频  ·  10% / 50% / 90% 取帧"))
        source_heading.setObjectName("previewHeading")
        self.source_preview = self._preview_label(tr("待匹配视频"))
        source_play = QPushButton(tr("播放待匹配视频"))
        source_play.clicked.connect(lambda: self._open_video(source))
        source_panel.addWidget(source_heading)
        source_panel.addWidget(self.source_preview, 1)
        source_panel.addWidget(source_play)

        candidate_panel = QVBoxLayout()
        candidate_panel.setSpacing(6)
        candidate_heading = QLabel(tr("候选高清源  ·  选择下方视频族后显示"))
        candidate_heading.setObjectName("previewHeading")
        self.candidate_preview = self._preview_label(tr("选择下方候选"))
        self.candidate_play = QPushButton(tr("播放候选高清源"))
        self.candidate_play.setEnabled(False)
        self.candidate_play.clicked.connect(self._open_selected_candidate)
        candidate_panel.addWidget(candidate_heading)
        candidate_panel.addWidget(self.candidate_preview, 1)
        candidate_panel.addWidget(self.candidate_play)
        previews.addLayout(source_panel, 1)
        previews.addLayout(candidate_panel, 1)
        layout.addLayout(previews)
        self._set_thumbnail(
            self.source_preview,
            source,
            str(item.get("sha256") or source.name),
        )

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [
                tr("候选视频族"),
                tr("相似度"),
                tr("分辨率"),
                tr("时长"),
                tr("画面差异"),
                tr("音频"),
                tr("判断"),
            ]
        )
        self.tree.setColumnWidth(0, 280)
        self.tree.setColumnWidth(1, 70)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 80)
        self.tree.setColumnWidth(4, 90)
        self.tree.setColumnWidth(5, 70)
        self.tree.itemSelectionChanged.connect(self._candidate_changed)
        self.tree.itemDoubleClicked.connect(lambda _item, _column: self.accept())
        layout.addWidget(self.tree, 1)
        for candidate in item.get("candidates", []):
            confidence = candidate.get("confidence") or {}
            row = QTreeWidgetItem(
                [
                    _display_video_name(candidate["family_name"]),
                    f"{float(candidate.get('score') or 0):.0f}",
                    f"{candidate['width']}×{candidate['height']}",
                    _format_duration(float(candidate.get("duration_sec") or 0)),
                    str(
                        confidence.get("frame_total_distance")
                        if confidence.get("frame_total_distance") is not None
                        else "—"
                    ),
                    tr("一致")
                    if confidence.get("audio_consistent")
                    else tr("不同/未知"),
                    tr("严格匹配") if candidate.get("strict_match") else tr("人工核对"),
                ]
            )
            row.setData(0, Qt.ItemDataRole.UserRole, candidate)
            self.tree.addTopLevelItem(row)
        if self.tree.topLevelItemCount():
            self.tree.setCurrentItem(self.tree.topLevelItem(0))

        self.remember_checkbox = QCheckBox(tr("记住此关联，后续同一压缩版本自动匹配"))
        self.remember_checkbox.setChecked(True)
        self.remember_checkbox.setVisible(allow_remember)
        layout.addWidget(self.remember_checkbox)

        buttons = QDialogButtonBox()
        self.link_button = buttons.addButton(
            tr("关联并继续"), QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.link_button.setObjectName("primaryAction")
        self.link_button.setEnabled(self.tree.topLevelItemCount() > 0)
        if allow_new_family:
            new_button = buttons.addButton(
                tr("新建视频族"), QDialogButtonBox.ButtonRole.ActionRole
            )
            new_button.clicked.connect(self._choose_new_family)
        skip_button = buttons.addButton(
            tr("跳过"), QDialogButtonBox.ButtonRole.RejectRole
        )
        self.link_button.clicked.connect(self.accept)
        skip_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

    @staticmethod
    def _preview_label(text: str) -> QLabel:
        label = ResponsiveVideoThumbnail(text)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setMinimumHeight(120)
        label.setMaximumHeight(170)
        label.setStyleSheet(
            "background:#0f1720;color:#94a3b8;"
            "border:1px solid #334155;border-radius:8px;"
        )
        return label

    def _set_thumbnail(self, label: QLabel, source: Path, digest: str) -> None:
        _set_video_thumbnail(label, source, digest)

    def _candidate_changed(self) -> None:
        row = self.tree.currentItem()
        candidate = row.data(0, Qt.ItemDataRole.UserRole) if row is not None else None
        self.selected_family_id = (
            str(candidate["family_id"]) if candidate is not None else None
        )
        if hasattr(self, "link_button"):
            self.link_button.setEnabled(candidate is not None)
        if candidate is not None:
            self.candidate_play.setEnabled(True)
            self._set_thumbnail(
                self.candidate_preview,
                Path(candidate["source_path"]),
                str(candidate["source_sha256"]),
            )

    def _open_selected_candidate(self) -> None:
        row = self.tree.currentItem()
        candidate = row.data(0, Qt.ItemDataRole.UserRole) if row is not None else None
        if candidate is not None:
            self._open_video(Path(candidate["source_path"]))

    def _open_video(self, path: Path) -> None:
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(
                self, tr("确认视频匹配"), f"{tr('无法打开视频：')}{path.name}"
            )

    def _choose_new_family(self) -> None:
        self.create_new_family = True
        self.accept()

    @property
    def remember(self) -> bool:
        return self.remember_checkbox.isChecked()


class PptxUpgradeReviewDialog(QDialog):
    """Review every embedded video before creating an upgraded PPTX."""

    def __init__(
        self,
        parent: QWidget,
        source_pptx: Path,
        items: list[dict[str, Any]],
        families: list[dict[str, Any]],
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr("确认 PPTX 高清回填"))
        self.setMinimumSize(1080, 620)
        self.resize(1220, 720)
        self.source_pptx = source_pptx
        self.items = items
        self.families = families
        self.plans: dict[str, dict[str, Any]] = {}
        self._loading = False
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        title = QLabel(tr("确认 PPTX 高清回填"))
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        intro = QLabel(
            f"{source_pptx.name}{tr(' · ')}{len(items)}{tr(' 个内嵌视频。')}"
            + tr("系统只给出建议；确认后另存新 PPTX，不覆盖原文件。")
        )
        intro.setObjectName("dialogSubtitle")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        counts = {
            kind: sum(item.get("match_kind") == kind for item in items)
            for kind in ("exact", "content", "unmatched")
        }
        summary = QLabel(
            f"{tr('精确匹配 ')}{counts['exact']}{tr(' · 内容匹配 ')}{counts['content']}{tr(' · ')}"
            + f"{tr('需人工确认 ')}{counts['unmatched']}{tr('。')}"
            + tr("“仅本次替换”不会把不同视频误记为同一族。")
        )
        summary.setObjectName("dialogSummary")
        summary.setWordWrap(True)
        layout.addWidget(summary)

        self.settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        tier_row = QHBoxLayout()
        tier_row.addWidget(QLabel(tr("回填质量")))
        self.tier_combo = QComboBox()
        for key, spec in BACKFILL_QUALITY_TIERS.items():
            self.tier_combo.addItem(
                f"{spec['label']}{tr('（≤')}{spec['max_height']}p"
                + (
                    f"{tr(' · ≤')}{spec['bitrate_kbps'] // 1000}Mbps"
                    if spec["bitrate_kbps"]
                    else ""
                )
                + tr("）"),
                key,
            )
        saved_tier = str(
            self.settings.value("backfill/quality_tier", DEFAULT_BACKFILL_TIER)
        )
        self.tier_combo.setCurrentIndex(max(0, self.tier_combo.findData(saved_tier)))
        self.tier_combo.currentIndexChanged.connect(self._tier_changed)
        tier_row.addWidget(self.tier_combo, 1)
        layout.addLayout(tier_row)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [
                tr("PPTX 视频"),
                tr("当前规格"),
                tr("匹配结果"),
                tr("引用"),
                tr("执行计划"),
            ]
        )
        self.tree.setColumnWidth(0, 150)
        self.tree.setColumnWidth(1, 120)
        self.tree.setColumnWidth(2, 150)
        self.tree.setColumnWidth(3, 45)
        self.tree.header().setStretchLastSection(True)
        self.tree.setMinimumWidth(560)
        self.tree.itemSelectionChanged.connect(self._selection_changed)
        splitter.addWidget(self.tree)

        detail = QWidget()
        detail.setMinimumWidth(480)
        detail_layout = QVBoxLayout(detail)
        self.detail_title = QLabel(tr("选择左侧视频查看匹配详情"))
        self.detail_title.setWordWrap(True)
        detail_layout.addWidget(self.detail_title)

        previews = QHBoxLayout()
        self.current_preview = VideoMatchDialog._preview_label(tr("PPTX 当前视频"))
        self.target_preview = VideoMatchDialog._preview_label(tr("目标高清源"))
        previews.addWidget(self.current_preview, 1)
        previews.addWidget(self.target_preview, 1)
        detail_layout.addLayout(previews)

        preview_actions = QHBoxLayout()
        self.current_play = QPushButton(tr("播放 PPTX 当前视频"))
        self.target_play = QPushButton(tr("播放目标高清源"))
        self.current_play.clicked.connect(self._play_current)
        self.target_play.clicked.connect(self._play_target)
        preview_actions.addWidget(self.current_play, 1)
        preview_actions.addWidget(self.target_play, 1)
        detail_layout.addLayout(preview_actions)

        family_row = QHBoxLayout()
        family_row.addWidget(QLabel(tr("目标视频族")))
        self.family_combo = QComboBox()
        self.family_combo.setEditable(True)
        self.family_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.family_combo.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.family_combo.addItem(tr("未选择"), None)
        for family in families:
            resolution = family.get("resolution") or ""
            label = (
                f"{family['name']}{tr(' · ')}{resolution}"
                if resolution
                else family["name"]
            )
            self.family_combo.addItem(label, family["id"])
        completer = self.family_combo.completer()
        if completer is not None:
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            completer.setFilterMode(Qt.MatchFlag.MatchContains)
            completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.family_combo.currentIndexChanged.connect(self._family_changed)
        family_row.addWidget(self.family_combo, 1)
        detail_layout.addLayout(family_row)

        action_box = QGroupBox(tr("本视频处理方式"))
        action_layout = QVBoxLayout(action_box)
        self.keep_radio = QRadioButton(tr("保持 PPTX 当前视频，不回填"))
        self.replace_radio = QRadioButton(tr("用所选高清源替换，仅本次输出使用"))
        self.remember_radio = QRadioButton(tr("确认是同一视频：替换并记住当前媒体哈希"))
        self.action_group = QButtonGroup(self)
        for button in (self.keep_radio, self.replace_radio, self.remember_radio):
            self.action_group.addButton(button)
            action_layout.addWidget(button)
            button.toggled.connect(self._action_changed)
        detail_layout.addWidget(action_box)

        self.explanation = QLabel()
        self.explanation.setWordWrap(True)
        self.explanation.setMinimumHeight(54)
        detail_layout.addWidget(self.explanation)
        detail_layout.addStretch(1)
        splitter.addWidget(detail)
        splitter.setStretchFactor(0, 6)
        splitter.setStretchFactor(1, 5)
        splitter.setSizes([660, 540])
        layout.addWidget(splitter, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText(tr("按清单另存回填"))
        buttons.button(QDialogButtonBox.StandardButton.Ok).setObjectName(
            "primaryAction"
        )
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("取消"))
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        for item in items:
            self._add_item(item)
        if self.tree.topLevelItemCount():
            self.tree.setCurrentItem(self.tree.topLevelItem(0))

    def _add_item(self, item: dict[str, Any]) -> None:
        metadata = item.get("metadata") or {}
        resolution = (
            f"{int(metadata.get('width') or 0)}×"
            + f"{int(metadata.get('height') or 0)}{tr(' · ')}"
            + f"{_format_duration(float(metadata.get('duration_sec') or 0))}"
        )
        kind = item.get("match_kind")
        target_available = bool(item.get("target_source"))
        if kind in {"exact", "content"} and not target_available:
            match = f"{tr('匹配但高清源不可用 · ')}{item.get('family_name') or tr('未知族')}"
        elif kind == "exact":
            match = f"{tr('精确匹配 · ')}{item.get('family_name') or tr('未知族')}"
        elif kind == "content":
            match = f"{tr('内容匹配 · ')}{item.get('family_name') or tr('未知族')}"
        else:
            first_candidate = next(iter(item.get("candidates") or []), None)
            match = (
                f"{tr('待确认 · 建议 ')}{first_candidate['family_name']}"
                if first_candidate
                else tr("未找到可靠候选")
            )
        suggested_family_id = next(
            (
                candidate.get("family_id")
                for candidate in item.get("candidates", [])
                if candidate.get("family_id")
            ),
            None,
        )
        planned_family_id = (
            item.get("family_id") if target_available else None
        ) or suggested_family_id
        if item.get("already_high_quality"):
            action = tr("保持当前（已是高清源）")
            plan_action = "keep"
        elif item.get("family_id") and target_available:
            if kind == "content":
                action = tr("内容匹配：回填并记住")
                plan_action = "remember"
            else:
                action = tr("回填建议高清源")
                plan_action = "replace"
        else:
            action = tr("保持当前（待人工选择）")
            plan_action = "keep"
        media_path = str(item["media_path"])
        self.plans[media_path] = {
            "action": plan_action,
            "family_id": planned_family_id,
        }
        row = QTreeWidgetItem(
            [
                Path(media_path).name,
                resolution,
                match,
                str(len(item.get("occurrences") or [])),
                action,
            ]
        )
        row.setData(0, Qt.ItemDataRole.UserRole, item)
        row.setToolTip(0, media_path)
        self.tree.addTopLevelItem(row)
        self._refresh_row(row)

    def _selection_changed(self) -> None:
        row = self.tree.currentItem()
        item = row.data(0, Qt.ItemDataRole.UserRole) if row is not None else None
        if not item:
            return
        self._loading = True
        plan = self.plans[str(item["media_path"])]
        source = Path(item["source"])
        metadata = item.get("metadata") or {}
        self.detail_title.setText(
            f"{Path(item['media_path']).name}{tr(' · ')}"
            + f"{int(metadata.get('width') or 0)}×"
            + f"{int(metadata.get('height') or 0)}{tr(' · ')}"
            + f"{_format_duration(float(metadata.get('duration_sec') or 0))}{tr(' · ')}"
            + f"{len(item.get('occurrences') or [])}{tr(' 处引用')}"
        )
        _set_video_thumbnail(
            self.current_preview, source, str(item.get("sha256") or source.name)
        )
        family_index = self.family_combo.findData(plan.get("family_id"))
        self.family_combo.setCurrentIndex(max(0, family_index))
        action = plan["action"]
        self.keep_radio.setChecked(action == "keep")
        self.replace_radio.setChecked(action == "replace")
        self.remember_radio.setChecked(action == "remember")
        self._loading = False
        self._refresh_target()
        self._refresh_explanation()

    def _family_changed(self) -> None:
        if self._loading:
            return
        row = self.tree.currentItem()
        if row is None:
            return
        item = row.data(0, Qt.ItemDataRole.UserRole)
        plan = self.plans[str(item["media_path"])]
        plan["family_id"] = self.family_combo.currentData()
        if plan["family_id"] and plan["action"] == "keep":
            self.replace_radio.setChecked(True)
        self._refresh_target()
        self._refresh_row()

    def _action_changed(self) -> None:
        if self._loading:
            return
        row = self.tree.currentItem()
        if row is None:
            return
        item = row.data(0, Qt.ItemDataRole.UserRole)
        plan = self.plans[str(item["media_path"])]
        if self.keep_radio.isChecked():
            plan["action"] = "keep"
        elif self.replace_radio.isChecked():
            plan["action"] = "replace"
        elif self.remember_radio.isChecked():
            plan["action"] = "remember"
        self._refresh_row()
        self._refresh_explanation()

    def quality_tier(self) -> str:
        tier = str(self.tier_combo.currentData() or DEFAULT_BACKFILL_TIER)
        return tier if tier in BACKFILL_QUALITY_TIERS else DEFAULT_BACKFILL_TIER

    def _tier_changed(self) -> None:
        self.settings.setValue("backfill/quality_tier", self.quality_tier())
        self._refresh_all_rows()

    def _refresh_all_rows(self) -> None:
        for index in range(self.tree.topLevelItemCount()):
            self._refresh_row(self.tree.topLevelItem(index))

    def _family_metadata(self, family: dict[str, Any]) -> dict[str, Any]:
        return {
            "suffix": family.get("suffix") or "",
            "video_codec": family.get("video_codec") or "",
            "audio_codec": family.get("audio_codec") or "",
            "width": int(family.get("width") or 0),
            "height": int(family.get("height") or 0),
            "bitrate_kbps": int(family.get("bitrate_kbps") or 0),
        }

    def _refresh_row(self, row: QTreeWidgetItem | None = None) -> None:
        row = row or self.tree.currentItem()
        if row is None:
            return
        item = row.data(0, Qt.ItemDataRole.UserRole)
        plan = self.plans[str(item["media_path"])]
        family = self._family_by_id(plan.get("family_id"))
        if plan["action"] == "keep":
            text = tr("保持当前")
        elif family is None:
            text = tr("请选择目标视频族")
        else:
            base = (
                f"{tr('关联并回填 · ')}{family['name']}"
                if plan["action"] == "remember"
                else f"{tr('仅本次回填 · ')}{family['name']}"
            )
            text = f"{base}{tr(' · ')}{plan_backfill_action(self._family_metadata(family), self.quality_tier())}"
        row.setText(4, text)

    def _refresh_target(self) -> None:
        family = self._family_by_id(self.family_combo.currentData())
        enabled = family is not None
        self.target_play.setEnabled(enabled)
        self.replace_radio.setEnabled(enabled)
        self.remember_radio.setEnabled(enabled)
        if family is None:
            self.target_preview.setPixmap(QPixmap())
            self.target_preview.setText(tr("选择目标视频族后显示高清源封面"))
            return
        _set_video_thumbnail(
            self.target_preview,
            Path(family["source_path"]),
            str(family["source_sha256"]),
        )

    def _refresh_explanation(self) -> None:
        if self.keep_radio.isChecked():
            text = tr("保持：输出 PPTX 中该视频完全不变，也不会新增关联。")
        elif self.remember_radio.isChecked():
            text = tr(
                "确认同一视频：成功输出并校验后，把当前媒体哈希记入该视频族；"
            ) + tr("以后同一压缩版可自动匹配。")
        else:
            text = tr(
                "仅本次替换：输出使用所选高清源，但不会把当前视频登记为同一内容，"
            ) + tr("适合临时指定或身份仍不确定的情况。")
        self.explanation.setText(text)

    def _family_by_id(self, family_id: str | None) -> dict[str, Any] | None:
        return next(
            (family for family in self.families if family["id"] == family_id), None
        )

    def _play_current(self) -> None:
        row = self.tree.currentItem()
        if row is not None:
            item = row.data(0, Qt.ItemDataRole.UserRole)
            self._open_video(Path(item["source"]))

    def _play_target(self) -> None:
        family = self._family_by_id(self.family_combo.currentData())
        if family is not None:
            self._open_video(Path(family["source_path"]))

    def _open_video(self, path: Path) -> None:
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(
                self, tr("确认 PPTX 高清回填"), f"{tr('无法打开视频：')}{path.name}"
            )

    def accept(self) -> None:
        invalid = [
            media_path
            for media_path, plan in self.plans.items()
            if plan["action"] != "keep" and not plan.get("family_id")
        ]
        if invalid:
            QMessageBox.warning(
                self,
                tr("确认 PPTX 高清回填"),
                tr("仍有计划回填但未选择目标视频族的项目，请先处理。"),
            )
            return
        super().accept()

    def decisions(self) -> tuple[dict[str, str], set[str], set[str]]:
        overrides: dict[str, str] = {}
        remembered: set[str] = set()
        kept: set[str] = set()
        for media_path, plan in self.plans.items():
            if plan["action"] == "keep":
                kept.add(media_path)
                continue
            overrides[media_path] = str(plan["family_id"])
            if plan["action"] == "remember":
                remembered.add(media_path)
        return overrides, remembered, kept


class CleanupDialog(QDialog):
    """Side-by-side evaluation of duplicate candidates with a recommended
    keep choice per group. The user confirms before anything is quarantined."""

    def __init__(
        self,
        parent: QWidget,
        groups: list[dict[str, Any]],
        ssim_threshold: float,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr("整理视频库"))
        self.setMinimumSize(760, 500)
        self.resize(920, 600)
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)
        self.groups = groups
        self.group_radios: list[dict[str, Any]] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        title = QLabel(tr("整理视频库"))
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        hint = QLabel(
            f"{tr('每组都是系统认为重复的视频。★ 是推荐保留项；画面相似度越接近 1 越一致')}"
            + f"{tr('（本次门槛 ')}{ssim_threshold:.2f}{tr('）。时长或音轨不一致的版本会被锁定，')}"
            + tr("不会自动清理。整理只把文件移到“待清理”，之后仍可恢复。")
        )
        hint.setObjectName("dialogSummary")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [
                tr("保留"),
                tr("视频 / 版本"),
                tr("分辨率"),
                tr("时长"),
                tr("大小"),
                tr("画面相似度"),
                tr("处理建议"),
            ]
        )
        # Leave enough room for the scaled indicator plus its column padding;
        # 54px let the indicator visually intrude into the video column on HiDPI.
        self.tree.setColumnWidth(0, 72)
        self.tree.setColumnWidth(1, 330)
        self.tree.setColumnWidth(2, 110)
        self.tree.setColumnWidth(3, 80)
        self.tree.setColumnWidth(4, 90)
        self.tree.setColumnWidth(5, 110)
        self.tree.header().setStretchLastSection(True)
        self.tree.itemClicked.connect(self._select_candidate_row)
        layout.addWidget(self.tree, 1)

        for group in groups:
            self._add_group(group)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText(tr("应用整理"))
        buttons.button(QDialogButtonBox.StandardButton.Ok).setObjectName(
            "primaryAction"
        )
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("取消"))
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _add_group(self, group: dict[str, Any]) -> None:
        rec = group["recommendation"]
        safe_to_apply = group.get("safe_to_apply", True)
        group_item = QTreeWidgetItem(self.tree, [group["title"], *([""] * 6)])
        group_item.setFirstColumnSpanned(True)
        group_item.setToolTip(0, group["title"])
        group_item.setForeground(0, QColor("#f8fafc"))
        group_item.setBackground(0, QBrush(QColor("#152235")))
        candidate_items: dict[str, QRadioButton] = {}
        candidate_group = QButtonGroup(self)
        candidate_group.setExclusive(True)
        for candidate in group["candidates"]:
            keep_radio = QRadioButton()
            candidate_group.addButton(keep_radio)
            keep_radio.setEnabled(candidate.get("can_keep", True))
            marker = "★ " if candidate["variant_id"] == rec["keep_variant_id"] else ""
            if candidate["variant_id"] == rec["keep_variant_id"]:
                keep_radio.setChecked(True)
            note = (
                tr("文件损坏或不完整，仅可清理")
                if candidate.get("integrity_error")
                else tr("；").join(candidate["block_reasons"])
            ) or (
                tr("推荐保留")
                if candidate["variant_id"] == rec["keep_variant_id"]
                else tr("可清理")
            )
            row = QTreeWidgetItem(
                group_item,
                [
                    "",
                    f"{marker}{candidate['family_name']}{tr(' · ')}{candidate['label']}",
                    f"{candidate['width']}×{candidate['height']}",
                    _format_duration(candidate["duration_sec"]),
                    _format_mb(candidate["size_bytes"]),
                    _format_ssim(candidate["ssim_to_best"]),
                    note,
                ],
            )
            row.setToolTip(
                1,
                tr("{}\n{}×{} · {}kbps · {} · 音轨：{}").format(
                    candidate["path"],
                    candidate["width"],
                    candidate["height"],
                    candidate["bitrate_kbps"],
                    candidate["video_codec"] or tr("未知编码"),
                    (
                        candidate["audio_codec"] or tr("有")
                        if candidate["has_audio"]
                        else tr("无")
                    ),
                ),
            )
            if not candidate.get("can_keep", True):
                for column in range(7):
                    row.setDisabled(True)
                keep_radio.setEnabled(False)
            self.tree.setItemWidget(row, 0, keep_radio)
            candidate_items[candidate["variant_id"]] = keep_radio
            row.setData(0, Qt.ItemDataRole.UserRole, candidate["variant_id"])
        group_item.setExpanded(True)

        action_panel = QWidget()
        action_panel.setObjectName("cleanupActionPanel")
        action_layout = QVBoxLayout(action_panel)
        action_layout.setContentsMargins(8, 3, 8, 3)
        action_layout.setSpacing(4)
        primary_actions = QHBoxLayout()
        primary_actions.setSpacing(12)
        action_label = QLabel(tr("处理这组："))
        action_label.setObjectName("cleanupActionLabel")
        action_label.setMinimumWidth(72)
        primary_actions.addWidget(action_label)
        unify_radio = QRadioButton(tr("生成兼容 1080p 后清理其他"))
        unify_radio.setEnabled(rec["unify_available"] and safe_to_apply)
        skip_radio = QRadioButton(tr("暂不处理"))
        keep_radio = QRadioButton(tr("保留勾选版本，其他移入待清理"))
        eligible_candidates = [
            candidate
            for candidate in group["candidates"]
            if candidate.get("auto_allowed", False) and candidate.get("can_keep", True)
        ]
        forced_candidates = [
            candidate
            for candidate in group["candidates"]
            if not candidate.get("auto_allowed", False)
            and candidate.get("can_keep", True)
        ]
        force_enabled = bool(
            group["kind"] == "within_family"
            and eligible_candidates
            and forced_candidates
        )
        keep_radio.setEnabled(len(eligible_candidates) >= 2 or force_enabled)
        if not keep_radio.isEnabled():
            keep_radio.setToolTip(
                tr("没有至少两个通过时长、音轨和内容一致性校验的版本；")
                + tr("被锁定的版本不会自动移入待清理。")
            )
        force_check = QCheckBox(tr("人工确认：连锁定版本也移入待清理"))
        force_check.setObjectName("cleanupForceCheck")
        force_check.setEnabled(force_enabled)
        force_visible = group["kind"] == "within_family" and bool(forced_candidates)
        force_check.setVisible(force_visible)
        force_check.setToolTip(
            tr("仅限族内整理；不会验证该版本与保留项完全一致，")
            + tr("确认后仍先移入可恢复的待清理目录。")
        )
        action_group = QButtonGroup(self)
        action_group.setExclusive(True)
        for action_radio in (keep_radio, unify_radio, skip_radio):
            action_group.addButton(action_radio)
        (keep_radio if safe_to_apply else skip_radio).setChecked(True)
        for candidate_radio in candidate_items.values():
            candidate_radio.toggled.connect(
                lambda checked, action=keep_radio: (
                    action.setChecked(True) if checked and action.isEnabled() else None
                )
            )
        primary_actions.addWidget(keep_radio)
        primary_actions.addWidget(unify_radio)
        primary_actions.addWidget(skip_radio)
        primary_actions.addStretch(1)
        action_layout.addLayout(primary_actions)
        if force_visible:
            force_panel = QWidget()
            force_panel.setObjectName("cleanupForcePanel")
            force_actions = QHBoxLayout(force_panel)
            force_actions.setContentsMargins(72, 5, 8, 5)
            force_actions.setSpacing(8)
            force_actions.addWidget(force_check)
            force_hint = QLabel(
                tr("仅族内可用；仍先移入可恢复的待清理目录")
                if force_enabled
                else tr("需要至少一个可安全保留的版本")
            )
            force_hint.setObjectName("cleanupForceHint")
            force_hint.setWordWrap(True)
            force_actions.addWidget(force_hint, 1)
            action_layout.addWidget(force_panel)
        action_item = QTreeWidgetItem(group_item, ["", *([""] * 6)])
        action_item.setFirstColumnSpanned(True)
        action_item.setSizeHint(0, action_panel.sizeHint())
        self.tree.setItemWidget(action_item, 0, action_panel)
        self.group_radios.append(
            {
                "group": group,
                "keep": keep_radio,
                "unify": unify_radio,
                "skip": skip_radio,
                "force": force_check,
                "candidates": candidate_items,
                "candidate_group": candidate_group,
                "action_group": action_group,
            }
        )
        force_check.toggled.connect(
            lambda checked, entry=self.group_radios[-1]: self._force_cleanup_toggled(
                entry, checked
            )
        )
        for action_radio in (keep_radio, unify_radio, skip_radio):
            action_radio.toggled.connect(
                lambda checked, action=action_radio, entry=self.group_radios[-1]: (
                    self._cleanup_action_toggled(entry, action, checked)
                )
            )

    def accept(self) -> None:
        for entry in self.group_radios:
            if not entry["force"].isChecked():
                continue
            keep_id = next(
                (
                    variant_id
                    for variant_id, radio in entry["candidates"].items()
                    if radio.isChecked()
                ),
                entry["group"]["recommendation"]["keep_variant_id"],
            )
            keep = next(
                candidate
                for candidate in entry["group"]["candidates"]
                if candidate["variant_id"] == keep_id
            )
            if not keep.get("auto_allowed", False):
                QMessageBox.warning(
                    self,
                    tr("无法强制整理"),
                    tr("强制整理必须保留已通过内容一致性校验的版本。"),
                )
                return
        super().accept()

    def _force_cleanup_toggled(self, entry: dict[str, Any], checked: bool) -> None:
        if not checked:
            return
        entry["skip"].setChecked(False)
        selected_id = next(
            (
                variant_id
                for variant_id, radio in entry["candidates"].items()
                if radio.isChecked()
            ),
            None,
        )
        selected = next(
            (
                candidate
                for candidate in entry["group"]["candidates"]
                if candidate["variant_id"] == selected_id
            ),
            None,
        )
        if selected is None or not selected.get("auto_allowed", False):
            safe_id = next(
                (
                    candidate["variant_id"]
                    for candidate in entry["group"]["candidates"]
                    if candidate.get("auto_allowed", False)
                    and candidate.get("can_keep", True)
                ),
                None,
            )
            if safe_id is not None:
                entry["candidates"][safe_id].setChecked(True)
        entry["keep"].setChecked(True)

    def _cleanup_action_toggled(
        self, entry: dict[str, Any], action: QRadioButton, checked: bool
    ) -> None:
        """Keep the manual force option as a modifier of the keep action only."""
        if checked and action in (entry["unify"], entry["skip"]):
            entry["force"].setChecked(False)

    def _select_candidate_row(self, item: QTreeWidgetItem, _column: int) -> None:
        variant_id = item.data(0, Qt.ItemDataRole.UserRole)
        if not variant_id:
            return
        for entry in self.group_radios:
            radio = entry["candidates"].get(variant_id)
            if radio is None or not radio.isEnabled():
                continue
            if entry["force"].isChecked():
                candidate = next(
                    candidate
                    for candidate in entry["group"]["candidates"]
                    if candidate["variant_id"] == variant_id
                )
                if not candidate.get("auto_allowed", False):
                    entry["force"].setChecked(False)
            radio.setChecked(True)
            if entry["keep"].isEnabled():
                entry["keep"].setChecked(True)
            return

    def decisions(self) -> list[dict[str, Any]]:
        decisions: list[dict[str, Any]] = []
        for entry in self.group_radios:
            if entry["skip"].isChecked() and not entry["force"].isChecked():
                continue
            group = entry["group"]
            keep_id = next(
                (
                    variant_id
                    for variant_id, radio in entry["candidates"].items()
                    if radio.isChecked()
                ),
                group["recommendation"]["keep_variant_id"],
            )
            force_ids = [
                candidate["variant_id"]
                for candidate in group["candidates"]
                if candidate["variant_id"] != keep_id
                and not candidate.get("auto_allowed", False)
                and candidate.get("can_keep", True)
            ]
            if not entry["force"].isChecked():
                force_ids = []
            remove_ids = [
                candidate["variant_id"]
                for candidate in group["candidates"]
                if candidate["variant_id"] != keep_id
                and (
                    candidate.get("auto_allowed", False)
                    or candidate["variant_id"] in force_ids
                )
            ]
            if not remove_ids and not entry["unify"].isChecked():
                continue
            decision: dict[str, Any] = {
                "kind": group["kind"],
                "keep_variant_id": keep_id,
                "remove_variant_ids": remove_ids,
                "force_remove_variant_ids": force_ids,
                "unify_first": entry["unify"].isChecked(),
                "reason": tr("整理视频库"),
            }
            if group["kind"] == "cross_family":
                keep_family = next(
                    candidate["family_id"]
                    for candidate in group["candidates"]
                    if candidate["variant_id"] == keep_id
                )
                decision["merge_into_family_id"] = keep_family
                decision["merge_family_ids"] = list(group["family_ids"])
            else:
                decision["family_id"] = group["family_ids"][0]
            decisions.append(decision)
        return decisions


class PendingCleanupDialog(QDialog):
    def __init__(self, parent: QWidget, window: MainWindow) -> None:
        super().__init__(parent)
        self.window = window
        self.setWindowTitle(tr("待清理目录"))
        self.setMinimumSize(720, 440)
        self.resize(860, 540)
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        title = QLabel(tr("待清理目录"))
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        self.summary = QLabel()
        self.summary.setObjectName("dialogSummary")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [tr("文件"), tr("来源视频族"), tr("大小"), tr("原因"), tr("隔离时间")]
        )
        self.tree.setColumnWidth(0, 240)
        self.tree.setColumnWidth(1, 140)
        self.tree.setColumnWidth(2, 80)
        self.tree.setColumnWidth(3, 120)
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        layout.addWidget(self.tree, 1)

        button_row = QHBoxLayout()
        self.restore_button = QPushButton(tr("还原选中"))
        self.restore_button.clicked.connect(self.restore_selected)
        self.empty_button = QPushButton(tr("清空待清理"))
        self.empty_button.setObjectName("dangerAction")
        self.empty_button.clicked.connect(self.empty_all)
        close_button = QPushButton(tr("关闭"))
        close_button.clicked.connect(self.accept)
        button_row.addWidget(self.restore_button)
        button_row.addWidget(self.empty_button)
        button_row.addStretch(1)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)
        self.reload()

    def reload(self) -> None:
        project = self.window.project
        self.tree.clear()
        if project is None:
            self.summary.setText(tr("尚未打开视频库。"))
            self.restore_button.setEnabled(False)
            self.empty_button.setEnabled(False)
            return
        entries = project.pending_cleanup()
        issues = project.cleanup_pending_issues()
        total = sum(entry["size_bytes"] for entry in entries)
        if not entries:
            self.summary.setText(tr("当前没有待清理文件。"))
        else:
            self.summary.setText(
                f"{tr('共 ')}{len(entries)}{tr(' 个文件，')}{_format_mb(total)}{tr('。')}"
                + (
                    f"\n⚠ {tr('；').join(issues)}"
                    if issues
                    else tr("\n迁移校验通过，可以清空。")
                )
            )
        for entry in entries:
            item = QTreeWidgetItem(
                self.tree,
                [
                    Path(entry["quarantined_path"]).name,
                    entry.get("family_name", ""),
                    _format_mb(entry["size_bytes"]),
                    entry.get("reason", ""),
                    entry.get("quarantined_at", ""),
                ],
            )
            item.setData(0, Qt.ItemDataRole.UserRole, entry["token"])
        self.empty_button.setEnabled(bool(entries) and not issues)
        self.restore_button.setEnabled(bool(entries))

    def restore_selected(self) -> None:
        item = self.tree.currentItem()
        if item is None or self.window.project is None:
            return
        token = item.data(0, Qt.ItemDataRole.UserRole)
        try:
            restored = self.window.project.restore_cleanup_entry(token)
        except Exception as exc:
            QMessageBox.warning(self, tr("待清理目录"), str(exc))
            return
        self.window.append_log(f"{tr('已还原：')}{restored}")
        self.window.refresh_views()
        self.reload()

    def empty_all(self) -> None:
        if self.window.project is None:
            return
        answer = QMessageBox.question(
            self,
            tr("清空待清理"),
            tr(
                "清空后文件将被永久删除，且 PPTX 哈希别名与视频族迁移已完成。是否继续？"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            removed = self.window.project.empty_cleanup()
        except Exception as exc:
            QMessageBox.warning(self, tr("待清理目录"), str(exc))
            return
        self.window.append_log(f"{tr('已清空 ')}{removed}{tr(' 个待清理文件。')}")
        self.window.refresh_views()
        self.reload()


class LibraryHealthDialog(QDialog):
    ISSUE_TITLES = {
        "invalid_family_pointer": tr("版本指针无效"),
        "missing_variant": tr("视频文件丢失"),
        "modified_variant": tr("视频文件被修改"),
        "hash_mismatch": tr("视频哈希不一致"),
        "variant_metadata_drift": tr("文件时间戳变化"),
        "unreadable_variant": tr("媒体不可读"),
        "ambiguous_known_hash": tr("哈希归属冲突"),
        "duplicate_variant_hash": tr("实体跨族重复"),
        "duplicate_variant_path": tr("路径重复引用"),
        "deck_primary_missing_alias_available": tr("PPTX 主路径失效"),
        "deck_source_missing": tr("PPTX 来源丢失"),
        "missing_output_record": tr("历史输出已删除"),
        "changed_output_record": tr("历史输出被修改"),
        "untracked_media": tr("未登记媒体文件"),
        "cleanup_index_issue": tr("待清理索引冲突"),
        "cleanup_index_invalid": tr("待清理索引损坏"),
    }

    def __init__(
        self,
        parent: QWidget,
        window: MainWindow,
        report: dict[str, Any],
    ) -> None:
        super().__init__(parent)
        self.setObjectName("libraryHealthDialog")
        self.window = window
        self.report = report
        self.setWindowTitle(tr("视频库体检"))
        self.setMinimumSize(760, 500)
        self.resize(900, 600)
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        title = QLabel(tr("视频库体检"))
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        self.summary = QLabel()
        self.summary.setObjectName("dialogSummary")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        self.tree = QTreeWidget()
        self.tree.setObjectName("healthTree")
        self.tree.setHeaderLabels(
            [tr("级别"), tr("检查项"), tr("数量"), tr("说明 / 对象")]
        )
        self.tree.setColumnWidth(0, 80)
        self.tree.setColumnWidth(1, 190)
        self.tree.setColumnWidth(2, 70)
        self.tree.header().setStretchLastSection(True)
        layout.addWidget(self.tree, 1)

        hint = QLabel(
            tr("“无关联”可能是独立素材，“多版本”可能是主动保留候选，")
            + tr("都不等于数据损坏。只有红色错误会阻断安全回填或清理。")
        )
        hint.setObjectName("healthHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QHBoxLayout()
        self.full_verify_button = QPushButton(tr("完整哈希复核"))
        self.full_verify_button.setToolTip(
            tr("逐个读取并计算全部库内视频哈希；可区分内容变化与时间戳变化，")
            + tr("耗时较长，但不会修改任何文件。")
        )
        self.full_verify_button.clicked.connect(self.run_full_verification)
        self.prune_button = QPushButton(tr("清理失效输出记录"))
        self.prune_button.setToolTip(
            tr("只删除指向已不存在 PPTX 的历史输出记录；")
            + tr("不会修改视频、视频族、PPTX 来源或形状关联。")
        )
        self.prune_button.clicked.connect(self.prune_stale_outputs)
        save_button = QPushButton(tr("保存 JSON 报告"))
        save_button.clicked.connect(self.save_report)
        close_button = QPushButton(tr("关闭"))
        close_button.clicked.connect(self.accept)
        buttons.addWidget(self.full_verify_button)
        buttons.addWidget(self.prune_button)
        buttons.addWidget(save_button)
        buttons.addStretch(1)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self.reload(report)

    def reload(self, report: dict[str, Any]) -> None:
        self.report = report
        stats = report["stats"]
        health = (
            tr("核心关联与媒体实体通过检查")
            if report["ok"]
            else f"{tr('发现 ')}{stats['errors']}{tr(' 个阻断性错误')}"
        )
        self.summary.setText(
            f"{health}{tr('。')}{stats['families']}{tr(' 个视频族 / ')}{stats['variants']}{tr(' 个版本 / ')}"
            + f"{stats['decks']}{tr(' 份 PPTX / ')}{stats['references']}{tr(' 处媒体引用；')}"
            + f"{tr('警告 ')}{stats['warnings']}{tr('，历史信息 ')}{stats['info']}{tr('。')}"
            + (
                f"{tr(' 本次为')}{tr('完整哈希') if report['mode'] == 'full_hash' else tr('快速')}{tr('检查。')}"
            )
        )
        self.tree.clear()
        grouped: dict[str, list[dict[str, Any]]] = {}
        for issue in report["issues"]:
            grouped.setdefault(issue["code"], []).append(issue)
        severity_order = {"error": 0, "warning": 1, "info": 2}
        severity_text = {"error": tr("错误"), "warning": tr("警告"), "info": tr("信息")}
        severity_color = {
            "error": QColor("#fb7185"),
            "warning": QColor("#fbbf24"),
            "info": QColor("#94a3b8"),
        }
        for code, issues in sorted(
            grouped.items(),
            key=lambda item: (
                severity_order[item[1][0]["severity"]],
                self.ISSUE_TITLES.get(item[0], item[0]),
            ),
        ):
            severity = issues[0]["severity"]
            group_item = QTreeWidgetItem(
                self.tree,
                [
                    severity_text[severity],
                    self.ISSUE_TITLES.get(code, code),
                    str(len(issues)),
                    issues[0]["message"],
                ],
            )
            for column in range(4):
                group_item.setForeground(column, severity_color[severity])
            for issue in issues[:50]:
                child = QTreeWidgetItem(
                    group_item,
                    [
                        "",
                        "",
                        "",
                        issue["message"],
                    ],
                )
                child.setToolTip(3, issue.get("path", issue["message"]))
            if len(issues) > 50:
                QTreeWidgetItem(
                    group_item,
                    [
                        "",
                        "",
                        "",
                        f"{tr('另有 ')}{len(issues) - 50}{tr(' 条，请保存 JSON 报告查看。')}",
                    ],
                )
        if not grouped:
            QTreeWidgetItem(
                self.tree,
                [tr("通过"), tr("未发现问题"), "0", tr("可以继续入库、清理和回填。")],
            )
        self.prune_button.setEnabled(stats["stale_output_records"] > 0)
        self.prune_button.setText(
            f"{tr('清理失效输出记录 (')}{stats['stale_output_records']})"
        )
        self.full_verify_button.setEnabled(report["mode"] != "full_hash")

    def run_full_verification(self) -> None:
        self.accept()
        self.window.run_library_health(verify_hashes=True)

    def prune_stale_outputs(self) -> None:
        project = self.window.project
        if project is None:
            return
        count = self.report["stats"]["stale_output_records"]
        answer = QMessageBox.question(
            self,
            tr("清理失效输出记录"),
            f"{tr('将删除 ')}{count}{tr(' 条指向已不存在 PPTX 的历史输出记录。')}\n\n"
            + tr("视频实体、视频族、PPTX 来源、哈希别名和形状关联不会改变。是否继续？"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            removed = prune_missing_output_records(project)
            report = audit_video_project(project)
        except Exception as exc:
            QMessageBox.warning(self, tr("清理失效输出记录"), str(exc))
            return
        self.window.append_log(f"{tr('已清理 ')}{removed}{tr(' 条失效历史输出记录。')}")
        self.window.refresh_views()
        self.reload(report)

    def save_report(self) -> None:
        project = self.window.project
        if project is None:
            return
        reports = project.root / "reports"
        reports.mkdir(parents=True, exist_ok=True)
        default = reports / f"library-health-{datetime.now():%Y%m%d-%H%M%S}.json"
        selected, _ = QFileDialog.getSaveFileName(
            self,
            tr("保存视频库体检报告"),
            str(default),
            "JSON (*.json)",
        )
        if not selected:
            return
        target = Path(selected)
        if target.suffix.lower() != ".json":
            target = target.with_suffix(".json")
        target.write_text(
            json.dumps(self.report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        self.window.append_log(f"{tr('已保存视频库体检报告：')}{target}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        app = QApplication.instance()
        if app is not None:
            configure_ui_font(app)
        super().__init__()
        configure_app_logging()
        self.setWindowTitle(tr("PPTX 视频资产库"))
        self.setAcceptDrops(True)
        self.setMinimumSize(880, 560)
        self.resize(960, 620)
        self.settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        self.project: VideoProject | None = None
        self.input_paths: list[Path] = []
        self.is_running = False
        self.worker_thread: QThread | None = None
        self.worker: OperationWorker | None = None
        self.operation_done: Callable[[Any], None] | None = None
        self.ai_thread: QThread | None = None
        self.ai_worker: OperationWorker | None = None
        self.ai_target_family_id = ""
        self.ai_ignore_result = False
        self.ai_item_names: dict[str, str] = {}
        self.ai_request_config = None
        self._build_ui()
        self._apply_style()
        install_control_help(self)
        last_project = self.settings.value("video_manager/last_project", "", str)
        if last_project and (Path(last_project) / "video-project.json").is_file():
            self.open_project(Path(last_project), report_errors=False)

    def _build_ui(self) -> None:
        central = QWidget()
        self.content_widget = central
        central.installEventFilter(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 6, 10, 6)
        root.setSpacing(6)

        header = QFrame()
        header.setObjectName("headerCard")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 8, 12, 8)
        title = QLabel(tr("PPTX 视频资产库"))
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            tr("管理高清源、版本与 PPTX 形状关联，为压缩文档安全回填可播放视频")
        )
        subtitle.setObjectName("pageSubtitle")
        title_stack = QVBoxLayout()
        title_stack.setContentsMargins(0, 0, 0, 0)
        title_stack.setSpacing(1)
        title_stack.addWidget(title)
        title_stack.addWidget(subtitle)
        header_layout.addLayout(title_stack, 1)
        self.help_button = QPushButton("?")
        self.help_button.setObjectName("helpIconButton")
        self.help_button.setToolTip(tr("使用说明"))
        self.help_button.setFixedSize(30, 30)
        self.help_button.clicked.connect(self.show_help)
        header_layout.addWidget(
            self.help_button,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        root.addWidget(header)

        project_frame = QFrame()
        project_frame.setObjectName("projectBar")
        project_bar = QHBoxLayout(project_frame)
        project_bar.setContentsMargins(12, 8, 8, 8)
        self.project_label = QLabel(tr("尚未打开视频库 · 可按主题新建多个独立库"))
        self.project_label.setObjectName("projectPath")
        project_bar.addWidget(self.project_label, 1)
        self.new_project_button = QPushButton(tr("新建视频库"))
        self.new_project_button.clicked.connect(self.choose_new_project)
        self.open_project_button = QPushButton(tr("切换 / 打开视频库"))
        self.open_project_button.setFixedHeight(32)
        self.open_project_button.setToolTip(
            tr("每个目录都是独立视频库；可按客户、主题或项目建立多个库。")
        )
        self.open_project_button.clicked.connect(self.choose_open_project)
        logs_button = QPushButton(tr("日志目录"))
        logs_button.clicked.connect(self.open_logs)
        self.log_panel_button = QPushButton(tr("操作记录"))
        self.log_panel_button.setCheckable(True)
        self.log_panel_button.setToolTip(
            tr("显示本次运行记录；长期滚动日志请点击“日志目录”。")
        )
        self.log_panel_button.toggled.connect(self.toggle_operation_log)
        self.health_button = QPushButton(tr("库体检"))
        self.health_button.setToolTip(
            tr("检查视频实体、哈希归属、PPTX 关联、历史输出和待清理索引；")
            + tr("默认只读，不会修改视频库。")
        )
        self.health_button.clicked.connect(self.run_library_health)
        project_bar.addWidget(self.open_project_button)
        self.project_menu_button = QPushButton(tr("更多"))
        self.project_menu_button.setFixedHeight(32)
        self.project_menu = QMenu(self.project_menu_button)
        self.project_menu_button.setMenu(self.project_menu)
        self.project_action_targets: list[tuple[Any, QPushButton]] = []
        for label, button in (
            (tr("新建视频库"), self.new_project_button),
            (tr("库体检"), self.health_button),
            (tr("操作记录"), self.log_panel_button),
            (tr("日志目录"), logs_button),
        ):
            action = self.project_menu.addAction(label)
            action.triggered.connect(button.click)
            self.project_action_targets.append((action, button))
        self.project_menu.aboutToShow.connect(self._sync_project_actions)
        project_bar.addWidget(self.project_menu_button)
        root.addWidget(project_frame)

        self.content_splitter = QSplitter(Qt.Orientation.Vertical)
        workflow_panel = self._build_workflow_panel()
        library_panel = self._build_library_panel()
        workflow_panel.setMinimumHeight(76)
        library_panel.setMinimumHeight(360)
        self.content_splitter.addWidget(workflow_panel)
        self.content_splitter.addWidget(library_panel)
        self.content_splitter.setStretchFactor(0, 0)
        self.content_splitter.setStretchFactor(1, 1)
        self.content_splitter.setSizes([84, 660])
        root.addWidget(self.content_splitter, 1)

        self.status_frame = QFrame()
        self.status_frame.setObjectName("jobStatus")
        status_row = QHBoxLayout(self.status_frame)
        status_row.setContentsMargins(0, 0, 0, 0)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.cancel_button = QPushButton(tr("停止"))
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.stop_job)
        status_row.addWidget(self.progress, 1)
        status_row.addWidget(self.cancel_button)
        self.status_frame.hide()
        root.addWidget(self.status_frame)

        self.log_shelf = QPushButton(tr("状态与日志 · 等待操作"))
        self.log_shelf.setObjectName("logShelf")
        self.log_shelf.setFixedHeight(26)
        self.log_shelf.clicked.connect(
            lambda: self.toggle_operation_log(self.log_output.isHidden())
        )
        root.addWidget(self.log_shelf)

        self.log_output = QPlainTextEdit(central)
        self.log_output.setObjectName("operationLog")
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText(tr("操作记录会显示在这里"))
        self.log_output.setMaximumBlockCount(1000)
        self.log_output.hide()
        self.log_hide_timer = QTimer(self)
        self.log_hide_timer.setSingleShot(True)
        self.log_hide_timer.timeout.connect(lambda: self.toggle_operation_log(False))
        self.setCentralWidget(central)
        self.refresh_views()

    def show_help(self) -> None:
        QMessageBox.information(
            self,
            tr("使用说明"),
            tr("选择或拖入 PPTX 后，可归档其中的视频并建立形状关联；")
            + tr("也可导入外部视频匹配，核对后再高清回填。"),
        )

    def _sync_project_actions(self) -> None:
        for action, button in self.project_action_targets:
            action.setEnabled(button.isEnabled())
            action.setCheckable(button.isCheckable())
            action.setChecked(button.isChecked())

    def _build_workflow_panel(self) -> QGroupBox:
        box = QGroupBox(tr("PPTX 工作流"))
        layout = QVBoxLayout(box)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        row1 = QHBoxLayout()
        row1.setSpacing(6)

        self.workflow_files = QFrame()
        self.workflow_files.setObjectName("workflowFiles")
        self.workflow_files.setAcceptDrops(True)
        self.workflow_files.installEventFilter(self)
        self.workflow_files_layout = QHBoxLayout(self.workflow_files)
        self.workflow_files_layout.setContentsMargins(6, 3, 6, 3)
        self.workflow_files_layout.setSpacing(6)
        self.input_summary = QLabel(tr("可拖入或多选 PPTX，批量归档或高清回填"))
        self.input_summary.setObjectName("inputSummary")
        self.input_summary.setAcceptDrops(True)
        self.input_summary.installEventFilter(self)
        self.input_summary.setToolTip(
            tr("视频源独立入库，PPTX 只记录形状关联。")
            + tr("匹配依据哈希和内容特征，改文件名或目录不受影响。")
        )
        self.workflow_files_layout.addWidget(self.input_summary, 1)
        self.workflow_chip_widgets: list[QWidget] = []
        row1.addWidget(self.workflow_files, 1)

        self.choose_input_button = QPushButton(tr("选择 PPTX（可多选）"))
        self.choose_input_button.setFixedHeight(32)
        self.choose_input_button.setToolTip(
            tr("可拖入或一次选择多个 PPTX，批量归档或高清回填。")
        )
        self.choose_input_button.clicked.connect(self.choose_input_pptx)
        row1.addWidget(self.choose_input_button)
        self.workflow_menu_button = QPushButton(tr("工作流设置"))
        self.workflow_menu_button.setFixedHeight(32)
        self.workflow_menu = QMenu(self.workflow_menu_button)
        self.workflow_menu.setToolTipsVisible(True)
        self.workflow_menu_button.setMenu(self.workflow_menu)
        row1.addWidget(self.workflow_menu_button)
        layout.addLayout(row1)

        self.workflow_settings = QFrame()
        self.workflow_settings.setObjectName("workflowSettings")
        self.workflow_settings.setMinimumHeight(38)
        action_row = QHBoxLayout(self.workflow_settings)
        action_row.setContentsMargins(8, 4, 8, 4)
        action_row.setSpacing(6)

        quality_label = QLabel(tr("入库质量"))
        quality_label.setObjectName("fieldLabel")
        action_row.addWidget(quality_label)

        self.source_quality_combo = QComboBox()
        self.source_quality_combo.addItem(tr("兼容 MP4（最高 1080p）"), "1080p")
        self.source_quality_combo.addItem(tr("兼容 MP4（保留分辨率）"), "mp4")
        self.source_quality_combo.addItem(tr("保留原片"), "original")
        self.source_quality_combo.setToolTip(
            tr("最高 1080p：超限时高质量降采样并转 MP4；")
            + tr("保留分辨率：WMV、M4V 等转为兼容 MP4，不缩放；")
            + tr("保留原片：库内文件字节不变。转换前后哈希仍关联到同一资源。")
        )
        self.source_quality_combo.setMinimumWidth(210)
        self.source_quality_combo.setMinimumHeight(26)
        saved_quality = self.settings.value(
            "compression/archive_source_quality", "1080p", str
        )
        self.source_quality_combo.setCurrentIndex(
            max(0, self.source_quality_combo.findData(saved_quality))
        )
        self.source_quality_combo.currentIndexChanged.connect(
            lambda _index: self.settings.setValue(
                "compression/archive_source_quality",
                self.source_quality_combo.currentData(),
            )
        )
        action_row.addWidget(self.source_quality_combo)

        category_label = QLabel(tr("分类目录"))
        category_label.setObjectName("fieldLabel")
        action_row.addWidget(category_label)

        self.category_input = QLineEdit()
        self.category_input.setPlaceholderText(tr("入库分类，例如：示例项目/2026"))
        self.category_input.setMinimumWidth(140)
        self.category_input.setMinimumHeight(26)
        self.category_input.setText(
            self.settings.value("compression/archive_category", "", str)
        )
        self.category_input.textChanged.connect(
            lambda value: self.settings.setValue("compression/archive_category", value)
        )
        action_row.addWidget(self.category_input, 1)

        self.archive_button = QPushButton(tr("归档 PPTX 视频"))
        self.archive_button.setObjectName("primaryAction")
        self.archive_button.setToolTip(
            tr("提取所选 PPTX 的内嵌视频并建立形状关联；不修改 PPTX。")
        )
        self.archive_button.clicked.connect(self.archive_pptx_videos)
        self.external_import_button = QPushButton(tr("导入外部视频并匹配"))
        self.external_import_button.setToolTip(
            tr("按哈希、时长、画面和音频指纹匹配；无匹配时新建视频，歧义时跳过。")
        )
        self.external_import_button.clicked.connect(self.import_external_videos)

        self.upgrade_button = QPushButton(tr("高清回填 PPTX（另存）"))
        self.upgrade_button.setToolTip(
            tr(
                "精确匹配后只替换包内视频；始终由你指定新文件或输出目录，不覆盖输入 PPTX。"
            )
        )
        self.upgrade_button.clicked.connect(self.upgrade_pptx)
        self.workflow_action_targets: list[tuple[Any, QPushButton]] = []
        for label, button in (
            (tr("归档 PPTX 视频"), self.archive_button),
            (tr("导入外部视频并匹配"), self.external_import_button),
            (tr("高清回填 PPTX（另存）"), self.upgrade_button),
        ):
            action = self.workflow_menu.addAction(label)
            action.triggered.connect(button.click)
            self.workflow_action_targets.append((action, button))
        self.workflow_menu.addSeparator()
        settings_action = self.workflow_menu.addAction(tr("展开 / 收起入库设置"))
        settings_action.triggered.connect(self._toggle_workflow_settings)
        self.workflow_menu.aboutToShow.connect(self._sync_workflow_actions)
        self.workflow_settings.hide()

        layout.addWidget(self.workflow_settings)
        return box

    def _sync_workflow_actions(self) -> None:
        for action, button in self.workflow_action_targets:
            action.setEnabled(button.isEnabled())
            action.setToolTip(button.toolTip())

    def _toggle_workflow_settings(self) -> None:
        visible = self.workflow_settings.isHidden()
        self.workflow_settings.setVisible(visible)
        if not hasattr(self, "content_splitter"):
            return
        sizes = self.content_splitter.sizes()
        workflow_height = 124 if visible else 84
        self.content_splitter.setSizes(
            [workflow_height, max(360, sum(sizes) - workflow_height)]
        )

    def _build_library_panel(self) -> QGroupBox:
        box = QGroupBox(tr("视频源与已知版本"))
        layout = QVBoxLayout(box)
        self.library_empty = QLabel(
            tr("先新建或打开一个视频库\n\n可从 PPTX 提取，或导入外部视频自动匹配")
        )
        self.library_empty.setObjectName("emptyState")
        self.library_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.library_empty.setWordWrap(True)
        self.video_tree = QTreeWidget()
        self.video_tree.setHeaderLabels(
            [
                tr("视频族 / 版本"),
                tr("分辨率"),
                tr("时长"),
                tr("大小"),
                tr("哈希数"),
                tr("关联状态"),
                tr("文件位置"),
            ]
        )
        self.video_tree.setColumnWidth(0, 280)
        self.video_tree.setColumnWidth(1, 110)
        self.video_tree.setColumnWidth(2, 85)
        self.video_tree.setColumnWidth(3, 85)
        self.video_tree.setColumnWidth(4, 80)
        self.video_tree.setColumnWidth(5, 190)
        self.video_tree.setAlternatingRowColors(False)
        self.video_tree.setAcceptDrops(True)
        self.video_tree.installEventFilter(self)
        self.video_tree_viewport = self.video_tree.viewport()
        self.video_tree_viewport.setAcceptDrops(True)
        self.video_tree_viewport.installEventFilter(self)
        self.video_tree.setToolTip(
            tr("可把视频拖到这里自动匹配；拖入不会修改已有视频族或 PPTX 关联。")
        )
        self.video_tree.itemSelectionChanged.connect(self._update_action_states)
        self.video_tree.itemSelectionChanged.connect(self._update_detail_drawer)
        self.video_tree.itemDoubleClicked.connect(
            lambda _item, _column: self.preview_selected()
        )
        self.video_tree.header().setSectionsClickable(True)
        self.video_tree.header().setSortIndicatorShown(True)
        self.video_tree.header().sectionClicked.connect(self._sort_by_header)
        saved_sort = self.settings.value("video_library/sort", "name", str)
        if saved_sort == "quality":
            saved_sort = "review"
        self.library_sort_mode = (
            saved_sort
            if saved_sort
            in {
                "name",
                "duration",
                "resolution",
                "size",
                "hashes",
                "review",
                "path",
            }
            else "name"
        )
        self.library_sort_descending = self.settings.value(
            "video_library/sort_descending",
            self.library_sort_mode != "name",
            bool,
        )
        self.video_tree.header().setSortIndicator(
            {
                "name": 0,
                "resolution": 1,
                "duration": 2,
                "size": 3,
                "hashes": 4,
                "review": 5,
                "path": 6,
            }.get(self.library_sort_mode, 0),
            (
                Qt.SortOrder.DescendingOrder
                if self.library_sort_descending
                else Qt.SortOrder.AscendingOrder
            ),
        )

        self.library_actions = QWidget()
        self.library_actions.setObjectName("libraryActions")
        self.library_actions.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        actions_layout = QGridLayout(self.library_actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)

        self.library_action_row = QWidget()
        version_row = QHBoxLayout(self.library_action_row)
        version_row.setContentsMargins(0, 0, 0, 0)
        version_row.setSpacing(6)
        self.preview_button = QPushButton(tr("播放"))
        self.preview_button.setToolTip(tr("在右侧详情面板播放；双击列表项也可播放。"))
        self.preview_button.clicked.connect(self.preview_selected)
        self.rename_button = QPushButton(tr("重命名"))
        self.rename_button.clicked.connect(self.rename_selected)
        self.rename_button.setToolTip(
            tr("选中视频族时同步重命名高清源；选中具体版本时只修改该文件名。")
        )
        self.move_button = QPushButton(tr("移动文件"))
        self.move_button.clicked.connect(self.move_selected_variant)
        self.move_button.setToolTip(
            tr("选中视频族时移动全部版本并更新库内分类；选中具体版本时只移动该文件。")
        )
        self.import_button = QPushButton(tr("添加版本"))
        self.import_button.clicked.connect(self.import_version)
        self.import_button.setToolTip(
            tr("把外部高清、重编码或改名视频加入当前视频族；")
            + tr("会先核对时长、画面和音频指纹，不安全时拒绝。")
        )
        self.activate_button = QPushButton(tr("设为高清源"))
        self.activate_button.clicked.connect(self.set_selected_source_variant)
        self.activate_button.setToolTip(
            tr("将选中版本设为以后高清回填 PPTX 使用的权威源。")
        )
        self.review_button = QPushButton(tr("核实版本"))
        self.review_button.setObjectName("primaryAction")
        self.review_button.clicked.connect(self.review_selected_family)
        self.review_button.setToolTip(
            tr("只核实当前视频族及其跨族重复候选；比较画面、时长、音轨和分辨率，")
            + tr("确认后才会把冗余文件移到待清理。")
        )
        self.quarantine_button = QPushButton(tr("隔离异常"))
        self.quarantine_button.setToolTip(
            tr("仅允许隔离未被引用、且不是高清源或当前版本的异常文件。")
        )
        self.quarantine_button.clicked.connect(self.quarantine_selected_abnormal)
        self.merge_button = QPushButton(tr("归并视频"))
        self.merge_button.clicked.connect(self.merge_selected_family)
        self.merge_button.setToolTip(
            tr("分析所选视频并按相似度显示候选封面；确认后合并视频族及全部 PPTX 关联。")
        )
        self.ai_button = QPushButton(tr("AI 整理建议"))
        self.ai_button.clicked.connect(self.request_ai_suggestion)
        self.ai_button.setToolTip(
            tr("先用代码规则筛选候选，再由 AI 建议同源归并、主视频和命名；")
            + tr("视觉模型可额外参考三帧联系图，不会自动修改视频库。")
        )
        self.relink_button = QPushButton(tr("核验 / 重新关联"))
        self.relink_button.setToolTip(
            tr(
                "跨机器复制后按 SHA-256 核验时间戳变化；也可按哈希重新定位缺失视频并恢复路径关联。"
            )
        )
        self.relink_button.clicked.connect(self.relink_missing)
        self.cleanup_button = QPushButton(tr("整理视频库"))
        self.cleanup_button.setToolTip(
            tr("扫描库内重复版本：并排评估 SSIM、时长、音轨与分辨率，")
            + tr("由你确认保留项后把冗余文件移到待清理目录。")
        )
        self.cleanup_button.clicked.connect(self.start_cleanup_scan)
        self.pending_button = QPushButton(tr("待清理 (0)"))
        self.pending_button.setToolTip(
            tr("查看整理后移入待清理目录的文件，可继续核对或恢复。")
        )
        self.pending_button.clicked.connect(self.show_pending_cleanup)

        self.more_actions_button = QPushButton(tr("更多操作"))
        self.more_actions_button.setToolTip(
            tr("收纳归并、文件恢复、异常处理和导入导出等低频操作。")
        )
        self.more_actions_menu = QMenu(self.more_actions_button)
        self.more_actions_menu.setToolTipsVisible(True)
        self.more_action_targets: list[tuple[Any, QPushButton]] = []

        self.more_actions_menu.addSection(tr("核对与归并"))
        action = self.more_actions_menu.addAction(tr("AI 整理建议"))
        action.triggered.connect(self.ai_button.click)
        self.more_action_targets.append((action, self.ai_button))
        action = self.more_actions_menu.addAction(tr("归并视频"))
        action.triggered.connect(self.merge_button.click)
        self.more_action_targets.append((action, self.merge_button))

        self.more_actions_menu.addSection(tr("版本管理"))
        action = self.more_actions_menu.addAction(tr("设为高清源"))
        action.triggered.connect(self.activate_button.click)
        self.more_action_targets.append((action, self.activate_button))

        self.more_actions_menu.addSection(tr("文件与恢复"))
        for label, button in (
            (tr("重命名"), self.rename_button),
            (tr("移动文件"), self.move_button),
            (tr("核验 / 重新关联"), self.relink_button),
        ):
            action = self.more_actions_menu.addAction(label)
            action.triggered.connect(button.click)
            self.more_action_targets.append((action, button))

        self.more_actions_menu.addSection(tr("异常处理"))
        action = self.more_actions_menu.addAction(tr("隔离异常"))
        action.triggered.connect(self.quarantine_button.click)
        self.more_action_targets.append((action, self.quarantine_button))

        self.more_actions_menu.addSection(tr("库维护"))
        for label, button in (
            (tr("整理视频库"), self.cleanup_button),
            (tr("待清理"), self.pending_button),
        ):
            action = self.more_actions_menu.addAction(label)
            action.triggered.connect(button.click)
            self.more_action_targets.append((action, button))

        self.more_actions_menu.addSection(tr("导出"))
        self.import_hash_catalog_action = self.more_actions_menu.addAction(
            tr("导入哈希目录")
        )
        self.import_hash_catalog_action.triggered.connect(self.import_hash_catalog)
        self.export_hash_catalog_action = self.more_actions_menu.addAction(
            tr("导出哈希目录")
        )
        self.export_hash_catalog_action.triggered.connect(self.export_hash_catalog)
        self.export_associations_action = self.more_actions_menu.addAction(
            tr("导出关联记录")
        )
        self.export_associations_action.triggered.connect(
            self.export_association_records
        )

        self.more_actions_menu.aboutToShow.connect(self._sync_more_actions)
        self.more_actions_button.setMenu(self.more_actions_menu)
        version_row.addWidget(self.review_button)
        version_row.addWidget(self.import_button)
        version_row.addWidget(self.activate_button)
        version_row.addWidget(self.cleanup_button)
        version_row.addWidget(self.pending_button)
        version_row.addWidget(self.more_actions_button)

        self.library_filter_row = QWidget()
        filter_row = QHBoxLayout()
        filter_row.setContentsMargins(0, 0, 0, 0)
        self.library_filter_row.setLayout(filter_row)
        self.library_search_label = QLabel(tr("查找"))
        filter_row.addWidget(self.library_search_label)
        self.library_filter_input = QLineEdit()
        self.library_filter_input.setPlaceholderText(tr("筛选名称、路径或哈希"))
        self.library_filter_input.setClearButtonEnabled(True)
        self.library_filter_input.setToolTip(
            tr("即时筛选视频族及版本；可输入名称、目录、分辨率或哈希片段。")
        )
        self.library_filter_input.setMinimumWidth(220)
        self.library_filter_input.setMaximumWidth(320)
        self.library_filter_input.textChanged.connect(self._apply_library_filter)
        filter_row.addWidget(self.library_filter_input)
        self.attention_filter_combo = QComboBox()
        self.attention_filter_combo.addItem(tr("全部视频"), "all")
        self.attention_filter_combo.addItem(tr("待核对"), "review")
        self.attention_filter_combo.addItem(tr("无 PPTX 关联"), "unlinked")
        self.attention_filter_combo.addItem(tr("多版本"), "multi")
        self.attention_filter_combo.addItem(tr("文件异常"), "abnormal")
        self.attention_filter_combo.setToolTip(
            tr("无关联可能是手工导入素材，多版本可能是主动保留候选；")
            + tr("筛选只用于核对，不会修改视频库。")
        )
        self.attention_filter_combo.setMinimumWidth(120)
        self.attention_filter_combo.setMaximumWidth(160)
        self.attention_filter_combo.currentIndexChanged.connect(
            self._apply_library_filter
        )
        self.attention_filter_combo.hide()
        filter_row.addWidget(self.attention_filter_combo)

        filter_row.addSpacing(8)
        self.library_stats_group = QButtonGroup(self)
        self.library_stats_group.setExclusive(True)
        self.library_stat_buttons: dict[str, QPushButton] = {}
        stat_definitions = (
            ("all", tr("全部"), tr("显示全部视频族；选择具体版本后可播放或设为高清源")),
            (
                "review",
                tr("待核对"),
                tr("无关联、多版本或文件异常的视频族去重并集；再按子类筛选处理"),
            ),
            (
                "unlinked",
                tr("无关联"),
                tr("没有 PPTX 引用的视频族；独立素材可保留，重复内容用“归并视频”"),
            ),
            (
                "multi",
                tr("多版本"),
                tr("包含多个清晰度或编码版本的视频族；用“整理视频库”安全保留最佳版本"),
            ),
            (
                "abnormal",
                tr("文件异常"),
                tr("库内文件缺失、不可读或探测失败；选择异常版本后用“隔离异常”"),
            ),
        )
        for key, title, tooltip in stat_definitions:
            button = QPushButton(f"{title} 0")
            button.setObjectName("statFilter")
            button.setCheckable(True)
            button.setToolTip(f"{tooltip}{tr('。点击只筛选列表，不会修改视频库。')}")
            button.clicked.connect(
                lambda _checked=False, name=key: self._filter_library_by_stat(name)
            )
            self.library_stats_group.addButton(button)
            self.library_stat_buttons[key] = button
            filter_row.addWidget(button)
        self.library_stat_buttons["all"].setChecked(True)

        self.library_version_label = QLabel(tr("0 版本"))
        self.library_version_label.setObjectName("statMeta")
        self.library_version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        filter_row.addWidget(self.library_version_label)
        self.library_stats_label = QLabel(tr("显示 0/0"))
        self.library_stats_label.setObjectName("statMeta")
        self.library_stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        filter_row.addWidget(self.library_stats_label)
        actions_layout.addWidget(self.library_filter_row, 0, 0)
        actions_layout.addWidget(self.library_action_row, 0, 1)
        actions_layout.setColumnStretch(0, 1)
        self.library_actions_layout = actions_layout
        layout.addWidget(self.library_actions)
        layout.addWidget(self.library_empty, 1)
        layout.addWidget(self.video_tree, 1)
        self.library_panel = box
        self.library_panel.installEventFilter(self)
        self._build_detail_drawer()

        self.project_action_buttons = [
            self.archive_button,
            self.external_import_button,
            self.upgrade_button,
            self.relink_button,
            self.cleanup_button,
            self.pending_button,
        ]
        self.selection_action_buttons = [
            self.preview_button,
            self.rename_button,
            self.move_button,
            self.import_button,
            self.activate_button,
            self.review_button,
            self.ai_button,
            self.quarantine_button,
            self.merge_button,
        ]
        return box

    def _sync_more_actions(self) -> None:
        for action, button in self.more_action_targets:
            action.setEnabled(button.isEnabled())
            action.setToolTip(button.toolTip())
            if button is self.pending_button:
                action.setText(button.text())
        self.export_associations_action.setEnabled(self.project is not None)
        self.import_hash_catalog_action.setEnabled(self.project is not None)
        self.export_hash_catalog_action.setEnabled(self.project is not None)

    def export_hash_catalog(self) -> None:
        if not self.require_project():
            return
        assert self.project is not None
        default = self.project.root / "video-hash-catalog.json"
        output, _ = QFileDialog.getSaveFileName(
            self, tr("导出视频哈希目录"), str(default), "JSON (*.json)"
        )
        if not output:
            return
        try:
            Path(output).write_text(
                json.dumps(self.project.hash_catalog(), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError as exc:
            self.show_error(f"{tr('无法导出哈希目录：')}{exc}")
            return
        self.append_log(f"{tr('已导出视频哈希目录：')}{output}")

    def import_hash_catalog(self) -> None:
        if not self.require_project():
            return
        assert self.project is not None
        source, _ = QFileDialog.getOpenFileName(
            self, tr("导入视频哈希目录"), str(self.project.root), "JSON (*.json)"
        )
        if not source:
            return
        try:
            report = self.project.merge_hash_catalog(
                json.loads(Path(source).read_text(encoding="utf-8"))
            )
        except (OSError, ValueError, json.JSONDecodeError, RuntimeError) as exc:
            self.show_error(f"{tr('无法导入哈希目录：')}{exc}")
            return
        self.refresh_views()
        QMessageBox.information(
            self,
            tr("哈希目录已合并"),
            (
                f"{tr('匹配视频族 ')}{report['matched']}{tr('，新增哈希 ')}{report['added']}{tr('；')}"
                + f"{tr('无本地锚点 ')}{report['skipped']}{tr('，冲突 ')}{report['conflicts']}{tr('。')}\n"
                + tr("未导入媒体文件，也未自动归并视频族。")
            ),
        )

    def export_association_records(self) -> None:
        if not self.require_project():
            return
        assert self.project is not None
        default = self.project.root / "video-associations.json"
        output, _ = QFileDialog.getSaveFileName(
            self, tr("导出视频关联记录"), str(default), "JSON (*.json)"
        )
        if not output:
            return
        try:
            Path(output).write_text(
                json.dumps(self.project.data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError as exc:
            self.show_error(f"{tr('无法导出关联记录：')}{exc}")
            return
        self.append_log(f"{tr('已导出关联记录：')}{output}")

    def _build_detail_drawer(self) -> None:
        self.detail_drawer = QFrame(self.library_panel)
        self.detail_drawer.setObjectName("videoDetailDrawer")
        drawer_layout = QVBoxLayout(self.detail_drawer)
        drawer_layout.setContentsMargins(14, 12, 14, 14)
        drawer_layout.setSpacing(6)

        header = QHBoxLayout()
        self.detail_title = QLabel(tr("视频详情"))
        self.detail_title.setObjectName("detailTitle")
        self.detail_title.setWordWrap(True)
        close_button = QPushButton(tr("关闭"))
        close_button.setFixedWidth(64)
        close_button.clicked.connect(self._close_detail_drawer)
        header.addWidget(self.detail_title, 1)
        header.addWidget(close_button, 0, Qt.AlignmentFlag.AlignTop)
        drawer_layout.addLayout(header)

        self.detail_preview = ResponsiveVideoThumbnail(tr("选择视频后显示封面"))
        self.detail_preview.setObjectName("detailPreview")
        self.detail_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.detail_preview.setCursor(Qt.CursorShape.PointingHandCursor)
        self.detail_preview.setToolTip(tr("点击放大封面；播放按钮可核对完整视频"))
        self.detail_preview.clicked.connect(self._show_detail_preview)
        self.detail_video = QVideoWidget()
        self.detail_video.setObjectName("detailVideo")
        self.detail_media_stack = QStackedWidget()
        self.detail_media_stack.addWidget(self.detail_preview)
        self.detail_media_stack.addWidget(self.detail_video)
        drawer_layout.addWidget(self.detail_media_stack)
        self.detail_audio_output = QAudioOutput(self)
        self.detail_media_player = QMediaPlayer(self)
        self.detail_media_player.setAudioOutput(self.detail_audio_output)
        self.detail_media_player.setVideoOutput(self.detail_video)
        self.detail_media_player.playbackStateChanged.connect(
            self._sync_detail_play_button
        )
        self.detail_media_player.errorOccurred.connect(self._detail_playback_failed)

        self.detail_seek_bar = QWidget()
        self.detail_seek_bar.setObjectName("detailSeekBar")
        seek_layout = QHBoxLayout(self.detail_seek_bar)
        seek_layout.setContentsMargins(0, 0, 0, 0)
        seek_layout.setSpacing(8)
        self.detail_position_label = QLabel("0:00")
        self.detail_position_label.setObjectName("detailPositionLabel")
        self.detail_position_label.setFixedWidth(40)
        self.detail_position_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.detail_seek_slider = QSlider(Qt.Orientation.Horizontal)
        self.detail_seek_slider.setObjectName("detailSeekSlider")
        self.detail_seek_slider.setRange(0, 0)
        self.detail_seek_slider.setEnabled(False)
        self.detail_duration_label = QLabel("0:00")
        self.detail_duration_label.setObjectName("detailDurationLabel")
        self.detail_duration_label.setFixedWidth(40)
        self.detail_duration_label.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        seek_layout.addWidget(self.detail_position_label)
        seek_layout.addWidget(self.detail_seek_slider, 1)
        seek_layout.addWidget(self.detail_duration_label)
        self.detail_seek_bar.hide()
        drawer_layout.addWidget(self.detail_seek_bar)
        self.detail_media_player.durationChanged.connect(self._detail_duration_changed)
        self.detail_media_player.positionChanged.connect(self._detail_position_changed)
        self.detail_seek_slider.sliderMoved.connect(self._detail_seek_moved)
        self.detail_seek_slider.sliderReleased.connect(self._detail_seek_released)

        self.detail_badge = QLabel()
        self.detail_badge.setObjectName("detailBadge")
        self.detail_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.detail_badge.setSizePolicy(
            QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed
        )
        drawer_layout.addWidget(self.detail_badge, 0, Qt.AlignmentFlag.AlignLeft)

        metrics = QHBoxLayout()
        metrics.setSpacing(8)
        for title, attribute in (
            (tr("分辨率"), "detail_resolution"),
            (tr("时长"), "detail_duration"),
            (tr("大小"), "detail_size"),
        ):
            card = QFrame()
            card.setObjectName("detailMetricCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(2)
            label = QLabel(title)
            label.setObjectName("detailMetricLabel")
            value = QLabel("—")
            value.setObjectName("detailMetricValue")
            card_layout.addWidget(label)
            card_layout.addWidget(value)
            metrics.addWidget(card, 1)
            setattr(self, attribute, value)
        drawer_layout.addLayout(metrics)

        for attribute, word_wrap in (
            ("detail_references", False),
            ("detail_path", True),
            ("detail_hash", False),
            ("detail_status", False),
        ):
            value = QLabel()
            value.setObjectName("detailValue")
            value.setWordWrap(word_wrap)
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            drawer_layout.addWidget(value)
            setattr(self, attribute, value)
        self.detail_status.setObjectName("detailStatus")
        drawer_layout.addStretch(1)

        actions = QHBoxLayout()
        self.detail_play_button = QPushButton(tr("播放"))
        self.detail_play_button.clicked.connect(self.preview_selected)
        self.detail_ai_button = QPushButton(tr("AI 建议"))
        self.detail_ai_button.clicked.connect(self.request_ai_suggestion)
        self.detail_ai_button.setToolTip(
            tr("由 AI 建议当前视频的命名、分类和同源归并；不会自动修改视频库。")
        )
        review_button = QPushButton(tr("核对关联"))
        review_button.setObjectName("primaryAction")
        review_button.clicked.connect(self.review_selected_family)
        location_button = QPushButton(tr("打开位置"))
        location_button.clicked.connect(self.open_selected_location)
        actions.addWidget(self.detail_play_button)
        actions.addWidget(self.detail_ai_button)
        actions.addWidget(review_button)
        actions.addWidget(location_button)
        drawer_layout.addLayout(actions)
        self.detail_drawer.hide()

    def _position_detail_drawer(self) -> None:
        if not hasattr(self, "detail_drawer"):
            return
        tree_rect = self.video_tree.geometry()
        width = min(540, max(420, round(tree_rect.width() * 0.40)))
        preview_width = max(1, width - 28)
        preview_height = round(preview_width * 9 / 16)
        media_height = max(160, min(preview_height, tree_rect.height() - 295))
        compact = tree_rect.height() < 620 or width < 480
        self.detail_path.setVisible(not compact)
        self.detail_hash.setVisible(not compact)
        self.detail_preview.setFixedHeight(media_height)
        self.detail_video.setFixedHeight(media_height)
        self.detail_media_stack.setFixedHeight(media_height)
        self.detail_drawer.setGeometry(
            tree_rect.right() - width + 1,
            tree_rect.top(),
            width,
            tree_rect.height(),
        )
        self.detail_drawer.raise_()

    def _update_detail_drawer(self) -> None:
        self.detail_media_player.stop()
        self.detail_media_stack.setCurrentWidget(self.detail_preview)
        self.detail_seek_bar.hide()
        self._reset_detail_seek_bar()
        if self.project is None or self.video_tree.currentItem() is None:
            self.detail_drawer.hide()
            return
        family, variant = self.selected_family_variant()
        if family is None:
            self.detail_drawer.hide()
            return
        target = variant or self.project.source_variant(family)
        references = sum(
            asset.get("family_id") == family["id"]
            for deck in self.project.decks()
            for asset in deck.get("assets", [])
        )
        width = int(target.get("width") or 0)
        height = int(target.get("height") or 0)
        duration = float(target.get("duration_sec") or 0)
        self.detail_title.setText(family["name"])
        self.detail_badge.setText(
            tr("高清源")
            if target["id"] == family.get("source_variant_id")
            else tr("候选版本")
        )
        self.detail_resolution.setText(f"{width}×{height}" if width and height else "—")
        self.detail_duration.setText(_format_duration(duration) if duration else "—")
        self.detail_size.setText(_format_mb(int(target.get("size_bytes") or 0)))
        self.detail_references.setText(
            f"{tr('关联　')}{len(family['variants'])}{tr(' 个版本 · ')}{references}{tr(' 个 PPTX 引用')}"
        )
        path_text = str(target.get("path") or "—")
        digest = str(target.get("sha256") or "—")
        self.detail_path.setText(f"{tr('文件位置  ')}{path_text}")
        self.detail_path.setToolTip(path_text)
        self.detail_hash.setText(f"{tr('文件哈希  ')}{digest[:16]}")
        self.detail_hash.setToolTip(digest)
        status = (
            tr("媒体不可读")
            if target.get("probe_error")
            else {
                "available": tr("正常"),
                "missing": tr("文件丢失"),
                "modified": tr("文件已修改"),
                "metadata_drift": tr("待校验（时间戳变化）"),
            }.get(self.project.status(target), tr("未知"))
        )
        self.detail_status.setText(f"{tr('状态　')}{status}")
        self.detail_status.setProperty("healthy", status == tr("正常"))
        self.detail_status.style().unpolish(self.detail_status)
        self.detail_status.style().polish(self.detail_status)
        source = self.project.root / str(target.get("path") or "")
        _set_video_thumbnail(
            self.detail_preview,
            source,
            str(target.get("sha256") or target["id"]),
            single_frame=True,
        )
        self._position_detail_drawer()
        self.detail_drawer.show()
        self.detail_drawer.raise_()

    def _sync_detail_play_button(self, state: QMediaPlayer.PlaybackState) -> None:
        self.detail_play_button.setText(
            tr("暂停")
            if state == QMediaPlayer.PlaybackState.PlayingState
            else tr("播放")
        )

    def _detail_duration_changed(self, duration: int) -> None:
        self.detail_seek_slider.setRange(0, max(0, duration))
        self.detail_duration_label.setText(_format_position(duration))
        self.detail_seek_slider.setEnabled(duration > 0)

    def _detail_position_changed(self, position: int) -> None:
        if self.detail_seek_slider.isSliderDown():
            return
        self.detail_seek_slider.setValue(max(0, position))
        self.detail_position_label.setText(_format_position(position))

    def _detail_seek_moved(self, value: int) -> None:
        self.detail_position_label.setText(_format_position(value))

    def _detail_seek_released(self) -> None:
        self.detail_media_player.setPosition(self.detail_seek_slider.value())

    def _reset_detail_seek_bar(self) -> None:
        self.detail_seek_slider.setRange(0, 0)
        self.detail_seek_slider.setValue(0)
        self.detail_seek_slider.setEnabled(False)
        self.detail_position_label.setText("0:00")
        self.detail_duration_label.setText("0:00")

    def _close_detail_drawer(self) -> None:
        self.detail_media_player.stop()
        self.detail_media_stack.setCurrentWidget(self.detail_preview)
        self.detail_seek_bar.hide()
        self._reset_detail_seek_bar()
        self.detail_drawer.hide()

    def _detail_playback_failed(self, _error: QMediaPlayer.Error, message: str) -> None:
        if not message:
            return
        self.detail_media_stack.setCurrentWidget(self.detail_preview)
        self.detail_seek_bar.hide()
        self._reset_detail_seek_bar()
        self.show_error(f"{tr('无法在窗口内播放该视频：')}{message}")

    def _show_detail_preview(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is None:
            return
        target = variant or self.project.source_variant(family)
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{family['name']}{tr(' · 封面预览')}")
        dialog.setMinimumSize(720, 480)
        dialog.resize(900, 600)
        dialog.setStyleSheet(MEDIA_MANAGER_STYLESHEET)
        layout = QVBoxLayout(dialog)
        preview = ResponsiveVideoThumbnail(tr("正在生成封面"))
        preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview.setStyleSheet(
            "background:#0b1017;border:1px solid #334155;border-radius:10px;"
        )
        layout.addWidget(preview, 1)
        source = self.project.root / str(target.get("path") or "")
        _set_video_thumbnail(
            preview,
            source,
            str(target.get("sha256") or target["id"]),
            single_frame=True,
        )
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def open_selected_location(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is None:
            return
        target = variant or self.project.source_variant(family)
        path = self.project.root / str(target.get("path") or "")
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.parent)))

    def _sort_by_header(self, section: int) -> None:
        mode = {
            0: "name",
            1: "resolution",
            2: "duration",
            3: "size",
            4: "hashes",
            5: "review",
            6: "path",
        }.get(section)
        if mode is None:
            return
        if self.library_sort_mode == mode:
            self.library_sort_descending = not self.library_sort_descending
        else:
            self.library_sort_mode = mode
            self.library_sort_descending = mode != "name"
        self.settings.setValue("video_library/sort", self.library_sort_mode)
        self.settings.setValue(
            "video_library/sort_descending", self.library_sort_descending
        )
        self.video_tree.header().setSortIndicator(
            section,
            (
                Qt.SortOrder.DescendingOrder
                if self.library_sort_descending
                else Qt.SortOrder.AscendingOrder
            ),
        )
        self.refresh_views()

    def _apply_library_filter(self, _text: str = "") -> None:
        query = self.library_filter_input.text().strip().casefold()
        review_filter = str(self.attention_filter_combo.currentData())
        active_button = self.library_stat_buttons.get(review_filter)
        if active_button is not None:
            active_button.setChecked(True)
        selected = self.video_tree.currentItem()
        selected_hidden = False
        visible_families = 0
        total_families = self.video_tree.topLevelItemCount()
        for index in range(self.video_tree.topLevelItemCount()):
            family_item = self.video_tree.topLevelItem(index)
            review_tags = set(
                str(family_item.data(0, REVIEW_TAGS_ROLE) or "").split(",")
            )
            review_matches = review_filter == "all" or review_filter in review_tags
            family_text = " ".join(
                family_item.text(column)
                for column in range(self.video_tree.columnCount())
            ).casefold()
            family_matches = review_matches and (not query or query in family_text)
            child_matches = False
            for child_index in range(family_item.childCount()):
                child = family_item.child(child_index)
                child_text = " ".join(
                    child.text(column)
                    for column in range(self.video_tree.columnCount())
                ).casefold()
                matches = review_matches and (
                    family_matches or not query or query in child_text
                )
                child.setHidden(not matches)
                child_matches = child_matches or matches
                if child is selected and not matches:
                    selected_hidden = True
            family_visible = family_matches or child_matches
            family_item.setHidden(not family_visible)
            visible_families += int(family_visible)
            if family_item is selected and not family_visible:
                selected_hidden = True
            if query and child_matches:
                family_item.setExpanded(True)
        if selected_hidden:
            self.video_tree.clearSelection()
        if total_families:
            self.video_tree.setVisible(visible_families > 0)
            self.library_empty.setVisible(visible_families == 0)
            if visible_families == 0:
                self.library_empty.setText(
                    tr("没有符合当前条件的视频\n\n清除查找文字，或切换为“全部视频”")
                )
        self.library_stats_label.setText(
            f"{tr('显示 ')}{visible_families}/{total_families}"
        )
        self._update_action_states()

    def _filter_library_by_stat(self, filter_name: str) -> None:
        index = self.attention_filter_combo.findData(filter_name)
        if index >= 0:
            self.attention_filter_combo.setCurrentIndex(index)

    def _set_library_stat_counts(
        self,
        family_count: int,
        variant_count: int,
        review_count: int,
        unlinked_count: int,
        multi_count: int,
        abnormal_count: int,
    ) -> None:
        compact = getattr(self, "_compact_layout", False)
        counts = {
            "all": (tr("全部"), family_count),
            "review": (tr("待核") if compact else tr("待核对"), review_count),
            "unlinked": (tr("无关联"), unlinked_count),
            "multi": (tr("多版") if compact else tr("多版本"), multi_count),
            "abnormal": (tr("异常") if compact else tr("文件异常"), abnormal_count),
        }
        for key, (title, count) in counts.items():
            self.library_stat_buttons[key].setText(f"{title} {count}")
        self.library_version_label.setText(f"{variant_count}{tr(' 版本')}")

    def _sorted_families(self) -> list[dict[str, Any]]:
        assert self.project is not None
        mode = self.library_sort_mode

        def metrics(family: dict[str, Any]) -> tuple[int, float, int, int]:
            source = self.project.source_variant(family)
            pixels = int(source.get("width") or 0) * int(source.get("height") or 0)
            return (
                pixels,
                float(source.get("duration_sec") or 0),
                int(source.get("bitrate_kbps") or 0),
                int(source.get("size_bytes") or 0),
            )

        families = list(self.project.families())
        if mode == "name":
            families.sort(key=metrics, reverse=True)
            return sorted(
                families,
                key=lambda family: _normalized_video_name(family["name"]),
                reverse=self.library_sort_descending,
            )
        index = {"resolution": 0, "duration": 1, "size": 3}.get(mode)
        if index is not None:
            return sorted(
                families,
                key=lambda family: metrics(family)[index],
                reverse=self.library_sort_descending,
            )
        if mode == "hashes":
            return sorted(
                families,
                key=lambda family: len(family.get("known_hashes", [])),
                reverse=self.library_sort_descending,
            )
        if mode == "review":
            reference_counts: dict[str, int] = {}
            for deck in self.project.decks():
                for asset in deck.get("assets", []):
                    family_id = asset["family_id"]
                    reference_counts[family_id] = reference_counts.get(family_id, 0) + 1
            return sorted(
                families,
                key=lambda family: (
                    any(
                        self.project.status(variant) != "available"
                        for variant in family["variants"]
                    ),
                    reference_counts.get(family["id"], 0) == 0,
                    len(family["variants"]) > 1,
                ),
                reverse=self.library_sort_descending,
            )
        if mode == "path":
            return sorted(
                families,
                key=lambda family: str(
                    self.project.source_variant(family).get("path") or ""
                ).casefold(),
                reverse=self.library_sort_descending,
            )
        return sorted(families, key=metrics, reverse=self.library_sort_descending)

    def _apply_style(self) -> None:
        self.setStyleSheet(MEDIA_MANAGER_STYLESHEET)

    def _update_action_states(self) -> None:
        project_ready = self.project is not None and not self.is_running
        self.health_button.setEnabled(project_ready)
        self.library_actions.setVisible(self.project is not None)
        selected = project_ready and self.video_tree.currentItem() is not None
        for button in self.project_action_buttons:
            button.setEnabled(project_ready)
        pptx_ready = project_ready and bool(self.input_paths)
        self.archive_button.setEnabled(pptx_ready)
        self.upgrade_button.setEnabled(pptx_ready)
        for button in self.selection_action_buttons:
            button.setEnabled(selected)
        item = self.video_tree.currentItem() if selected else None
        self.activate_button.setEnabled(
            bool(item and item.data(1, Qt.ItemDataRole.UserRole))
        )
        family, variant = self.selected_family_variant() if selected else (None, None)
        referenced = bool(
            self.project
            and variant
            and any(
                asset.get("original_variant_id") == variant["id"]
                for deck in self.project.decks()
                for asset in deck.get("assets", [])
            )
        )
        self.quarantine_button.setEnabled(
            bool(
                family
                and variant
                and variant.get("probe_error")
                and variant["id"]
                not in {
                    family.get("source_variant_id"),
                    family.get("active_variant_id"),
                }
                and not referenced
            )
        )
        self.merge_button.setEnabled(
            bool(selected and self.project and len(self.project.families()) > 1)
        )
        ai_configured = self._ai_config() is not None
        ai_enabled = bool(selected and ai_configured)
        self.ai_button.setEnabled(ai_enabled)
        self.detail_ai_button.setEnabled(ai_enabled)
        ai_tooltip = (
            tr("先用代码规则筛选候选，再由 AI 建议同源归并、主视频和命名；")
            + tr("视觉模型可额外参考三帧联系图，不会自动修改视频库。")
            if ai_configured
            else tr("请先点击顶栏齿轮配置并验证 AI；未配置时视频库功能不受影响。")
        )
        self.ai_button.setToolTip(ai_tooltip)
        self.detail_ai_button.setToolTip(ai_tooltip)

    def on_settings_changed(self) -> None:
        self._update_action_states()

    @staticmethod
    def _ai_config():
        from pptx_tools.ai_client import AIConfig

        app = QApplication.instance()
        if app is None:
            return None
        base_url = str(app.property("doc_media_ai_base_url") or "").strip()
        model = str(app.property("doc_media_ai_model") or "").strip()
        if not base_url or not model:
            return None
        return AIConfig(
            base_url,
            model,
            str(app.property("doc_media_ai_api_key") or ""),
            vision_enabled=bool(app.property("doc_media_ai_vision_enabled")),
            timeout_seconds=45,
            context=str(app.property("doc_media_ai_context") or ""),
        )

    def choose_new_project(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, tr("选择视频库文件夹"))
        if not directory:
            return
        try:
            self.project = VideoProject.create(Path(directory))
            self._project_opened()
        except Exception as exc:
            self.show_error(str(exc))

    def choose_open_project(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, tr("打开视频库"))
        if not directory:
            return
        root = Path(directory)
        if not (root / "video-project.json").is_file():
            nested = [
                path.parent
                for path in root.glob("*/video-project.json")
                if path.is_file()
            ]
            if len(nested) == 1:
                root = nested[0]
        self.open_project(root)

    def open_project(self, root: Path, *, report_errors: bool = True) -> None:
        try:
            self.project = VideoProject.open(root)
            self._project_opened(report_recovery=report_errors)
        except Exception as exc:
            if report_errors:
                self.show_error(str(exc))

    def _project_opened(self, *, report_recovery: bool = True) -> None:
        assert self.project is not None
        self.settings.setValue("video_manager/last_project", str(self.project.root))
        self.settings.sync()
        self._close_detail_drawer()
        self.library_filter_input.clear()
        self.library_stat_buttons["all"].setChecked(True)
        recovered = self.project.recovered_from_backup
        warning = tr("  ·  ⚠ 已从备份恢复") if recovered else ""
        project_path = str(self.project.root)
        if len(project_path) > 72:
            project_path = f"…/{Path(*self.project.root.parts[-3:])}"
        self.project_label.setText(
            f"{tr('当前视频库：')}{self.project.data['name']}{warning}{tr('  ·  ')}{project_path}"
        )
        self.project_label.setToolTip(
            f"{tr('视频库目录：')}{self.project.root}\n"
            + tr("库内保存：video-project.json、media、_cleanup、reports\n")
            + f"{tr('应用日志：')}{log_directory()}\n"
            + f"{tr('全局偏好：')}{self.settings.fileName()}"
        )
        self.append_log(f"{tr('已打开视频库：')}{self.project.root}", reveal=False)
        self.log_shelf.setText(
            f"{tr('状态与日志 · ')}{tr('已打开视频库：')}{self.project.root}"
        )
        if recovered:
            message = tr(
                "主视频库清单损坏或不可读，程序已使用最近的有效备份恢复。"
            ) + tr("建议立即检查视频状态。")
            self.append_log(f"{tr('警告：')}{message} {self.project.recovery_detail}")
            if report_recovery:
                QMessageBox.warning(self, tr("视频库已从备份恢复"), message)
        self.refresh_views()

    def on_activated(self) -> None:
        if self.project is None:
            last_project = self.settings.value("video_manager/last_project", "", str)
            if last_project and (Path(last_project) / "video-project.json").is_file():
                self.open_project(Path(last_project), report_errors=False)
            return
        if self.is_running:
            return
        try:
            self.project.reload()
            self.refresh_views()
        except Exception as exc:
            self.append_log(f"{tr('项目刷新失败：')}{exc}")

    def run_library_health(self, *, verify_hashes: bool = False) -> None:
        if not self.require_project():
            return

        def operation(progress, cancelled):
            assert self.project is not None
            progress(
                tr("正在完整核对全部视频哈希")
                if verify_hashes
                else tr("正在检查视频库清单与文件状态")
            )
            if cancelled():
                raise RuntimeError("Operation cancelled")
            return audit_video_project(
                self.project,
                verify_hashes=verify_hashes,
                progress_callback=progress,
                cancel_callback=cancelled,
            )

        self.run_operation(
            tr("正在进行视频库完整哈希复核")
            if verify_hashes
            else tr("正在进行视频库体检"),
            operation,
            self._library_health_finished,
        )

    def _library_health_finished(self, report: dict[str, Any]) -> None:
        stats = report["stats"]
        self.append_log(
            tr(
                "视频库体检完成：错误 {}，警告 {}，历史信息 {}，失效输出记录 {}。"
            ).format(
                stats["errors"],
                stats["warnings"],
                stats["info"],
                stats["stale_output_records"],
            )
        )
        LibraryHealthDialog(self, self, report).exec()

    def set_files(self, paths: Iterable[Path]) -> None:
        self.on_activated()
        self.input_paths = list(
            dict.fromkeys(
                resolved
                for path in paths
                if (resolved := Path(path).expanduser().resolve()).is_file()
                and resolved.suffix.lower() == ".pptx"
            )
        )
        self._refresh_workflow_files()
        self._update_action_states()
        if self.input_paths:
            self.input_summary.setToolTip(
                "\n".join(str(path) for path in self.input_paths[:10])
                + (
                    f"\n{tr('另有 ')}{len(self.input_paths) - 10}{tr(' 个文件')}"
                    if len(self.input_paths) > 10
                    else ""
                )
            )
            self.append_log(
                f"{tr('已接收 ')}{len(self.input_paths)}{tr(' 个 PPTX。')}",
                reveal=False,
            )

    def _refresh_workflow_files(self) -> None:
        while self.workflow_files_layout.count():
            item = self.workflow_files_layout.takeAt(0)
            widget = item.widget()
            if widget is not None and widget is not self.input_summary:
                widget.deleteLater()
        self.workflow_chip_widgets.clear()
        self.input_summary.setVisible(not self.input_paths)
        if not self.input_paths:
            self.workflow_files_layout.addWidget(self.input_summary, 1)
            return
        for path in self.input_paths[:2]:
            name = path.name if len(path.name) <= 34 else f"{path.stem[:25]}….pptx"
            chip = QPushButton(f"▣  {name}  ×")
            chip.setObjectName("inputChip")
            chip.setToolTip(f"{path}\n{tr('点击从当前工作流移除')}")
            chip.clicked.connect(
                lambda _checked=False, value=path: self._remove_input(value)
            )
            self.workflow_files_layout.addWidget(chip)
            self.workflow_chip_widgets.append(chip)
        if len(self.input_paths) > 2:
            remaining = QPushButton(
                f"{tr('还有 ')}{len(self.input_paths) - 2}{tr(' 个')}"
            )
            remaining.setObjectName("inputChip")
            menu = QMenu(remaining)
            for path in self.input_paths[2:]:
                action = menu.addAction(f"{path.name}  ×")
                action.setToolTip(str(path))
                action.triggered.connect(
                    lambda _checked=False, value=path: self._remove_input(value)
                )
            remaining.setMenu(menu)
            self.workflow_files_layout.addWidget(remaining)
            self.workflow_chip_widgets.append(remaining)
        self.workflow_files_layout.addStretch(1)
        clear_button = QPushButton(tr("清空"))
        clear_button.clicked.connect(self.clear_input_paths)
        self.workflow_files_layout.addWidget(clear_button)
        self.workflow_chip_widgets.append(clear_button)

    def _remove_input(self, path: Path) -> None:
        self.input_paths = [item for item in self.input_paths if item != path]
        self._refresh_workflow_files()
        self._update_action_states()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        suffixes = {".pptx", *VIDEO_SUFFIXES}
        if self._dropped_paths(event, suffixes):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:  # noqa: N802
        self.dragEnterEvent(event)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        if (
            watched is getattr(self, "content_widget", None)
            and event.type() == QEvent.Type.Resize
        ):
            self.update_responsive_layout(event.size().width())
            self._position_operation_log()
        if watched is getattr(self, "library_panel", None):
            if event.type() == QEvent.Type.Resize:
                self._position_detail_drawer()
            return super().eventFilter(watched, event)
        if watched in {
            getattr(self, "input_summary", None),
            getattr(self, "workflow_files", None),
        }:
            if event.type() == QEvent.Type.DragEnter:
                if self._dropped_paths(event, {".pptx"}):
                    event.acceptProposedAction()
                else:
                    event.ignore()
                return event.isAccepted()
            if event.type() == QEvent.Type.DragMove:
                if self._dropped_paths(event, {".pptx"}):
                    event.acceptProposedAction()
                else:
                    event.ignore()
                return event.isAccepted()
            if event.type() == QEvent.Type.Drop:
                paths = self._dropped_paths(event, {".pptx"})
                if paths:
                    self.set_files(paths)
                    event.acceptProposedAction()
                else:
                    event.ignore()
                return event.isAccepted()
        tree = getattr(self, "video_tree", None)
        if watched is tree or watched is getattr(self, "video_tree_viewport", None):
            if event.type() in {QEvent.Type.DragEnter, QEvent.Type.DragMove}:
                if self._dropped_paths(event, VIDEO_SUFFIXES):
                    event.acceptProposedAction()
                else:
                    event.ignore()
                return event.isAccepted()
            if event.type() == QEvent.Type.Drop:
                paths = self._dropped_paths(event, VIDEO_SUFFIXES)
                if paths:
                    event.acceptProposedAction()
                    self.import_external_paths(paths)
                else:
                    event.ignore()
                return event.isAccepted()
        return super().eventFilter(watched, event)

    def update_responsive_layout(self, width: int) -> None:
        if not hasattr(self, "library_actions"):
            return
        compact = width < 1180
        if compact == getattr(self, "_compact_layout", None):
            return
        self._compact_layout = compact
        self.library_actions.setProperty("compact", compact)
        self.library_search_label.setVisible(True)
        self.library_version_label.setVisible(not compact)
        self.library_stats_label.setVisible(not compact)
        self.library_filter_input.setMinimumWidth(120 if compact else 220)
        self.library_filter_input.setMaximumWidth(180 if compact else 320)
        self.attention_filter_combo.setVisible(compact)
        self.activate_button.setVisible(not compact)
        self.pending_button.setVisible(not compact)
        for button in self.library_stat_buttons.values():
            button.setVisible(not compact)
            button.setMinimumWidth(56 if compact else 0)
        self.library_actions.style().unpolish(self.library_actions)
        self.library_actions.style().polish(self.library_actions)
        self._set_library_stat_counts(
            self.video_tree.topLevelItemCount(),
            int(self.library_version_label.text().split()[0] or 0),
            int(self.library_stat_buttons["review"].text().split()[-1]),
            int(self.library_stat_buttons["unlinked"].text().split()[-1]),
            int(self.library_stat_buttons["multi"].text().split()[-1]),
            int(self.library_stat_buttons["abnormal"].text().split()[-1]),
        )

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        pptx_paths = self._dropped_paths(event, {".pptx"})
        if pptx_paths:
            self.set_files(pptx_paths)
            event.acceptProposedAction()
            return
        video_paths = self._dropped_paths(event, VIDEO_SUFFIXES)
        if video_paths:
            self.import_external_paths(video_paths)
            event.acceptProposedAction()
            return
        event.ignore()

    @staticmethod
    def _dropped_paths(
        event: QDragEnterEvent | QDragMoveEvent | QDropEvent,
        suffixes: set[str],
    ) -> list[Path]:
        if not event.mimeData().hasUrls():
            return []
        return [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile() and Path(url.toLocalFile()).suffix.lower() in suffixes
        ]

    def choose_input_pptx(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(
            self, tr("选择 PPTX"), "", "PowerPoint (*.pptx)"
        )
        if selected:
            self.set_files(Path(path) for path in selected)

    def clear_input_paths(self) -> None:
        self.input_paths = []
        self.input_summary.setText(tr("可拖入或多选 PPTX，批量归档或高清回填"))
        self.input_summary.setToolTip(
            tr("视频源独立入库，PPTX 只记录形状关联。")
            + tr("匹配依据哈希和内容特征，改文件名或目录不受影响。")
        )
        self._refresh_workflow_files()
        self._update_action_states()

    def selected_family_variant(
        self,
    ) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        if self.project is None or self.video_tree.currentItem() is None:
            return None, None
        item = self.video_tree.currentItem()
        family_id = item.data(0, Qt.ItemDataRole.UserRole)
        variant_id = item.data(1, Qt.ItemDataRole.UserRole)
        family = self.project.family(family_id) if family_id else None
        if variant_id:
            _, variant = self.project.find_variant(variant_id)
            return family, variant
        return family, None

    def refresh_views(self) -> None:
        self.video_tree.clear()
        if self.project is None:
            self.pending_button.setText(tr("待清理 (0)"))
            self._set_library_stat_counts(0, 0, 0, 0, 0, 0)
            self.library_stats_label.setText(tr("显示 0/0"))
            self.video_tree.hide()
            self.library_empty.setText(
                tr("先新建或打开一个视频库\n\n可从 PPTX 提取，或导入外部视频自动匹配")
            )
            self.library_empty.show()
            self._update_action_states()
            return
        try:
            pending_count = len(self.project.pending_cleanup())
        except Exception:
            pending_count = 0
        self.pending_button.setText(f"{tr('待清理 (')}{pending_count})")
        has_families = bool(self.project.families())
        self.video_tree.setVisible(has_families)
        self.library_empty.setVisible(not has_families)
        if not has_families:
            self.library_empty.setText(
                tr("视频库还是空的\n\n可归档 PPTX 高清视频，或导入外部视频创建视频族")
            )
        reference_counts: dict[str, int] = {}
        for deck in self.project.decks():
            for asset in deck.get("assets", []):
                family_id = asset["family_id"]
                reference_counts[family_id] = reference_counts.get(family_id, 0) + 1
        review_tags_by_family: dict[str, set[str]] = {}
        variant_statuses: dict[str, str] = {}
        for family in self.project.families():
            tags: set[str] = set()
            if reference_counts.get(family["id"], 0) == 0:
                tags.add("unlinked")
            if len(family["variants"]) > 1:
                tags.add("multi")
            for variant in family["variants"]:
                status = self.project.status(variant)
                variant_statuses[variant["id"]] = status
                if status in {"missing", "modified"} or variant.get("probe_error"):
                    tags.add("abnormal")
                elif status == "metadata_drift":
                    tags.add("review")
            if tags:
                tags.add("review")
            review_tags_by_family[family["id"]] = tags
        families = self.project.families()
        family_count = len(families)
        variant_count = sum(len(family["variants"]) for family in families)
        review_count = sum("review" in tags for tags in review_tags_by_family.values())
        unlinked_count = sum(
            "unlinked" in tags for tags in review_tags_by_family.values()
        )
        multi_count = sum("multi" in tags for tags in review_tags_by_family.values())
        abnormal_count = sum(
            "abnormal" in tags for tags in review_tags_by_family.values()
        )
        self._set_library_stat_counts(
            family_count,
            variant_count,
            review_count,
            unlinked_count,
            multi_count,
            abnormal_count,
        )
        for family in self._sorted_families():
            source = self.project.source_variant(family)
            source_id = source["id"]
            width = source.get("width", 0)
            height = source.get("height", 0)
            duration = source.get("duration_sec", 0)
            family_item = QTreeWidgetItem(
                [
                    family["name"],
                    f"{width}×{height}" if width and height else tr("未识别"),
                    _format_duration(duration) if duration else tr("未识别"),
                    f"{source['size_bytes'] / 1024 / 1024:.1f} MB",
                    f"{len(family.get('known_hashes', []))}{tr(' 个')}",
                    f"{len(family['variants'])}{tr(' 版本 · ')}"
                    + f"{reference_counts.get(family['id'], 0)}{tr(' 引用')}",
                    Path(source["path"]).parent.as_posix(),
                ]
            )
            for column in range(self.video_tree.columnCount()):
                font = family_item.font(column)
                font.setBold(True)
                family_item.setFont(column, font)
                family_item.setForeground(column, QColor("#f1f5f9"))
                family_item.setBackground(column, QBrush(QColor("#152235")))
                family_item.setToolTip(column, family_item.text(column))
            family_item.setData(0, Qt.ItemDataRole.UserRole, family["id"])
            family_item.setData(
                0,
                REVIEW_TAGS_ROLE,
                ",".join(sorted(review_tags_by_family[family["id"]])),
            )
            self.video_tree.addTopLevelItem(family_item)
            variants = sorted(
                family["variants"],
                key=lambda variant: (
                    variant["id"] == source_id,
                    self.project._variant_quality_key(variant),
                ),
                reverse=True,
            )
            for variant in variants:
                marker = (
                    tr("高清源 · ") if variant["id"] == source_id else tr("候选 · ")
                )
                width = variant.get("width", 0)
                height = variant.get("height", 0)
                duration = variant.get("duration_sec", 0)
                resolution = f"{width}×{height}" if width and height else tr("未识别")
                duration_text = _format_duration(duration) if duration else tr("未识别")
                size = f"{variant['size_bytes'] / 1024 / 1024:.1f} MB"
                status = (
                    tr("媒体不可读")
                    if variant.get("probe_error")
                    else {
                        "available": tr("正常"),
                        "missing": tr("丢失"),
                        "modified": tr("已修改"),
                        "metadata_drift": tr("待校验"),
                    }[variant_statuses[variant["id"]]]
                )
                child = QTreeWidgetItem(
                    [
                        f"{marker}{variant['label']}",
                        resolution,
                        duration_text,
                        size,
                        variant["sha256"][:8],
                        status,
                        variant["path"],
                    ]
                )
                child.setData(0, Qt.ItemDataRole.UserRole, family["id"])
                child.setData(1, Qt.ItemDataRole.UserRole, variant["id"])
                for column in range(self.video_tree.columnCount()):
                    child.setForeground(column, QColor("#b8c4d4"))
                    child.setToolTip(column, child.text(column))
                origins = variant.get("origin_paths", [])
                if origins:
                    child.setToolTip(
                        6,
                        f"{tr('库内：')}{variant['path']}\n{tr('来源：')}"
                        + "\n".join(str(path) for path in origins),
                    )
                family_item.addChild(child)
            family_item.setExpanded(True)
        self._apply_library_filter()
        self._update_action_states()

    def archive_pptx_videos(self) -> None:
        if not self.require_project():
            return
        paths = list(self.input_paths)
        if not paths:
            selected, _ = QFileDialog.getOpenFileNames(
                self, tr("选择 PPTX"), "", "PowerPoint (*.pptx)"
            )
            paths = [Path(path) for path in selected]
        if not paths:
            return
        source_quality = str(self.source_quality_combo.currentData())
        category = self.category_input.text().strip()
        try:
            normalize_library_category(category)
        except ValueError as exc:
            self.show_error(f"{tr('入库分类无效：')}{exc}")
            return

        def operation(progress, cancelled):
            assert self.project is not None
            results = []
            for path in paths:
                result = self.project.archive_and_register_pptx(
                    path,
                    source_quality=source_quality,
                    category=category,
                    progress_callback=progress,
                    cancel_callback=cancelled,
                )
                results.append(result)
            return results

        self.run_operation(
            tr("正在提取并归档高清源视频"),
            operation,
            self._archive_finished,
        )

    def _archive_finished(self, results: list[dict[str, Any]]) -> None:
        self.append_log(
            tr("入库完成：登记 {} 份 PPTX，新增 {} 个视频族，复用 {} 个，")
            + tr("待确认高清候选 {} 个。").format(
                len({item["deck"]["id"] for item in results}),
                sum(item["added"] for item in results),
                sum(item["reused"] for item in results),
                sum(item.get("candidates_added", 0) for item in results),
            )
        )
        self.clear_input_paths()

    def import_external_videos(self) -> None:
        if not self.require_project():
            return
        start = self.settings.value("video_manager/last_external_dir", "", str)
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            tr("导入外部视频"),
            start,
            "Video files (*.mp4 *.mov *.m4v *.wmv *.avi *.mkv *.webm)",
        )
        if not selected:
            return
        paths = [Path(path) for path in selected]
        self.settings.setValue("video_manager/last_external_dir", str(paths[0].parent))
        self.import_external_paths(paths)

    def import_external_paths(self, paths: Iterable[Path]) -> None:
        if not self.require_project():
            return
        paths = [
            path.expanduser().resolve()
            for path in paths
            if path.is_file() and path.suffix.lower() in VIDEO_SUFFIXES
        ]
        if not paths:
            self.show_error(tr("没有可导入的受支持视频文件。"))
            return
        category = self.category_input.text().strip()
        try:
            normalize_library_category(category)
        except ValueError as exc:
            self.show_error(f"{tr('入库分类无效：')}{exc}")
            return
        source_quality = str(self.source_quality_combo.currentData())

        def operation(progress, cancelled):
            assert self.project is not None
            results = []
            for index, path in enumerate(paths, start=1):
                if cancelled():
                    raise RuntimeError("Operation cancelled")
                progress(
                    f"{tr('正在匹配外部视频 ')}{index}/{len(paths)}{tr('：')}{path.name}"
                )
                try:
                    result = self.project.import_external_video(
                        path,
                        source_quality=source_quality,
                        category=category,
                        defer_suggestions=True,
                    )
                except Exception as exc:
                    result = {
                        "status": "failed",
                        "source": str(path),
                        "error": str(exc),
                    }
                results.append(result)
            return results

        self.run_operation(
            tr("正在导入并匹配外部视频"),
            operation,
            lambda results: self._review_external_imports(
                results, source_quality, category
            ),
        )

    def _review_external_imports(
        self,
        results: list[dict[str, Any]],
        source_quality: str,
        category: str,
    ) -> None:
        decisions: list[tuple[Path, str | None, bool]] = []
        completed = []
        for item in results:
            if item["status"] not in {"ambiguous", "suggested"}:
                completed.append(item)
                continue
            dialog = VideoMatchDialog(
                self,
                item,
                allow_new_family=True,
                allow_remember=False,
            )
            if dialog.exec() != QDialog.DialogCode.Accepted:
                completed.append({**item, "status": "skipped"})
                continue
            decisions.append(
                (
                    Path(item["source"]),
                    None if dialog.create_new_family else dialog.selected_family_id,
                    dialog.create_new_family,
                )
            )
        if not decisions:
            self._external_import_finished(completed)
            return

        def operation(progress, cancelled):
            assert self.project is not None
            reviewed = []
            for index, (path, family_id, force_new) in enumerate(decisions, start=1):
                if cancelled():
                    raise RuntimeError("Operation cancelled")
                progress(
                    f"{tr('正在应用人工匹配 ')}{index}/{len(decisions)}{tr('：')}{path.name}"
                )
                try:
                    reviewed.append(
                        self.project.import_external_video(
                            path,
                            source_quality=source_quality,
                            category=category,
                            family_id=family_id,
                            manual_confirmed=family_id is not None,
                            force_new_family=force_new,
                        )
                    )
                except Exception as exc:
                    reviewed.append(
                        {
                            "status": "failed",
                            "source": str(path),
                            "error": str(exc),
                        }
                    )
            return reviewed

        self.run_operation(
            tr("正在保存人工确认的视频关联"),
            operation,
            lambda reviewed: self._external_import_finished([*completed, *reviewed]),
        )

    def _external_import_finished(self, results: list[dict[str, Any]]) -> None:
        counts = {
            status: sum(item["status"] == status for item in results)
            for status in (
                "matched",
                "created",
                "existing",
                "ambiguous",
                "suggested",
                "skipped",
                "failed",
            )
        }
        self.append_log(
            tr(
                "外部视频导入完成：匹配 {}，新建 {}，已存在 {}，跳过 {}，失败 {}。"
            ).format(
                counts["matched"],
                counts["created"],
                counts["existing"],
                counts["skipped"] + counts["ambiguous"] + counts["suggested"],
                counts["failed"],
            )
        )
        for item in results:
            if item["status"] in {"ambiguous", "suggested", "skipped"}:
                self.append_log(f"{tr('已跳过，未导入：')}{item['source']}")
            elif item["status"] == "failed":
                self.append_log(
                    f"{tr('失败：')}{item['source']}{tr('；')}{item['error']}"
                )
            elif item["status"] == "matched" and not item.get("promoted"):
                self.append_log(
                    f"{tr('已加入候选版本：')}{item['source']} → {item['family_name']}{tr('；')}"
                    + tr("如需替换母版，请选择该版本后点击“设为高清源”。")
                )
        if counts["ambiguous"] or counts["suggested"] or counts["failed"]:
            self.show_error(tr("部分视频未导入，详情见操作记录。"))

    def upgrade_pptx(self) -> None:
        if not self.require_project():
            return
        paths = list(self.input_paths)
        if not paths:
            selected, _ = QFileDialog.getOpenFileNames(
                self, tr("选择需要高清优化的 PPTX"), "", "PowerPoint (*.pptx)"
            )
            if not selected:
                return
            paths = [Path(path) for path in selected]

        review_root = Path(tempfile.mkdtemp(prefix="pptx-tools-match-review-"))

        def operation(progress, cancelled):
            assert self.project is not None
            try:
                return [
                    {
                        "source": str(source),
                        "items": self.project.review_pptx_matches(
                            source,
                            review_root / str(index),
                            include_resolved=True,
                            progress_callback=progress,
                            cancel_callback=cancelled,
                        ),
                    }
                    for index, source in enumerate(paths)
                ]
            except Exception:
                shutil.rmtree(review_root, ignore_errors=True)
                raise

        self.run_operation(
            tr("正在分析 PPTX 视频匹配"),
            operation,
            lambda analyses: self._review_pptx_upgrade(paths, analyses, review_root),
        )

    def _review_pptx_upgrade(
        self,
        paths: list[Path],
        analyses: list[dict[str, Any]],
        review_root: Path,
    ) -> None:
        overrides: dict[Path, dict[str, str]] = {}
        remembered: dict[Path, set[str]] = {}
        kept: dict[Path, set[str]] = {}
        tiers: dict[Path, str] = {}
        try:
            assert self.project is not None
            family_choices: list[dict[str, Any]] = []
            for family in self.project.families():
                try:
                    variant = self.project.source_variant(family)
                    source_path = self.project.require_variant_path(variant)
                except (FileNotFoundError, KeyError):
                    continue
                family_choices.append(
                    {
                        "id": family["id"],
                        "name": family["name"],
                        "source_path": str(source_path),
                        "source_sha256": variant["sha256"],
                        "resolution": (
                            f"{int(variant.get('width') or 0)}×"
                            + f"{int(variant.get('height') or 0)}"
                        ),
                        "width": int(variant.get("width") or 0),
                        "height": int(variant.get("height") or 0),
                        "bitrate_kbps": int(variant.get("bitrate_kbps") or 0),
                        "video_codec": str(variant.get("video_codec") or ""),
                        "audio_codec": str(variant.get("audio_codec") or ""),
                        "suffix": source_path.suffix.lower(),
                    }
                )
            family_choices.sort(key=lambda item: item["name"].casefold())
            for analysis in analyses:
                source = Path(analysis["source"])
                items = analysis["items"]
                if not items:
                    self.append_log(f"{tr('未发现内嵌视频：')}{source}")
                    continue
                dialog = PptxUpgradeReviewDialog(self, source, items, family_choices)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    shutil.rmtree(review_root, ignore_errors=True)
                    self.append_log(tr("已取消高清回填，未修改视频库或 PPTX。"))
                    return
                source_overrides, source_remembered, source_kept = dialog.decisions()
                overrides[source] = source_overrides
                remembered[source] = source_remembered
                kept[source] = source_kept
                tiers[source] = dialog.quality_tier()
        except Exception:
            shutil.rmtree(review_root, ignore_errors=True)
            raise

        if len(paths) == 1:
            source = paths[0]
            tier_key = tiers.get(source, DEFAULT_BACKFILL_TIER)
            default = source.with_name(
                f"{source.stem}_{BACKFILL_QUALITY_TIERS[tier_key]['suffix']}.pptx"
            )
            output, _ = QFileDialog.getSaveFileName(
                self, tr("保存高清优化 PPTX"), str(default), "PowerPoint (*.pptx)"
            )
            if not output:
                shutil.rmtree(review_root, ignore_errors=True)
                return
            outputs = {source: Path(output)}
        else:
            directory = QFileDialog.getExistingDirectory(
                self, tr("选择高清优化输出目录"), str(paths[0].parent)
            )
            if not directory:
                shutil.rmtree(review_root, ignore_errors=True)
                return
            outputs = {
                source: Path(directory)
                / (
                    f"{source.stem}_"
                    + f"{BACKFILL_QUALITY_TIERS[tiers.get(source, DEFAULT_BACKFILL_TIER)]['suffix']}"
                    + ".pptx"
                )
                for source in paths
            }

        def operation(progress, cancelled):
            assert self.project is not None
            return [
                self.project.upgrade_pptx_from_library(
                    source,
                    output_path=outputs[source],
                    family_overrides=overrides.get(source),
                    remember_manual_matches=remembered.get(source, set()),
                    keep_current_media=kept.get(source, set()),
                    quality_tier=tiers.get(source, DEFAULT_BACKFILL_TIER),
                    progress_callback=progress,
                    cancel_callback=cancelled,
                )
                for source in paths
            ]

        self.run_operation(
            tr("正在匹配视频库并高清优化 PPTX"),
            operation,
            lambda results: self._finish_reviewed_upgrade(results, review_root),
        )

    def _finish_reviewed_upgrade(
        self, results: list[dict[str, Any]], review_root: Path
    ) -> None:
        shutil.rmtree(review_root, ignore_errors=True)
        self._upgrade_finished(results)

    def _upgrade_finished(self, results: list[dict[str, Any]]) -> None:
        generated = [item for item in results if item.get("output_pptx")]
        self.append_log(
            tr("高清优化完成：生成 {} 个 PPTX，替换 {} 个视频，已是高清 {} 个，")
            + tr("保持当前 {} 个，固化压缩哈希 {} 个，未匹配 {} 个。").format(
                len(generated),
                sum(item["matched"] for item in results),
                sum(item.get("already_high_quality", 0) for item in results),
                sum(item.get("kept_current", 0) for item in results),
                sum(item.get("aliases_added", 0) for item in results),
                sum(len(item["unmatched"]) for item in results),
            )
        )
        manual_matched = sum(item.get("manual_matched", 0) for item in results)
        if manual_matched:
            self.append_log(f"{tr('人工确认并回填 ')}{manual_matched}{tr(' 个视频。')}")
        labels = sorted(
            {
                spec["label"]
                for item in generated
                for spec in [
                    BACKFILL_QUALITY_TIERS.get(item.get("quality_tier") or "best")
                ]
                if spec and spec["label"] != tr("最佳")
            }
        )
        if labels:
            self.append_log(f"{tr('回填档位：')}{tr('、').join(labels)}")
        for item in generated:
            self.append_log(f"{tr('输出：')}{item['output_pptx']}")
        if not generated:
            self.show_error(
                tr("没有生成输出：视频均已是高清源，或没有安全匹配的视频。")
            )
        self.clear_input_paths()

    def rename_selected(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is None:
            self.show_error(tr("请选择一个视频或版本。"))
            return
        current = Path(variant["path"]).stem if variant else family["name"]
        value, ok = QInputDialog.getText(self, tr("改名"), tr("新名称"), text=current)
        if not ok or not value.strip():
            return
        try:
            if variant:
                self.run_operation(
                    tr("正在重命名视频"),
                    lambda _progress, _cancelled: self.project.rename_variant_file(
                        variant["id"], value
                    ),
                )
                return
            self.project.rename_family_and_source(family["id"], value)
            self.refresh_views()
        except Exception as exc:
            self.show_error(str(exc))

    def preview_selected(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is None:
            return
        target = variant or self.project.source_variant(family)
        try:
            path = self.project.require_variant_path(target)
        except Exception as exc:
            self.show_error(str(exc))
            return
        if (
            self.detail_media_player.source().toLocalFile() == str(path)
            and self.detail_media_player.playbackState()
            == QMediaPlayer.PlaybackState.PlayingState
        ):
            self.detail_media_player.pause()
            return
        if self.detail_media_player.source().toLocalFile() != str(path):
            self._reset_detail_seek_bar()
        self.detail_media_player.setSource(QUrl.fromLocalFile(str(path)))
        self.detail_media_stack.setCurrentWidget(self.detail_video)
        self.detail_seek_bar.show()
        self.detail_drawer.show()
        self.detail_drawer.raise_()
        self.detail_media_player.play()

    def move_selected_variant(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is not None and variant is None:
            category, accepted = QInputDialog.getText(
                self,
                tr("移动整个视频族"),
                tr("库内分类目录（留空表示媒体根目录）："),
                text=family.get("category", ""),
            )
            if not accepted:
                return
            try:
                normalize_library_category(category)
            except ValueError as exc:
                self.show_error(str(exc))
                return
            self.run_operation(
                tr("正在移动视频族"),
                lambda _progress, _cancelled: self.project.move_family(
                    family["id"], category
                ),
            )
            return
        if variant is None:
            self.show_error(tr("请选择一个视频或具体版本。"))
            return
        directory = QFileDialog.getExistingDirectory(self, tr("选择归档目录"))
        if not directory:
            return
        self.run_operation(
            tr("正在移动视频"),
            lambda _progress, _cancelled: self.project.move_variant(
                variant["id"], Path(directory)
            ),
        )

    def import_version(self) -> None:
        if self.project is None:
            return
        family, _ = self.selected_family_variant()
        if family is None:
            self.show_error(tr("请选择要添加版本的视频。"))
            return
        source, _ = QFileDialog.getOpenFileName(
            self,
            tr("选择视频版本"),
            "",
            "Video files (*.mp4 *.mov *.m4v *.wmv *.avi *.mkv *.webm)",
        )
        if not source:
            return
        self.run_operation(
            tr("正在验证并添加视频版本"),
            lambda _progress, _cancelled: self.project.import_external_video(
                Path(source),
                source_quality=str(self.source_quality_combo.currentData()),
                family_id=family["id"],
            ),
        )

    def set_selected_source_variant(self) -> None:
        if self.project is None:
            return
        _, variant = self.selected_family_variant()
        if variant is None:
            self.show_error(tr("请选择一个具体视频版本。"))
            return
        try:
            warnings = self.project.compatibility_warnings(variant["id"])
            if warnings:
                answer = QMessageBox.question(
                    self,
                    tr("高清源兼容性提示"),
                    "\n".join(warnings) + tr("\n\n仍设为高清源吗？"),
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            self.project.set_source_variant(variant["id"])
            self.refresh_views()
        except Exception as exc:
            self.show_error(str(exc))

    def request_ai_suggestion(self) -> None:
        from pptx_tools.ai_client import OpenAICompatibleClient, privacy_scope

        if self.ai_thread is not None:
            self.ai_ignore_result = True
            if self.ai_worker is not None:
                self.ai_worker.cancel()
            self.ai_button.setText(tr("已取消显示"))
            self.append_log(
                tr("已取消显示本次 AI 建议；已发出的网络请求会在后台结束。")
            )
            return
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        config = self._ai_config()
        if family is None or config is None:
            self.show_error(tr("请先在顶栏设置中配置支持图片输入的 AI。"))
            return
        privacy_key = f"ai/privacy_confirmed/{privacy_scope(config)}"
        if not self.settings.value(privacy_key, False, bool):
            sent_content = (
                tr("三帧联系图、名称、规格、时长和大小")
                if config.vision_enabled
                else tr("名称、规格、时长、大小与代码相似度（不发送视频画面）")
            )
            answer = QMessageBox.question(
                self,
                tr("AI 视频分析"),
                (
                    tr("将向已配置的 AI 服务发送代码筛选出的最多 6 个视频候选的")
                    + f"{sent_content}{tr('，不发送完整视频、PPTX ')}"
                    + tr("或本地路径。是否继续？")
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            self.settings.setValue(privacy_key, True)
        selected_variant = variant or self.project.source_variant(family)
        selected_path = self.project.require_variant_path(selected_variant)
        self.ai_target_family_id = family["id"]
        self.ai_ignore_result = False
        self.ai_request_config = config
        items: list[dict[str, Any]] = []
        item_names: dict[str, str] = {}
        seen: set[str] = set()

        def append_item(
            item_family: dict[str, Any],
            item_variant: dict[str, Any],
            code_similarity: float | None,
        ) -> None:
            if item_variant["id"] in seen or len(items) >= 6:
                return
            seen.add(item_variant["id"])
            path = self.project.require_variant_path(item_variant)
            preview = ""
            if config.vision_enabled:
                preview_root = (
                    Path(tempfile.gettempdir()) / "pptx-tools-ai-video-previews"
                )
                preview_root.mkdir(parents=True, exist_ok=True)
                cutoff = datetime.now().timestamp() - 7 * 24 * 60 * 60
                for cached in preview_root.glob("*.jpg"):
                    try:
                        if cached.stat().st_mtime < cutoff:
                            cached.unlink()
                    except OSError:
                        pass
                preview_path = preview_root / f"{item_variant['sha256'][:32]}.jpg"
                if not preview_path.is_file():
                    create_video_thumbnail(path, preview_path)
                if preview_path.is_file():
                    preview = preview_path
            label = (
                f"{item_family['name']}{tr(' · ')}"
                + f"{item_variant.get('profile') or Path(path).suffix.lstrip('.')}"
            )
            item_names[item_variant["id"]] = label
            items.append(
                {
                    "id": item_variant["id"],
                    "name": label,
                    "width": int(item_variant.get("width") or 0),
                    "height": int(item_variant.get("height") or 0),
                    "format": Path(path).suffix.lstrip("."),
                    "size_bytes": int(item_variant.get("size_bytes") or 0),
                    "duration_sec": float(item_variant.get("duration_sec") or 0),
                    "bitrate_kbps": int(
                        item_variant.get("video_bitrate_kbps")
                        or item_variant.get("bitrate_kbps")
                        or 0
                    ),
                    "health": (
                        tr("异常")
                        if item_variant.get("probe_error")
                        or self.project.status(item_variant) in {"missing", "modified"}
                        else tr("待校验")
                        if self.project.status(item_variant) == "metadata_drift"
                        else tr("正常")
                    ),
                    "code_similarity": code_similarity,
                    "preview_path": preview,
                }
            )

        append_item(family, selected_variant, 100.0)
        for current in family["variants"]:
            append_item(family, current, 100.0)
        try:
            candidates = self.project.suggest_video_matches(selected_path, limit=6)[
                "candidates"
            ]
        except Exception as exc:
            self.append_log(f"{tr('AI 候选预筛失败，改用当前视频族：')}{exc}")
            candidates = []
        for candidate in candidates:
            candidate_family = self.project.family(candidate["family_id"])
            candidate_variant = self.project.find_variant(
                candidate["source_variant_id"]
            )[1]
            append_item(candidate_family, candidate_variant, candidate["score"])
        self.ai_item_names = item_names
        self.ai_button.setEnabled(True)
        self.ai_button.setText(tr("取消 AI 分析"))
        self.detail_ai_button.setEnabled(False)
        self.detail_ai_button.setText(tr("分析中…"))
        self.append_log(tr("AI 正在分析代码候选、视频联系图与规格…"))
        self.log_hide_timer.stop()
        self.ai_thread = QThread(self)
        self.ai_worker = OperationWorker(
            lambda _message, cancelled: OpenAICompatibleClient(config).organize_media(
                items, media_kind=tr("视频"), cancel_callback=cancelled
            )
        )
        self.ai_worker.moveToThread(self.ai_thread)
        self.ai_thread.started.connect(self.ai_worker.run)
        self.ai_worker.finished.connect(self._show_ai_suggestion)
        self.ai_worker.failed.connect(self._ai_failed)
        self.ai_worker.finished.connect(self.ai_thread.quit)
        self.ai_worker.failed.connect(self.ai_thread.quit)
        self.ai_worker.finished.connect(self.ai_worker.deleteLater)
        self.ai_worker.failed.connect(self.ai_worker.deleteLater)
        self.ai_thread.finished.connect(self._ai_thread_finished)
        self.ai_thread.finished.connect(self.ai_thread.deleteLater)
        self.ai_thread.start()

    def _show_ai_suggestion(self, result: object) -> None:
        if self.ai_ignore_result:
            return
        if not isinstance(result, dict):
            self._ai_failed(tr("AI 返回结果无效。"))
            return
        if self.project is None:
            return
        try:
            family = self.project.family(self.ai_target_family_id)
        except KeyError:
            self.append_log(tr("AI 建议对应的视频已不存在，结果已忽略。"))
            return
        from pptx_tools.ai_review_dialog import AISuggestionDialog

        dialog = AISuggestionDialog(
            self,
            result,
            {"suggested_name": family["name"]},
            {"suggested_name": tr("名称")},
            self.ai_item_names,
        )
        applied: list[str] = []
        if dialog.exec() == QDialog.DialogCode.Accepted:
            values = dialog.selected_values()
            if values.get("suggested_name"):
                try:
                    self.project.rename_family_and_source(
                        self.ai_target_family_id, values["suggested_name"]
                    )
                    applied.append(tr("名称"))
                    self.refresh_views()
                except Exception as exc:
                    self.show_error(str(exc))
        merged = self._review_ai_video_merge_groups(result.get("merge_groups") or [])
        if merged:
            applied.append(f"{tr('归并 ')}{merged}{tr(' 个视频族')}")
        self.append_log(
            tr("AI 视频整理建议已完成")
            + (
                f"{tr('，已应用：')}{tr('、').join(applied)}{tr('。')}"
                if applied
                else tr("，未修改视频库。")
            )
        )
        if self.ai_request_config is not None:
            from pptx_tools.app_logging import write_ai_audit_event

            write_ai_audit_event(
                media_kind="video",
                target_id=self.ai_target_family_id,
                provider=self.ai_request_config.base_url,
                model=self.ai_request_config.model,
                vision_enabled=self.ai_request_config.vision_enabled,
                applied_fields=applied,
                merge_group_count=len(result.get("merge_groups") or []),
            )

    def _review_ai_video_merge_groups(self, groups: list[dict[str, Any]]) -> int:
        if self.project is None:
            return 0
        merged = 0
        for group in groups:
            try:
                primary_family, _ = self.project.find_variant(group["primary_id"])
            except (KeyError, TypeError):
                continue
            candidate_families: list[dict[str, Any]] = []
            seen: set[str] = set()
            for variant_id in group.get("item_ids") or []:
                try:
                    family, _ = self.project.find_variant(variant_id)
                except KeyError:
                    continue
                if family["id"] not in seen:
                    candidate_families.append(family)
                    seen.add(family["id"])
            for source in candidate_families:
                if source["id"] == primary_family["id"]:
                    continue
                impact = self.project.family_merge_impact(
                    source["id"], primary_family["id"]
                )
                answer = QMessageBox.question(
                    self,
                    tr("核对 AI 视频归并建议"),
                    (
                        f"{tr('疑似同一视频：')}{source['name']} → {primary_family['name']}\n"
                        + f"{tr('建议主资源：')}{primary_family['name']}\n"
                        + f"{tr('置信度：')}{float(group.get('confidence') or 0):.0%}\n"
                        + f"{tr('理由：')}{group.get('reason') or tr('未说明')}\n\n"
                        + f"{tr('确认后将迁移 ')}{impact['variant_count']}{tr(' 个版本、')}"
                        + f"{impact['reference_count']}{tr(' 处 PPTX 引用和 ')}"
                        + f"{impact['known_hash_count']}{tr(' 个哈希别名。')}"
                        + tr("是否确认它们内容相同？")
                    ),
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    continue
                try:
                    self.project.merge_families(
                        source["id"],
                        primary_family["id"],
                        confirmed_same_content=True,
                    )
                except Exception as exc:
                    self.show_error(str(exc))
                    continue
                merged += 1
        if merged:
            self.refresh_views()
        return merged

    def _ai_failed(self, message: str) -> None:
        if self.ai_ignore_result:
            return
        self.append_log(f"{tr('AI 分析失败：')}{message}")
        self.show_error(message)

    def _ai_thread_finished(self) -> None:
        self.ai_worker = None
        self.ai_thread = None
        self.ai_button.setText(tr("AI 整理建议"))
        self.detail_ai_button.setText(tr("AI 建议"))
        self.log_hide_timer.start(2500)
        self._update_action_states()

    def _select_family(self, family_id: str) -> None:
        for index in range(self.video_tree.topLevelItemCount()):
            item = self.video_tree.topLevelItem(index)
            if item.data(0, Qt.ItemDataRole.UserRole) == family_id:
                self.video_tree.setCurrentItem(item)
                return

    def quarantine_selected_abnormal(self) -> None:
        if self.project is None:
            return
        family, variant = self.selected_family_variant()
        if family is None or variant is None:
            self.show_error(tr("请选择一个异常视频版本。"))
            return
        answer = QMessageBox.question(
            self,
            tr("隔离异常版本"),
            f"{tr('将“')}{family['name']}{tr(' · ')}{variant.get('label', '')}{tr('”移到待清理目录。')}\n"
            + tr("不会删除文件，也不会改变其他版本或 PPTX 关联。是否继续？"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.run_operation(
            tr("正在隔离异常视频版本"),
            lambda _progress, _cancelled: self.project.quarantine_abnormal_variant(
                variant["id"]
            ),
            lambda _result: self.append_log(
                tr("异常版本已移到待清理目录，可在“待清理”中恢复。")
            ),
        )

    def merge_selected_family(self) -> None:
        if self.project is None:
            return
        source, _ = self.selected_family_variant()
        if source is None:
            self.show_error(tr("请选择需要归并的视频。"))
            return
        targets = [
            family for family in self.project.families() if family["id"] != source["id"]
        ]
        if not targets:
            self.show_error(tr("没有其他视频可供归并。"))
            return
        source_variant = self.project.source_variant(source)
        try:
            source_path = self.project.require_variant_path(source_variant)
        except Exception as exc:
            self.show_error(str(exc))
            return

        self.run_operation(
            tr("正在查找相似视频族"),
            lambda _progress, _cancelled: {
                **self.project.suggest_video_matches(
                    source_path, limit=min(12, len(self.project.families()))
                ),
                "source_family_id": source["id"],
            },
            self._merge_candidates_ready,
        )

    def _merge_candidates_ready(self, item: dict[str, Any]) -> None:
        if self.project is None:
            return
        source_family_id = item["source_family_id"]
        item["candidates"] = [
            candidate
            for candidate in item["candidates"]
            if candidate["family_id"] != source_family_id
        ]
        if not item["candidates"]:
            self.show_error(tr("没有其他可供核对的视频族。"))
            return
        dialog = VideoMatchDialog(
            self,
            item,
            allow_new_family=False,
            allow_remember=False,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.selected_family_id is None:
            return
        source = self.project.family(source_family_id)
        target = self.project.family(dialog.selected_family_id)
        impact = self.project.family_merge_impact(source["id"], target["id"])
        deck_names = tr("、").join(impact["deck_names"][:3])
        if len(impact["deck_names"]) > 3:
            deck_names += f"{tr(' 等 ')}{impact['deck_count']}{tr(' 个')}"
        association_text = (
            f"\n{tr('将迁移 ')}{impact['reference_count']}{tr(' 处 PPTX 视频引用')}"
            + f"{tr('（')}{impact['deck_count']}{tr(' 个 PPTX）')}"
        )
        if deck_names:
            association_text += f"{tr('：')}{deck_names}"
        answer = QMessageBox.question(
            self,
            tr("确认归并视频"),
            f"{tr('将“')}{source['name']}{tr('”归并到“')}{target['name']}{tr('”。')}\n"
            + f"{tr('将迁移 ')}{impact['variant_count']}{tr(' 个版本和 ')}"
            + f"{impact['known_hash_count']}{tr(' 个已知哈希别名。')}"
            + f"{association_text}\n\n"
            + tr("目标视频当前的高清源会继续作为高清源；原 PPTX 引用和压缩版哈希")
            + tr("会一并改到目标族，并在保存时校验。是否确认它们是同一视频？"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.project.merge_families(
                source["id"],
                target["id"],
                confirmed_same_content=True,
            )
            self.refresh_views()
        except Exception as exc:
            self.show_error(str(exc))

    def review_selected_family(self) -> None:
        if self.project is None:
            return
        family, _ = self.selected_family_variant()
        if family is None:
            self.show_error(tr("请选择需要核实的视频族或版本。"))
            return
        threshold = self.settings.value("cleanup/ssim_threshold", 0.95, float)
        family_id = family["id"]
        self.run_operation(
            f"{tr('正在核实“')}{family['name']}{tr('”的版本和重复候选…')}",
            lambda progress, cancelled: self.project.scan_cleanup_groups(
                ssim_threshold=threshold,
                focus_family_id=family_id,
                progress_callback=progress,
                cancel_callback=cancelled,
            ),
            lambda groups: self._selected_review_scan_finished(
                groups, threshold, family_id
            ),
        )

    def _selected_review_scan_finished(
        self,
        groups: list[dict[str, Any]],
        threshold: float,
        family_id: str,
    ) -> None:
        selected_groups = [
            group for group in groups if family_id in group.get("family_ids", [])
        ]
        if not selected_groups:
            self.append_log(tr("当前视频没有发现可整理的重复版本。"))
            _exec_centered_message(
                self,
                QMessageBox.Icon.Information,
                tr("核实版本"),
                tr("当前视频没有发现可安全归并或清理的重复版本。\n")
                + tr("如需指定高清基线，请选择具体版本后点击“设为高清源”。"),
            )
            return
        self._cleanup_scan_finished(selected_groups, threshold)

    def relink_missing(self) -> None:
        if self.project is None:
            return
        metadata_drift_count = sum(
            self.project.status(variant) == "metadata_drift"
            for family in self.project.families()
            for variant in family["variants"]
        )
        has_missing = any(
            self.project.status(variant) == "missing"
            for family in self.project.families()
            for variant in family["variants"]
        )
        roots: list[Path] = []
        if has_missing:
            root = QFileDialog.getExistingDirectory(
                self, tr("选择搜索目录"), str(self.project.root)
            )
            if not root:
                return
            roots.append(Path(root))
        elif not metadata_drift_count:
            self.append_log(tr("当前没有待校验或丢失的视频。"))
            return

        def operation(progress, cancelled):
            assert self.project is not None
            relinked = self.project.relink_missing(
                roots,
                progress_callback=progress,
                cancel_callback=cancelled,
            )
            remaining_drift = sum(
                self.project.status(variant) == "metadata_drift"
                for family in self.project.families()
                for variant in family["variants"]
            )
            return {
                "relinked": relinked,
                "metadata_refreshed": max(0, metadata_drift_count - remaining_drift),
                "remaining_drift": remaining_drift,
            }

        self.run_operation(
            tr("正在核验并重新关联视频"),
            operation,
            self._relink_missing_finished,
        )

    def _relink_missing_finished(self, result: dict[str, Any]) -> None:
        relinked = list(result.get("relinked") or [])
        refreshed = int(result.get("metadata_refreshed") or 0)
        remaining = int(result.get("remaining_drift") or 0)
        self.append_log(
            f"{tr('已核验并刷新 ')}{refreshed}{tr(' 个时间戳状态，重新关联 ')}"
            f"{len(relinked)}{tr(' 个视频版本。')}"
        )
        if remaining:
            self.append_log(
                f"{tr('仍有 ')}{remaining}{tr(' 个文件未通过哈希核验，请运行库体检。')}"
            )
        self.refresh_views()

    def start_cleanup_scan(self) -> None:
        if not self.require_project():
            return
        threshold = self.settings.value("cleanup/ssim_threshold", 0.95, float)
        threshold, ok = QInputDialog.getDouble(
            self,
            tr("整理视频库"),
            tr("“体积更小但质量接近”的 SSIM 阈值（0.90 - 1.00）："),
            threshold,
            0.90,
            1.00,
            2,
        )
        if not ok:
            return
        self.settings.setValue("cleanup/ssim_threshold", threshold)
        self.run_operation(
            tr("正在扫描视频库重复候选…"),
            lambda progress, cancelled: self.project.scan_cleanup_groups(
                ssim_threshold=threshold,
                progress_callback=progress,
                cancel_callback=cancelled,
            ),
            lambda groups: self._cleanup_scan_finished(groups, threshold),
        )

    def _cleanup_scan_finished(
        self, groups: list[dict[str, Any]], threshold: float
    ) -> None:
        if not groups:
            self.append_log(tr("没有发现可整理的重复版本。"))
            return
        self.append_log(
            f"{tr('发现 ')}{len(groups)}{tr(' 组重复候选，等待确认保留项。')}"
        )
        dialog = CleanupDialog(self, groups, threshold)
        if dialog.exec() != CleanupDialog.DialogCode.Accepted:
            self.append_log(tr("已取消整理。"))
            return
        decisions = dialog.decisions()
        if not decisions:
            actionable_groups = sum(
                1
                for group in groups
                if sum(
                    1
                    for candidate in group.get("candidates", [])
                    if candidate.get("auto_allowed", False)
                    and candidate.get("can_keep", True)
                )
                >= 2
            )
            blocked = [
                f"{candidate.get('label') or candidate.get('path') or tr('未命名版本')}{tr('：')}"
                + tr("、").join(candidate.get("block_reasons", []))
                for group in groups
                for candidate in group.get("candidates", [])
                if candidate.get("block_reasons")
            ]
            forceable_groups = sum(
                1
                for group in groups
                if group.get("kind") == "within_family"
                and any(
                    not candidate.get("auto_allowed", False)
                    and candidate.get("can_keep", True)
                    for candidate in group.get("candidates", [])
                )
                and any(
                    candidate.get("auto_allowed", False)
                    and candidate.get("can_keep", True)
                    for candidate in group.get("candidates", [])
                )
            )
            if actionable_groups == 0 and blocked:
                message = (
                    tr("没有执行整理：候选版本未通过安全一致性校验，")
                    + tr("不会自动移入待清理。\n\n")
                    + "\n".join(blocked[:6])
                )
                if forceable_groups:
                    message += tr(
                        "\n\n如确认要处理族内锁定版本，请勾选“人工确认：连锁定版本也移入待清理”，"
                    ) + tr("并完成二次确认。")
            else:
                message = tr("没有选择需要处理的组。")
            self.append_log(message.replace("\n", " "))
            QMessageBox.information(self, tr("整理视频库"), message)
            return
        forced = [
            variant_id
            for decision in decisions
            for variant_id in decision.get("force_remove_variant_ids", [])
        ]
        if forced:
            answer = QMessageBox.question(
                self,
                tr("确认强制整理"),
                tr("有版本未通过时长、音轨或内容一致性校验。\n")
                + f"{tr('将 ')}{len(forced)}{tr(' 个版本移入可恢复的待清理目录，可能不是同一视频；')}"
                + tr("不会直接删除。是否继续？"),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                self.append_log(tr("已取消强制整理。"))
                return
        self.run_operation(
            f"{tr('正在应用整理（')}{len(decisions)}{tr(' 组）…')}",
            lambda progress, cancelled: self.project.apply_cleanup_plan(
                decisions,
                progress_callback=progress,
                cancel_callback=cancelled,
            ),
            self._cleanup_apply_finished,
        )

    def _cleanup_apply_finished(self, result: dict[str, Any]) -> None:
        applied = result["applied"]
        failed = result["failed"]
        self.append_log(
            f"{tr('整理完成：')}{applied}{tr(' 组已处理，')}{failed}{tr(' 组失败。')}"
        )
        for item in result["results"]:
            if not item["ok"]:
                self.append_log(f"{tr('失败：')}{item['error']}")
        if failed:
            self.show_error(f"{failed}{tr(' 组整理失败，详情见日志。')}")
        self.refresh_views()

    def show_pending_cleanup(self) -> None:
        if not self.require_project():
            return
        PendingCleanupDialog(self, self).exec()
        self.refresh_views()

    def run_operation(
        self,
        title: str,
        operation: Callable[[Callable[[str], None], Callable[[], bool]], Any],
        done: Callable[[Any], None] | None = None,
    ) -> None:
        if self.is_running:
            return
        self.is_running = True
        self.status_frame.show()
        self.content_splitter.setEnabled(False)
        self.new_project_button.setEnabled(False)
        self.open_project_button.setEnabled(False)
        self.operation_done = done
        self.progress.setRange(0, 0)
        self.cancel_button.setEnabled(True)
        self.append_log(title)
        thread = QThread(self)
        worker = OperationWorker(operation)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.message.connect(self.append_log)
        worker.finished.connect(self.operation_succeeded)
        worker.failed.connect(self.operation_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(self.operation_thread_finished)
        self.worker_thread = thread
        self.worker = worker
        thread.start()

    @Slot(object)
    def operation_succeeded(self, result: Any) -> None:
        self.append_log(tr("操作完成。"))
        self.refresh_views()
        if self.operation_done:
            self.operation_done(result)

    @Slot(str)
    def operation_failed(self, message: str) -> None:
        self.append_log(message)
        self.refresh_views()
        self.show_error(message.split("\n", 1)[0])

    @Slot()
    def operation_thread_finished(self) -> None:
        if self.worker:
            self.worker.deleteLater()
        if self.worker_thread:
            self.worker_thread.deleteLater()
        self.worker = None
        self.worker_thread = None
        self.operation_done = None
        self.is_running = False
        self.content_splitter.setEnabled(True)
        self.new_project_button.setEnabled(True)
        self.open_project_button.setEnabled(True)
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.cancel_button.setEnabled(False)
        self.status_frame.hide()
        self.log_hide_timer.start(2500)
        self._update_action_states()

    def stop_job(self) -> None:
        if self.worker is None:
            return
        if self.worker:
            self.worker.cancel()
        if self.worker.thread_id is not None:
            terminate_active_processes(
                grace_seconds=0.2,
                owner_thread_id=self.worker.thread_id,
            )
        self.append_log(tr("正在停止当前操作…"))

    def require_project(self) -> bool:
        if self.project is not None:
            return True
        self.show_error(tr("请先新建或打开一个视频库。"))
        return False

    def append_log(self, message: str, *, reveal: bool = True) -> None:
        if message:
            self.log_output.appendPlainText(message)
            if reveal:
                self.toggle_operation_log(True)
                if not self.is_running:
                    self.log_hide_timer.start(2500)
            self.log_shelf.setText(f"{tr('状态与日志 · ')}{message.splitlines()[0]}")
            LOGGER.info("%s", message)

    def _position_operation_log(self) -> None:
        if self.log_output.isHidden():
            return
        origin = self.log_shelf.mapTo(
            self.content_widget, self.log_shelf.rect().topLeft()
        )
        height = min(220, max(140, self.content_widget.height() // 3))
        self.log_output.setGeometry(
            origin.x(),
            max(0, origin.y() - height - 4),
            self.log_shelf.width(),
            height,
        )
        self.log_output.raise_()

    def toggle_operation_log(self, visible: bool) -> None:
        self.log_output.setVisible(visible)
        self.log_panel_button.setChecked(visible)
        if visible:
            self._position_operation_log()
        else:
            self.log_hide_timer.stop()

    def open_logs(self) -> None:
        directory = log_directory()
        directory.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(directory)))

    def show_error(self, message: str) -> None:
        _exec_centered_message(
            self,
            QMessageBox.Icon.Warning,
            tr("视频库"),
            message,
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        self.detail_media_player.stop()
        if self.ai_thread is not None and self.ai_thread.isRunning():
            self.ai_ignore_result = True
            if self.ai_worker is not None:
                self.ai_worker.cancel()
            event.ignore()
            self.hide()
            self.ai_thread.finished.connect(self.close)
            return
        self.stop_job()
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            while not self.worker_thread.wait(250):
                terminate_active_processes(grace_seconds=0.1)
        super().closeEvent(event)


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
