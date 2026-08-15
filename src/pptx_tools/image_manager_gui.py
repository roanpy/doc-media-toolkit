from __future__ import annotations

import io
from pathlib import Path

from PIL import Image, ImageOps
from PySide6.QtCore import QSettings, QThread, Qt, QUrl
from PySide6.QtGui import QDesktopServices, QImage, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QSpinBox,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pptx_tools.manager_i18n import tr
from pptx_tools.image_manager import ImageProject
from pptx_tools.media_manager_ui import MEDIA_MANAGER_STYLESHEET, OperationWorker
from pptx_tools.ui_theme import (
    configure_ui_font,
    format_user_file_size,
    install_control_help,
)


IMAGE_MANAGER_STYLESHEET = (
    MEDIA_MANAGER_STYLESHEET
    + """
QFrame#imageDetailDrawer {
    background: #111827;
    border: 1px solid #334155;
    border-radius: 10px;
}
QLabel#imagePreview {
    background: #0b1017;
    border: 1px solid #334155;
    border-radius: 8px;
    color: #64748b;
}
"""
)
HEALTH_FILTER_TITLES = {
    "all": tr("全部"),
    "duplicate_origins": tr("重复来源"),
    "similar": tr("相似"),
    "undersized": tr("过小"),
    "no_origin": tr("无来源"),
}


class MetadataDialog(QDialog):
    def __init__(self, parent: QWidget, asset: dict) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr("编辑图片信息"))
        self.setMinimumWidth(460)
        self.setStyleSheet(IMAGE_MANAGER_STYLESHEET)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.name_input = QLineEdit(str(asset.get("name") or ""))
        self.category_input = QLineEdit(str(asset.get("category") or ""))
        self.tags_input = QLineEdit(tr("，").join(asset.get("tags") or []))
        self.summary_input = QLineEdit(str(asset.get("summary") or ""))
        form.addRow(tr("名称"), self.name_input)
        form.addRow(tr("分类"), self.category_input)
        form.addRow(tr("标签"), self.tags_input)
        form.addRow(tr("说明"), self.summary_input)
        layout.addLayout(form)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton(tr("取消"))
        cancel.clicked.connect(self.reject)
        save = QPushButton(tr("保存"))
        save.setObjectName("primaryAction")
        save.clicked.connect(self.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        layout.addLayout(buttons)

    def values(self) -> dict:
        tags = self.tags_input.text().replace(",", tr("，")).split(tr("，"))
        return {
            "name": self.name_input.text(),
            "category": self.category_input.text(),
            "tags": tags,
            "summary": self.summary_input.text(),
        }


class ImportPreviewDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        paths: list[Path],
        *,
        min_width: int = 0,
        min_height: int = 0,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr("预览待入库图片"))
        self.resize(860, 560)
        self.setMinimumSize(700, 460)
        self.setStyleSheet(IMAGE_MANAGER_STYLESHEET)
        self.records: list[dict] = []

        layout = QVBoxLayout(self)
        hint = QLabel(
            tr("以下为将从图片或文档中提取的实际图片。这里只预览，不修改源文件。")
        )
        hint.setObjectName("dialogSummary")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.list = QTreeWidget()
        self.list.setHeaderLabels([tr("来源文件"), tr("文档内位置"), tr("规格")])
        self.list.setColumnWidth(0, 210)
        self.list.setColumnWidth(1, 250)
        self.list.itemSelectionChanged.connect(self._show_selected)
        splitter.addWidget(self.list)
        self.preview = QLabel(tr("暂无可预览图片"))
        self.preview.setObjectName("imagePreview")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setMinimumSize(280, 260)
        splitter.addWidget(self.preview)
        splitter.setSizes([500, 320])
        layout.addWidget(splitter, 1)

        for source in paths:
            try:
                _source_type, images = ImageProject.source_images(source)
            except (OSError, ValueError) as exc:
                item = QTreeWidgetItem([source.name, str(exc), ""])
                item.setDisabled(True)
                self.list.addTopLevelItem(item)
                continue
            for member, data in images:
                pixmap = MainWindow._load_preview_bytes(data)
                size = (
                    f"{pixmap.width()}×{pixmap.height()}"
                    if not pixmap.isNull()
                    else tr("无法预览")
                )
                excluded = not pixmap.isNull() and (
                    pixmap.width() < min_width or pixmap.height() < min_height
                )
                if excluded:
                    size += tr(" · 将排除")
                index = len(self.records)
                self.records.append(
                    {"source": source, "member": member, "pixmap": pixmap}
                )
                item = QTreeWidgetItem([source.name, member, size])
                item.setData(0, Qt.ItemDataRole.UserRole, index)
                self.list.addTopLevelItem(item)
        if self.list.topLevelItemCount():
            self.list.setCurrentItem(self.list.topLevelItem(0))

        close = QPushButton(tr("关闭"))
        close.clicked.connect(self.accept)
        actions = QHBoxLayout()
        actions.addStretch(1)
        actions.addWidget(close)
        layout.addLayout(actions)

    def _show_selected(self) -> None:
        item = self.list.currentItem()
        index = item.data(0, Qt.ItemDataRole.UserRole) if item is not None else None
        if not isinstance(index, int) or not 0 <= index < len(self.records):
            self.preview.setPixmap(QPixmap())
            self.preview.setText(tr("暂无可预览图片"))
            return
        pixmap = self.records[index]["pixmap"]
        self.preview.setText("")
        self.preview.setPixmap(
            pixmap.scaled(
                self.preview.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )


class SimilarImageReviewDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        project: ImageProject,
        left: dict,
        right: dict,
        score: float,
    ) -> None:
        super().__init__(parent)
        self.decision = ""
        self.setWindowTitle(tr("核对相似图片"))
        self.resize(760, 520)
        self.setMinimumSize(680, 460)
        self.setStyleSheet(IMAGE_MANAGER_STYLESHEET)
        layout = QVBoxLayout(self)
        hint = QLabel(
            f"{tr('代码相似分 ')}{score:.1f}{tr('。请查看两张图片后决定；')}"
            + tr("合并只整理图片库，原始图片和文档不受影响。")
        )
        hint.setObjectName("dialogSummary")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        previews = QHBoxLayout()
        for label, asset in ((tr("左图"), left), (tr("右图"), right)):
            pane = QVBoxLayout()
            title = QLabel(
                f"{label}{tr(' · ')}{asset['name']}\n"
                + f"{asset['width']}×{asset['height']}{tr(' · ')}{asset['format']}{tr(' · ')}"
                + f"{format_user_file_size(asset['size_bytes'])}"
            )
            title.setWordWrap(True)
            preview = QLabel(tr("无法预览"))
            preview.setObjectName("imagePreview")
            preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
            preview.setMinimumSize(300, 280)
            pixmap = MainWindow._load_preview(project.asset_path(asset))
            if not pixmap.isNull():
                preview.setText("")
                preview.setPixmap(
                    pixmap.scaled(
                        300,
                        280,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
            pane.addWidget(title)
            pane.addWidget(preview, 1)
            previews.addLayout(pane, 1)
        layout.addLayout(previews, 1)
        buttons = QHBoxLayout()
        ignore = QPushButton(tr("不是同一张"))
        ignore.setToolTip(tr("记住这对图片，以后不再作为相似候选提示。"))
        ignore.clicked.connect(lambda: self._choose("ignore"))
        close = QPushButton(tr("关闭"))
        close.clicked.connect(self.reject)
        keep_left = QPushButton(tr("保留左图并合并"))
        keep_left.clicked.connect(lambda: self._choose("left"))
        keep_right = QPushButton(tr("保留右图并合并"))
        keep_right.clicked.connect(lambda: self._choose("right"))
        buttons.addWidget(ignore)
        buttons.addStretch(1)
        buttons.addWidget(close)
        buttons.addWidget(keep_left)
        buttons.addWidget(keep_right)
        layout.addLayout(buttons)

    def _choose(self, decision: str) -> None:
        self.decision = decision
        self.accept()


class PendingImageCleanupDialog(QDialog):
    def __init__(self, parent: QWidget, window: MainWindow) -> None:
        super().__init__(parent)
        self.window = window
        self.setWindowTitle(tr("图片待清理目录"))
        self.resize(820, 500)
        self.setMinimumSize(680, 420)
        self.setStyleSheet(IMAGE_MANAGER_STYLESHEET)
        layout = QVBoxLayout(self)
        self.summary = QLabel()
        self.summary.setObjectName("dialogSummary")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [tr("文件"), tr("原位置"), tr("大小"), tr("原因"), tr("隔离时间")]
        )
        self.tree.setColumnWidth(0, 220)
        self.tree.setColumnWidth(1, 220)
        self.tree.setColumnWidth(2, 90)
        self.tree.setColumnWidth(3, 180)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(self.tree, 1)
        buttons = QHBoxLayout()
        self.restore_button = QPushButton(tr("还原选中"))
        self.restore_button.clicked.connect(self.restore_selected)
        self.empty_button = QPushButton(tr("永久清空"))
        self.empty_button.setObjectName("dangerAction")
        self.empty_button.clicked.connect(self.empty_all)
        close_button = QPushButton(tr("关闭"))
        close_button.clicked.connect(self.accept)
        buttons.addWidget(self.restore_button)
        buttons.addWidget(self.empty_button)
        buttons.addStretch(1)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self.reload()

    def reload(self) -> None:
        project = self.window.project
        self.tree.clear()
        if project is None:
            self.summary.setText(tr("尚未打开图片库。"))
            self.restore_button.setEnabled(False)
            self.empty_button.setEnabled(False)
            return
        entries = project.pending_cleanup()
        issues = project.cleanup_pending_issues()
        total = sum(entry["size_bytes"] for entry in entries)
        self.summary.setText(
            (
                f"{tr('共 ')}{len(entries)}{tr(' 个文件，')}{format_user_file_size(total)}{tr('。')}"
                + (
                    f"\n⚠ {tr('；').join(issues)}"
                    if issues
                    else tr("\n可还原或永久清空。")
                )
            )
            if entries
            else tr("当前没有待清理文件。")
        )
        for entry in entries:
            item = QTreeWidgetItem(
                [
                    Path(entry["quarantined_path"]).name,
                    str(entry.get("original_path") or ""),
                    format_user_file_size(entry["size_bytes"]),
                    str(entry.get("reason") or ""),
                    str(entry.get("quarantined_at") or ""),
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, entry["token"])
            self.tree.addTopLevelItem(item)
        self.restore_button.setEnabled(bool(entries))
        self.empty_button.setEnabled(bool(entries) and not issues)

    def restore_selected(self) -> None:
        item = self.tree.currentItem()
        if item is None or self.window.project is None:
            return
        try:
            restored = self.window.project.restore_cleanup_entry(
                str(item.data(0, Qt.ItemDataRole.UserRole))
            )
        except Exception as exc:
            QMessageBox.warning(self, tr("图片待清理目录"), str(exc))
            return
        self.window.status_label.setText(f"{tr('状态与日志 · 已还原 ')}{restored.name}")
        self.window.refresh_views()
        self.reload()

    def empty_all(self) -> None:
        if self.window.project is None:
            return
        answer = QMessageBox.question(
            self,
            tr("永久清空"),
            tr("清空后无法从工具箱恢复，原始文档不受影响。是否继续？"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            removed = self.window.project.empty_cleanup()
        except Exception as exc:
            QMessageBox.warning(self, tr("图片待清理目录"), str(exc))
            return
        self.window.status_label.setText(
            f"{tr('状态与日志 · 已永久清空 ')}{removed}{tr(' 个文件')}"
        )
        self.window.refresh_views()
        self.reload()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        configure_ui_font(QApplication.instance())
        self.project: ImageProject | None = None
        self.pending_paths: list[Path] = []
        self.worker_thread: QThread | None = None
        self.worker: OperationWorker | None = None
        self.ai_thread: QThread | None = None
        self.ai_worker: OperationWorker | None = None
        self.ai_target_asset_id = ""
        self.ai_ignore_result = False
        self.ai_item_names: dict[str, str] = {}
        self.ai_request_config = None
        self.health_filter = "all"
        self.setWindowTitle(tr("图片库 - 文档媒体工具箱"))
        self.resize(1280, 760)
        self.setMinimumSize(900, 600)
        self.setAcceptDrops(True)
        self.setStyleSheet(IMAGE_MANAGER_STYLESHEET)
        self._build_ui()
        install_control_help(self)
        last_project = QSettings("Doc Media Toolkit", "Doc Media Toolkit").value(
            "image_manager/last_project", "", str
        )
        if last_project and (Path(last_project) / "image-project.json").is_file():
            self.open_project(Path(last_project), report_errors=False)

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("central")
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(8)

        header = QFrame()
        header.setObjectName("headerCard")
        header_layout = QHBoxLayout(header)
        title_stack = QVBoxLayout()
        title = QLabel(tr("文档图片资产库"))
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            tr("管理独立图片及 PPTX / DOCX / 数字版 PDF 内实际引用的图片，不做图片回填")
        )
        subtitle.setObjectName("pageSubtitle")
        title_stack.addWidget(title)
        title_stack.addWidget(subtitle)
        header_layout.addLayout(title_stack, 1)
        root.addWidget(header)

        project_bar = QFrame()
        project_bar.setObjectName("projectBar")
        project_layout = QHBoxLayout(project_bar)
        project_layout.setContentsMargins(12, 7, 8, 7)
        self.project_path = QLabel(tr("尚未打开图片库"))
        self.project_path.setObjectName("projectPath")
        project_layout.addWidget(self.project_path, 1)
        self.new_project_button = QPushButton(tr("新建图片库"))
        self.new_project_button.setFixedHeight(32)
        self.new_project_button.clicked.connect(self.choose_create_project)
        self.open_project_button = QPushButton(tr("切换 / 打开图片库"))
        self.open_project_button.setFixedHeight(32)
        self.open_project_button.clicked.connect(self.choose_open_project)
        project_layout.addWidget(self.new_project_button)
        project_layout.addWidget(self.open_project_button)
        root.addWidget(project_bar)

        workflow = QGroupBox(tr("图片入库工作流"))
        workflow_layout = QVBoxLayout(workflow)
        workflow_layout.setContentsMargins(10, 8, 10, 8)
        workflow_layout.setSpacing(6)
        source_row = QHBoxLayout()
        source_row.setSpacing(6)
        self.workflow_files = QFrame()
        self.workflow_files.setObjectName("workflowFiles")
        workflow_files_layout = QHBoxLayout(self.workflow_files)
        workflow_files_layout.setContentsMargins(6, 3, 6, 3)
        workflow_files_layout.setSpacing(6)
        self.import_drop_label = QLabel(tr("可拖入或多选图片、PPTX、DOCX、数字版 PDF"))
        self.import_drop_label.setObjectName("inputSummary")
        workflow_files_layout.addWidget(self.import_drop_label, 1)
        source_row.addWidget(self.workflow_files, 1)
        self.choose_import_button = QPushButton(tr("选择图片 / 文档（可多选）"))
        self.choose_import_button.setFixedHeight(32)
        self.choose_import_button.clicked.connect(self.choose_import)
        source_row.addWidget(self.choose_import_button)
        self.workflow_menu_button = QPushButton(tr("工作流设置"))
        self.workflow_menu_button.setFixedHeight(32)
        self.workflow_menu = QMenu(self.workflow_menu_button)
        self.workflow_menu.setToolTipsVisible(True)
        self.workflow_menu_button.setMenu(self.workflow_menu)
        source_row.addWidget(self.workflow_menu_button)
        workflow_layout.addLayout(source_row)

        self.workflow_settings = QFrame()
        self.workflow_settings.setObjectName("workflowSettings")
        self.workflow_settings.setMinimumHeight(38)
        settings_row = QHBoxLayout(self.workflow_settings)
        settings_row.setContentsMargins(8, 4, 8, 4)
        settings_row.setSpacing(6)
        settings_row.addWidget(QLabel(tr("默认分类")))
        self.import_category_input = QLineEdit()
        self.import_category_input.setPlaceholderText(tr("例如：示例项目/产品图"))
        self.import_category_input.setMinimumHeight(32)
        settings_row.addWidget(self.import_category_input, 1)
        settings_row.addWidget(QLabel(tr("最小尺寸")))
        self.min_width_spin = QSpinBox()
        self.min_height_spin = QSpinBox()
        image_settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        for spinbox, key, prefix in (
            (self.min_width_spin, "image_manager/min_width", tr("宽 ")),
            (self.min_height_spin, "image_manager/min_height", tr("高 ")),
        ):
            spinbox.setRange(0, 10000)
            spinbox.setSpecialValueText(tr("不限"))
            spinbox.setPrefix(prefix)
            spinbox.setSuffix(" px")
            spinbox.setFixedSize(82, 32)
            spinbox.setValue(image_settings.value(key, 0, int))
            spinbox.valueChanged.connect(
                lambda value, setting_key=key: self._minimum_size_changed(
                    setting_key, value
                )
            )
            settings_row.addWidget(spinbox)
        self.clear_import_button = QPushButton(tr("清空"))
        self.clear_import_button.clicked.connect(self.clear_pending_import)
        self.preview_import_button = QPushButton(tr("预览提取内容"))
        self.preview_import_button.clicked.connect(self.preview_pending_import)
        self.import_button = QPushButton(tr("开始入库"))
        self.import_button.setObjectName("primaryAction")
        self.import_button.clicked.connect(self.import_pending)
        self.workflow_action_targets = []
        for label, button in (
            (tr("清空待入库"), self.clear_import_button),
            (tr("预览提取内容"), self.preview_import_button),
            (tr("开始入库"), self.import_button),
        ):
            action = self.workflow_menu.addAction(label)
            action.triggered.connect(button.click)
            self.workflow_action_targets.append((action, button))
        self.workflow_menu.addSeparator()
        settings_action = self.workflow_menu.addAction(tr("展开 / 收起入库设置"))
        settings_action.triggered.connect(self._toggle_workflow_settings)
        self.workflow_menu.aboutToShow.connect(self._sync_workflow_actions)
        self.workflow_settings.hide()
        workflow_layout.addWidget(self.workflow_settings)
        root.addWidget(workflow)

        box = QGroupBox(tr("图片资产"))
        box_layout = QVBoxLayout(box)
        box_layout.setContentsMargins(10, 12, 10, 10)
        box_layout.setSpacing(8)

        filter_row = QHBoxLayout()
        self.library_toolbar_layout = filter_row
        self.library_filter_input = QLineEdit()
        self.library_filter_input.setPlaceholderText(tr("筛选名称、分类、格式或来源"))
        self.library_filter_input.setClearButtonEnabled(True)
        self.library_filter_input.setMinimumWidth(220)
        self.library_filter_input.setMaximumWidth(320)
        self.library_filter_input.textChanged.connect(self.refresh_views)
        filter_row.addWidget(QLabel(tr("查找")))
        filter_row.addWidget(self.library_filter_input)
        self.health_filter_combo = QComboBox()
        for key, title in HEALTH_FILTER_TITLES.items():
            self.health_filter_combo.addItem(title, key)
        self.health_filter_combo.setToolTip(
            tr("按图片健康状态筛选；只影响当前列表，不修改图片库。")
        )
        self.health_filter_combo.setMinimumWidth(120)
        self.health_filter_combo.setMaximumWidth(160)
        self.health_filter_combo.currentIndexChanged.connect(
            lambda: self._filter_by_health(str(self.health_filter_combo.currentData()))
        )
        self.health_filter_combo.hide()
        filter_row.addWidget(self.health_filter_combo)
        self.health_filter_group = QButtonGroup(self)
        self.health_filter_group.setExclusive(True)
        self.health_filter_buttons: dict[str, QPushButton] = {}
        for key, title, tooltip in (
            ("all", tr("全部"), tr("显示图片库全部资产")),
            ("duplicate_origins", tr("重复来源"), tr("同一图片被多个文件或文档引用")),
            ("similar", tr("相似"), tr("存在尚未确认的代码相似候选")),
            ("undersized", tr("过小"), tr("低于当前入库最小尺寸设置")),
            ("no_origin", tr("无来源"), tr("图片资产没有任何来源记录")),
        ):
            button = QPushButton(f"{title} 0")
            button.setObjectName("statFilter")
            button.setCheckable(True)
            button.setFixedHeight(32)
            button.setToolTip(f"{tooltip}{tr('；点击只筛选，不修改图片库。')}")
            button.clicked.connect(
                lambda _checked=False, name=key: self._filter_by_health(name)
            )
            self.health_filter_group.addButton(button)
            self.health_filter_buttons[key] = button
            filter_row.addWidget(button)
        self.health_filter_buttons["all"].setChecked(True)
        filter_row.addStretch(1)
        self.pending_cleanup_button = QPushButton(tr("待清理 0"), box)
        self.pending_cleanup_button.setFixedHeight(32)
        self.pending_cleanup_button.setToolTip(
            tr("查看已移出图片库但尚未永久删除的文件，可还原或手动清空。")
        )
        self.pending_cleanup_button.clicked.connect(self.show_pending_cleanup)

        self.edit_button = QPushButton(tr("编辑信息"))
        self.edit_button.setToolTip(
            tr("修改所选图片的名称、分类、标签和说明；不改动图片内容。")
        )
        self.edit_button.clicked.connect(self.edit_selected)
        self.similar_button = QPushButton(tr("查找相似图"))
        self.similar_button.setToolTip(
            tr("按代码指纹列出相近候选；只供核对，不会自动合并。")
        )
        self.similar_button.clicked.connect(self.show_similar)
        self.ai_button = QPushButton(tr("AI 整理建议"), box)
        self.ai_button.setToolTip(
            tr("先用代码指纹筛选候选，再由 AI 建议命名、分类、合并组和主资源；")
            + tr("视觉模型可额外参考压缩预览。")
        )
        self.ai_button.clicked.connect(self.request_ai_suggestion)
        self.open_location_button = QPushButton(tr("打开位置"), box)
        self.open_location_button.setToolTip(
            tr("在系统文件管理器中打开所选图片所在目录。")
        )
        self.open_location_button.clicked.connect(self.open_selected_location)
        self.delete_button = QPushButton(tr("移除选中"), box)
        self.delete_button.setToolTip(
            tr("确认后从图片库移除所选记录并移入待清理目录；原始文档不受影响。")
        )
        self.delete_button.clicked.connect(self.remove_selected)
        self.cleanup_button = QPushButton(tr("清理未引用文件"), box)
        self.cleanup_button.setToolTip(
            tr("扫描图片库 images 目录；确认后把未被清单引用的文件移入待清理目录。")
        )
        self.cleanup_button.clicked.connect(self.cleanup_orphans)
        self.more_actions_button = QPushButton(tr("更多操作"))
        self.more_actions_button.setToolTip(tr("收纳移除和图片库维护等低频操作。"))
        self.more_actions_menu = QMenu(self.more_actions_button)
        self.more_actions_menu.setToolTipsVisible(True)
        self.more_actions_menu.addSection(tr("资源"))
        ai_action = self.more_actions_menu.addAction(tr("AI 整理建议"))
        ai_action.triggered.connect(self.ai_button.click)
        open_location_action = self.more_actions_menu.addAction(tr("打开位置"))
        open_location_action.triggered.connect(self.open_location_button.click)
        remove_action = self.more_actions_menu.addAction(tr("移除选中"))
        remove_action.triggered.connect(self.delete_button.click)
        self.more_actions_menu.addSection(tr("库维护"))
        health_action = self.more_actions_menu.addAction(tr("库体检"))
        health_action.triggered.connect(self.show_library_health)
        health_action.setToolTip(
            tr("只读检查媒体缺失、大小变化、未登记文件和待清理索引。")
        )
        self.health_action = health_action
        cleanup_action = self.more_actions_menu.addAction(tr("清理未引用文件"))
        cleanup_action.triggered.connect(self.cleanup_button.click)
        pending_action = self.more_actions_menu.addAction(tr("待清理 (0)"))
        pending_action.triggered.connect(self.show_pending_cleanup)
        pending_action.setToolTip(tr("还原或永久清空已移出图片库的文件。"))
        self.pending_cleanup_action = pending_action
        reset_ignored_action = self.more_actions_menu.addAction(tr("重置已忽略候选"))
        reset_ignored_action.triggered.connect(self.reset_ignored_similar)
        reset_ignored_action.setToolTip(tr("重新显示曾确认“不是同一张”的相似候选。"))
        self.reset_ignored_action = reset_ignored_action
        self.more_actions_button.setMenu(self.more_actions_menu)
        self.more_action_targets = (
            (ai_action, self.ai_button),
            (open_location_action, self.open_location_button),
            (remove_action, self.delete_button),
            (cleanup_action, self.cleanup_button),
        )
        self.more_actions_menu.aboutToShow.connect(self._sync_more_actions)
        for button in (
            self.edit_button,
            self.similar_button,
            self.ai_button,
            self.open_location_button,
            self.more_actions_button,
        ):
            button.setFixedHeight(32)
        for button in (
            self.ai_button,
            self.open_location_button,
            self.pending_cleanup_button,
            self.delete_button,
            self.cleanup_button,
        ):
            button.hide()
        filter_row.addWidget(self.edit_button)
        filter_row.addWidget(self.similar_button)
        filter_row.addWidget(self.more_actions_button)
        box_layout.addLayout(filter_row)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        list_panel = QWidget()
        list_layout = QVBoxLayout(list_panel)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(0)
        self.library_empty = QLabel(
            tr("先新建或打开一个图片库\n\n可导入图片、PPTX、DOCX 或数字版 PDF")
        )
        self.library_empty.setObjectName("emptyState")
        self.library_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.library_empty.setWordWrap(True)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(
            [
                tr("名称"),
                tr("分类"),
                tr("规格"),
                tr("格式"),
                tr("大小"),
                tr("来源"),
                tr("相似"),
            ]
        )
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemSelectionChanged.connect(self.refresh_detail)
        self.tree.setColumnWidth(0, 260)
        self.tree.setColumnWidth(1, 130)
        self.tree.setColumnWidth(2, 110)
        self.tree.setColumnWidth(3, 70)
        self.tree.setColumnWidth(4, 90)
        self.tree.setColumnWidth(5, 100)
        self.tree.setColumnWidth(6, 60)
        list_layout.addWidget(self.library_empty, 1)
        list_layout.addWidget(self.tree, 1)
        splitter.addWidget(list_panel)

        detail = QFrame()
        detail.setObjectName("imageDetailDrawer")
        self.detail_drawer = detail
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(12, 12, 12, 12)
        self.detail_title = QLabel(tr("选择图片查看详情"))
        self.detail_title.setObjectName("detailTitle")
        self.preview = QLabel(tr("暂无预览"))
        self.preview.setObjectName("imagePreview")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setMinimumSize(280, 240)
        self.preview.setScaledContents(False)
        self.detail_meta = QLabel("")
        self.detail_meta.setObjectName("detailMeta")
        self.detail_meta.setWordWrap(True)
        self.detail_summary = QLabel("")
        self.detail_summary.setObjectName("detailValue")
        self.detail_summary.setWordWrap(True)
        self.detail_origins = QTreeWidget()
        self.detail_origins.setHeaderLabels([tr("来源"), tr("文档内位置")])
        self.detail_origins.setRootIsDecorated(False)
        self.detail_origins.setMaximumHeight(110)
        detail_layout.addWidget(self.detail_title)
        detail_layout.addWidget(self.preview, 1)
        detail_layout.addWidget(self.detail_meta)
        detail_layout.addWidget(self.detail_summary)
        detail_layout.addWidget(self.detail_origins)
        splitter.addWidget(detail)
        splitter.setSizes([820, 360])
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        box_layout.addWidget(splitter, 1)
        root.addWidget(box, 1)

        self.status_label = QLabel(tr("状态与日志 · 等待开始"))
        self.status_label.setObjectName("projectPath")
        root.addWidget(self.status_label)
        self.setCentralWidget(central)
        self._sync_actions()

    def on_activated(self) -> None:
        if self.project is None:
            settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
            last_project = settings.value("image_manager/last_project", "", str)
            if last_project and (Path(last_project) / "image-project.json").is_file():
                self.open_project(Path(last_project), report_errors=False)
            return
        if self.worker_thread is not None:
            return
        try:
            self.project.reload()
        except Exception as exc:
            self.status_label.setText(f"{tr('状态与日志 · 重新载入失败：')}{exc}")
            return
        self.refresh_views()

    def on_settings_changed(self) -> None:
        self._sync_actions()

    def choose_create_project(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, tr("新建图片库"))
        if directory:
            try:
                project = ImageProject.create(Path(directory), tr("图片库"))
            except Exception as exc:
                self.show_error(str(exc))
                return
            self._set_project(project)

    def choose_open_project(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, tr("打开图片库"))
        if directory:
            self.open_project(Path(directory))

    def open_project(self, root: Path, *, report_errors: bool = True) -> None:
        try:
            project = ImageProject.open(root)
        except Exception as exc:
            if report_errors:
                self.show_error(str(exc))
            return
        self._set_project(project)

    def _set_project(self, project: ImageProject) -> None:
        self.project = project
        QSettings("Doc Media Toolkit", "Doc Media Toolkit").setValue(
            "image_manager/last_project", str(project.root)
        )
        self.project_path.setText(
            f"{tr('当前图片库：')}{project.data['name']}{tr(' · ')}{project.root}"
        )
        self.status_label.setText(
            tr("状态与日志 · 图片库已打开")
            + (tr("，已从最近有效备份恢复") if project.recovered_from_backup else "")
        )
        if project.recovered_from_backup:
            QMessageBox.warning(
                self,
                tr("图片库已恢复"),
                tr("当前清单不可读，已从最近有效备份恢复。建议立即执行“库体检”。"),
            )
        self.refresh_views()

    def _minimum_size_changed(self, key: str, value: int) -> None:
        QSettings("Doc Media Toolkit", "Doc Media Toolkit").setValue(key, value)
        if self.project is not None:
            self.refresh_views()

    def choose_import(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            tr("导入图片或文档"),
            "",
            (
                tr("支持的文件 (*.png *.jpg *.jpeg *.webp *.gif *.bmp *.tif *.tiff ")
                + tr("*.pptx *.docx *.pdf);;所有文件 (*)")
            ),
        )
        if files:
            self.add_pending_paths([Path(path) for path in files])

    @property
    def input_paths(self) -> list[Path]:
        return self.pending_paths

    def set_files(self, paths: list[Path]) -> None:
        self.add_pending_paths(paths)

    def add_pending_paths(self, paths: list[Path]) -> None:
        supported = {
            ".png",
            ".jpg",
            ".jpeg",
            ".webp",
            ".gif",
            ".bmp",
            ".tif",
            ".tiff",
            ".pptx",
            ".docx",
            ".pdf",
        }
        known = {path.resolve() for path in self.pending_paths}
        for path in paths:
            resolved = path.expanduser().resolve()
            if resolved.is_file() and resolved.suffix.lower() in supported:
                if resolved not in known:
                    self.pending_paths.append(resolved)
                    known.add(resolved)
        self._refresh_pending_import()

    def _refresh_pending_import(self) -> None:
        if self.pending_paths:
            names = tr("、").join(path.name for path in self.pending_paths[:3])
            more = (
                f"{tr(' 等 ')}{len(self.pending_paths)}{tr(' 个文件')}"
                if len(self.pending_paths) > 3
                else ""
            )
            self.import_drop_label.setText(f"{tr('待入库：')}{names}{more}")
        else:
            self.import_drop_label.setText(
                tr("可拖入或多选图片、PPTX、DOCX、数字版 PDF")
            )
        self._sync_actions()

    def clear_pending_import(self) -> None:
        self.pending_paths.clear()
        self._refresh_pending_import()

    def preview_pending_import(self) -> None:
        if self.pending_paths:
            ImportPreviewDialog(
                self,
                self.pending_paths,
                min_width=self.min_width_spin.value(),
                min_height=self.min_height_spin.value(),
            ).exec()

    def import_pending(self) -> None:
        if not self.pending_paths:
            self.choose_import()
            return
        self.import_paths(list(self.pending_paths))

    def import_paths(self, paths: list[Path]) -> None:
        if self.project is None or self.worker_thread is not None:
            self.show_error(tr("请先新建或打开图片库。"))
            return
        project = self.project
        category = self.import_category_input.text().strip()
        min_width = self.min_width_spin.value()
        min_height = self.min_height_spin.value()
        self.status_label.setText(tr("状态与日志 · 正在提取并整理图片…"))
        self.worker_thread = QThread(self)
        self.worker = OperationWorker(
            lambda _message, cancelled: project.import_paths(
                paths,
                category=category,
                min_width=min_width,
                min_height=min_height,
                cancel_callback=cancelled,
            )
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._import_succeeded)
        self.worker.failed.connect(self._import_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self._import_thread_finished)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()
        self._sync_actions()

    def _import_succeeded(self, value: object) -> None:
        if not isinstance(value, dict):
            self._import_failed(tr("图片入库返回了无效结果。"))
            return
        result = value
        failed = len(result["failed"])
        skipped = len(result["skipped"])
        self.status_label.setText(
            f"{tr('状态与日志 · 新增 ')}{result['added']}{tr('，复用 ')}{result['reused']}{tr('，')}"
            + f"{tr('排除 ')}{skipped}{tr('，失败 ')}{failed}"
            + (tr("，已取消后续文件") if result.get("cancelled") else "")
        )
        self.refresh_views()
        if not failed and not result.get("cancelled"):
            self.clear_pending_import()
        if failed:
            details = "\n".join(
                f"{item['source']} {item['member']}: {item['error']}"
                for item in result["failed"][:20]
            )
            QMessageBox.warning(self, tr("部分图片未导入"), details)

    def _import_failed(self, message: str) -> None:
        self.status_label.setText(tr("状态与日志 · 图片入库失败"))
        self.show_error(message.splitlines()[0])

    def _import_thread_finished(self) -> None:
        self.worker = None
        self.worker_thread = None
        self._sync_actions()

    def refresh_views(self) -> None:
        selected_id = self.selected_asset_id()
        self.tree.clear()
        if self.project is None:
            for index, (key, button) in enumerate(self.health_filter_buttons.items()):
                button.setText(f"{HEALTH_FILTER_TITLES[key]} 0")
                self.health_filter_combo.setItemText(
                    index, f"{HEALTH_FILTER_TITLES[key]} 0"
                )
            self.pending_cleanup_button.setText(tr("待清理 0"))
            self.pending_cleanup_action.setText(tr("待清理 (0)"))
            self.library_empty.setText(
                tr("先新建或打开一个图片库\n\n可导入图片、PPTX、DOCX 或数字版 PDF")
            )
            self.library_empty.show()
            self.tree.hide()
            self.detail_drawer.hide()
            self._sync_actions()
            self.refresh_detail()
            return
        counts = self.project.health_counts(
            min_width=self.min_width_spin.value(),
            min_height=self.min_height_spin.value(),
        )
        for index, (key, button) in enumerate(self.health_filter_buttons.items()):
            button.setText(f"{HEALTH_FILTER_TITLES[key]} {counts[key]}")
            self.health_filter_combo.setItemText(
                index, f"{HEALTH_FILTER_TITLES[key]} {counts[key]}"
            )
        pending = self.project.pending_cleanup()
        pending_size = sum(item["size_bytes"] for item in pending)
        self.pending_cleanup_button.setText(f"{tr('待清理 ')}{len(pending)}")
        self.pending_cleanup_action.setText(f"{tr('待清理 (')}{len(pending)})")
        self.pending_cleanup_button.setToolTip(
            f"{tr('当前待清理 ')}{len(pending)}{tr(' 个文件，')}"
            + f"{format_user_file_size(pending_size)}{tr('；可还原或手动永久清空。')}"
        )
        needle = self.library_filter_input.text().strip().casefold()
        for asset in sorted(
            self.project.assets(), key=lambda item: item["name"].casefold()
        ):
            source_types = sorted(
                {str(item.get("source_type") or "") for item in asset["origins"]}
            )
            searchable = " ".join(
                (
                    str(asset.get("name") or ""),
                    str(asset.get("category") or ""),
                    str(asset.get("format") or ""),
                    " ".join(source_types),
                    str(asset.get("sha256") or ""),
                )
            ).casefold()
            if needle and needle not in searchable:
                continue
            health = self.project.asset_health(
                asset["id"],
                min_width=self.min_width_spin.value(),
                min_height=self.min_height_spin.value(),
            )
            if self.health_filter != "all" and not health[self.health_filter]:
                continue
            candidate_count = (
                str(len(self.project.similar_candidates(asset["id"])))
                if asset["id"] == selected_id
                else tr("按需")
            )
            item = QTreeWidgetItem(
                [
                    asset["name"],
                    asset.get("category") or tr("未分类"),
                    f"{asset['width']}×{asset['height']}",
                    asset["format"],
                    format_user_file_size(asset["size_bytes"]),
                    "/".join(source_types),
                    candidate_count,
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, asset["id"])
            self.tree.addTopLevelItem(item)
            if asset["id"] == selected_id:
                self.tree.setCurrentItem(item)
        visible_assets = self.tree.topLevelItemCount()
        self.tree.setVisible(visible_assets > 0)
        self.detail_drawer.setVisible(visible_assets > 0)
        self.library_empty.setVisible(visible_assets == 0)
        if visible_assets == 0:
            self.library_empty.setText(
                tr("没有匹配的图片资产")
                if needle or self.health_filter != "all"
                else tr("图片库还是空的\n\n选择图片或文档后即可预览并入库")
            )
        self._sync_actions()
        self.refresh_detail()

    def selected_asset_id(self) -> str | None:
        item = self.tree.currentItem()
        return str(item.data(0, Qt.ItemDataRole.UserRole)) if item is not None else None

    def selected_asset(self) -> dict | None:
        if self.project is None:
            return None
        asset_id = self.selected_asset_id()
        if asset_id is None:
            return None
        try:
            return self.project.asset(asset_id)
        except KeyError:
            return None

    def _sync_actions(self) -> None:
        busy = self.worker_thread is not None
        has_project = self.project is not None
        has_selection = self.selected_asset() is not None
        has_pending = bool(self.pending_paths)
        self.import_button.setEnabled(has_project and has_pending and not busy)
        self.preview_import_button.setEnabled(has_pending and not busy)
        self.clear_import_button.setEnabled(has_pending and not busy)
        self.choose_import_button.setEnabled(not busy)
        self.new_project_button.setEnabled(not busy)
        self.open_project_button.setEnabled(not busy)
        for button in (
            self.edit_button,
            self.similar_button,
            self.open_location_button,
            self.delete_button,
        ):
            button.setEnabled(has_selection and not busy)
        ai_configured = self._ai_config() is not None
        self.ai_button.setEnabled(has_selection and ai_configured and not busy)
        self.ai_button.setToolTip(
            (
                tr("先用代码指纹筛选候选，再由 AI 建议命名、分类、合并组和主资源；")
                + tr("视觉模型可额外参考压缩预览。")
            )
            if ai_configured
            else tr("请先点击顶栏齿轮配置并验证 AI；未配置时图片库功能不受影响。")
        )
        self.cleanup_button.setEnabled(has_project and not busy)
        self.library_filter_input.setEnabled(has_project and not busy)
        self.pending_cleanup_button.setEnabled(has_project and not busy)
        for button in self.health_filter_buttons.values():
            button.setEnabled(has_project and not busy)

    def _sync_more_actions(self) -> None:
        for action, button in self.more_action_targets:
            action.setEnabled(button.isEnabled())
            action.setToolTip(button.toolTip())
        self.pending_cleanup_action.setEnabled(
            self.project is not None and self.worker_thread is None
        )
        self.reset_ignored_action.setEnabled(
            self.project is not None
            and bool(self.project.data.get("ignored_similar_pairs"))
            and self.worker_thread is None
        )
        self.health_action.setEnabled(
            self.project is not None and self.worker_thread is None
        )

    def _sync_workflow_actions(self) -> None:
        for action, button in self.workflow_action_targets:
            action.setEnabled(button.isEnabled())
            action.setToolTip(button.toolTip())

    def _toggle_workflow_settings(self) -> None:
        self.workflow_settings.setVisible(self.workflow_settings.isHidden())

    def _filter_by_health(self, name: str) -> None:
        self.health_filter = name
        self.health_filter_buttons[name].setChecked(True)
        index = self.health_filter_combo.findData(name)
        if index >= 0 and index != self.health_filter_combo.currentIndex():
            self.health_filter_combo.blockSignals(True)
            self.health_filter_combo.setCurrentIndex(index)
            self.health_filter_combo.blockSignals(False)
        self.refresh_views()

    def update_responsive_layout(self, width: int) -> None:
        compact = width < 1180
        self.library_filter_input.setMinimumWidth(120 if compact else 220)
        self.library_filter_input.setMaximumWidth(180 if compact else 320)
        self.health_filter_combo.setVisible(compact)
        for button in self.health_filter_buttons.values():
            button.setVisible(not compact)

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

    def refresh_detail(self) -> None:
        self._sync_actions()
        asset = self.selected_asset()
        if asset is None or self.project is None:
            self.detail_title.setText(tr("选择图片查看详情"))
            self.preview.setPixmap(QPixmap())
            self.preview.setText(tr("暂无预览"))
            self.detail_meta.clear()
            self.detail_summary.clear()
            self.detail_origins.clear()
            return
        self.detail_title.setText(asset["name"])
        self.preview.setText("")
        pixmap = self._load_preview(self.project.asset_path(asset))
        if pixmap.isNull():
            self.preview.setText(tr("无法预览"))
        else:
            self.preview.setPixmap(
                pixmap.scaled(
                    self.preview.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        origins = asset.get("origins") or []
        self.detail_meta.setText(
            f"{asset['width']}×{asset['height']}{tr(' · ')}{asset['format']}{tr(' · ')}"
            + f"{format_user_file_size(asset['size_bytes'])}\n"
            + f"{tr('SHA-256：')}{asset['sha256'][:16]}{tr('… · 来源 ')}{len(origins)}{tr(' 处')}\n"
            + f"{tr('分类：')}{asset.get('category') or tr('未分类')}{tr(' · ')}"
            + f"{tr('标签：')}{tr('、').join(asset.get('tags') or []) or tr('无')}"
        )
        self.detail_summary.setText(asset.get("summary") or tr("暂无说明。"))
        self.detail_origins.clear()
        for origin in origins:
            source_path = str(origin.get("source_path") or tr("未知来源"))
            source_item = QTreeWidgetItem(
                [
                    source_path,
                    str(origin.get("member_path") or tr("独立图片")),
                ]
            )
            source_item.setToolTip(0, source_path)
            self.detail_origins.addTopLevelItem(source_item)
        self.detail_origins.setVisible(bool(origins))

    @staticmethod
    def _load_preview(path: Path) -> QPixmap:
        try:
            return MainWindow._load_preview_bytes(path.read_bytes())
        except Exception:
            return QPixmap()

    @staticmethod
    def _load_preview_bytes(data: bytes) -> QPixmap:
        try:
            with Image.open(io.BytesIO(data)) as image:
                image = ImageOps.exif_transpose(image).convert("RGBA")
                pixels = image.tobytes("raw", "RGBA")
                qimage = QImage(
                    pixels,
                    image.width,
                    image.height,
                    image.width * 4,
                    QImage.Format.Format_RGBA8888,
                )
                return QPixmap.fromImage(qimage.copy())
        except Exception:
            return QPixmap()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self.update_responsive_layout(event.size().width())
        if self.selected_asset() is not None:
            self.refresh_detail()

    def edit_selected(self) -> None:
        asset = self.selected_asset()
        if asset is None or self.project is None:
            return
        dialog = MetadataDialog(self, asset)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            self.project.update_metadata(asset["id"], **dialog.values())
        except Exception as exc:
            self.show_error(str(exc))
            return
        self.refresh_views()

    def show_similar(self) -> None:
        asset = self.selected_asset()
        if asset is None or self.project is None:
            return
        current_id = asset["id"]
        reviewed = 0
        while True:
            try:
                current = self.project.asset(current_id)
            except KeyError:
                break
            candidates = self.project.similar_candidates(current_id)
            if not candidates:
                break
            candidate_info = candidates[0]
            candidate = self.project.asset(candidate_info["asset_id"])
            dialog = SimilarImageReviewDialog(
                self,
                self.project,
                current,
                candidate,
                candidate_info["score"],
            )
            if dialog.exec() != QDialog.DialogCode.Accepted:
                break
            try:
                if dialog.decision == "ignore":
                    self.project.ignore_similar_pair(current_id, candidate["id"])
                else:
                    primary_id = (
                        current_id if dialog.decision == "left" else candidate["id"]
                    )
                    self.project.merge_assets(
                        primary_id,
                        [current_id, candidate["id"]],
                        confirmed_same_content=True,
                    )
                    current_id = primary_id
            except Exception as exc:
                self.show_error(str(exc))
                break
            reviewed += 1
            self.refresh_views()
        if reviewed:
            self.status_label.setText(
                f"{tr('状态与日志 · 已核对 ')}{reviewed}{tr(' 组相似图片')}"
            )
            self.refresh_views()
            return
        if not self.project.similar_candidates(current_id):
            QMessageBox.information(
                self, tr("相似候选"), tr("代码指纹没有找到足够接近的候选。")
            )

    def reset_ignored_similar(self) -> None:
        if self.project is None:
            return
        try:
            count = self.project.reset_ignored_similar_pairs()
        except Exception as exc:
            self.show_error(str(exc))
            return
        self.status_label.setText(
            f"{tr('状态与日志 · 已重置 ')}{count}{tr(' 组忽略的相似候选')}"
        )
        self.refresh_views()

    def request_ai_suggestion(self) -> None:
        from pptx_tools.ai_client import OpenAICompatibleClient, privacy_scope

        if self.ai_thread is not None:
            self.ai_ignore_result = True
            if self.ai_worker is not None:
                self.ai_worker.cancel()
            self.ai_button.setText(tr("已取消显示"))
            self.status_label.setText(
                tr("状态与日志 · 已取消显示；网络请求会在后台结束")
            )
            return

        asset = self.selected_asset()
        config = self._ai_config()
        if asset is None or self.project is None or config is None:
            self.show_error(tr("请先在顶栏设置中配置支持图片输入的 AI。"))
            return
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        privacy_key = f"ai/privacy_confirmed/{privacy_scope(config)}"
        if not settings.value(privacy_key, False, bool):
            sent_content = (
                tr("压缩预览、名称、规格与大小")
                if config.vision_enabled
                else tr("名称、规格、大小与代码相似度（不发送图片）")
            )
            answer = QMessageBox.question(
                self,
                tr("AI 图片分析"),
                (
                    tr("将向已配置的 AI 服务发送所选图片和最多 5 张代码相似候选的")
                    + f"{sent_content}{tr('，不发送原文档或本地路径。是否继续？')}"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            settings.setValue(privacy_key, True)
        candidates = self.project.similar_candidates(asset["id"], limit=5)
        ids = [asset["id"]] + [item["asset_id"] for item in candidates]
        candidate_scores = {item["asset_id"]: item["score"] for item in candidates}
        items = []
        self.ai_item_names = {}
        for item_id in ids:
            item = self.project.asset(item_id)
            self.ai_item_names[item_id] = item["name"]
            items.append(
                {
                    **item,
                    "preview_path": self.project.asset_path(item),
                    "code_similarity": candidate_scores.get(item_id, 100.0),
                    "health": tr("正常"),
                }
            )
        self.ai_target_asset_id = asset["id"]
        self.ai_ignore_result = False
        self.ai_request_config = config
        self.ai_button.setEnabled(True)
        self.ai_button.setText(tr("取消 AI 分析"))
        self.status_label.setText(tr("状态与日志 · AI 正在分析图片与相似候选…"))
        self.ai_thread = QThread(self)
        self.ai_worker = OperationWorker(
            lambda _message, cancelled: OpenAICompatibleClient(config).organize_media(
                items, media_kind=tr("图片"), cancel_callback=cancelled
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
            self._finish_ai()
            return
        try:
            asset = self.project.asset(self.ai_target_asset_id)
        except KeyError:
            self.status_label.setText(
                tr("状态与日志 · 对应图片已不存在，AI 结果已忽略")
            )
            self._finish_ai()
            return
        from pptx_tools.ai_review_dialog import AISuggestionDialog

        dialog = AISuggestionDialog(
            self,
            result,
            {
                "suggested_name": asset["name"],
                "category": asset.get("category", ""),
                "tags": asset.get("tags", []),
                "summary": asset.get("summary", ""),
            },
            {
                "suggested_name": tr("名称"),
                "category": tr("分类"),
                "tags": tr("标签"),
                "summary": tr("说明"),
            },
            self.ai_item_names,
        )
        applied: list[str] = []
        if dialog.exec() == QDialog.DialogCode.Accepted:
            values = dialog.selected_values()
            try:
                self.project.update_metadata(
                    asset["id"],
                    name=values.get("suggested_name"),
                    category=values.get("category"),
                    tags=values.get("tags"),
                    summary=values.get("summary"),
                )
                applied = [
                    label
                    for key, label in {
                        "suggested_name": tr("名称"),
                        "category": tr("分类"),
                        "tags": tr("标签"),
                        "summary": tr("说明"),
                    }.items()
                    if key in values
                ]
            except Exception as exc:
                self.show_error(str(exc))
        merged = self._review_ai_merge_groups(result.get("merge_groups") or [])
        if merged:
            applied.append(f"{tr('合并 ')}{merged}{tr(' 张重复图片')}")
        self.status_label.setText(
            tr("状态与日志 · AI 建议已完成")
            + (
                f"{tr('，已应用：')}{tr('、').join(applied)}"
                if applied
                else tr("，未修改图片库")
            )
        )
        if self.ai_request_config is not None:
            from pptx_tools.app_logging import write_ai_audit_event

            write_ai_audit_event(
                media_kind="image",
                target_id=self.ai_target_asset_id,
                provider=self.ai_request_config.base_url,
                model=self.ai_request_config.model,
                vision_enabled=self.ai_request_config.vision_enabled,
                applied_fields=applied,
                merge_group_count=len(result.get("merge_groups") or []),
            )
        self._finish_ai()
        self.refresh_views()

    def _review_ai_merge_groups(self, groups: list[dict]) -> int:
        if self.project is None:
            return 0
        merged = 0
        for group in groups:
            try:
                assets = [self.project.asset(item_id) for item_id in group["item_ids"]]
                primary = self.project.asset(group["primary_id"])
            except (KeyError, TypeError):
                continue
            names = tr("、").join(asset["name"] for asset in assets)
            answer = QMessageBox.question(
                self,
                tr("核对 AI 图片合并建议"),
                (
                    f"{tr('疑似同一图片：')}{names}\n"
                    + f"{tr('建议保留：')}{primary['name']}\n"
                    + f"{tr('置信度：')}{float(group.get('confidence') or 0):.0%}\n"
                    + f"{tr('理由：')}{group.get('reason') or tr('未说明')}\n\n"
                    + tr("确认后会合并来源、标签和说明，并删除库内其他副本；")
                    + tr("原始文档和原始图片不受影响。是否确认它们内容相同？")
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                continue
            try:
                result = self.project.merge_assets(
                    primary["id"],
                    [asset["id"] for asset in assets],
                    confirmed_same_content=True,
                )
            except Exception as exc:
                self.show_error(str(exc))
                continue
            merged += len(result["removed_ids"])
        return merged

    def _ai_failed(self, message: str) -> None:
        if self.ai_ignore_result:
            return
        self.status_label.setText(tr("状态与日志 · AI 分析失败"))
        self._finish_ai()
        self.show_error(message)

    def _finish_ai(self) -> None:
        self.ai_button.setEnabled(False)

    def _ai_thread_finished(self) -> None:
        self.ai_worker = None
        self.ai_thread = None
        self.ai_button.setText(tr("AI 整理建议"))
        self._sync_actions()

    def open_selected_location(self) -> None:
        asset = self.selected_asset()
        if asset is None or self.project is None:
            return
        QDesktopServices.openUrl(
            QUrl.fromLocalFile(str(self.project.asset_path(asset).parent))
        )

    def remove_selected(self) -> None:
        asset = self.selected_asset()
        if asset is None or self.project is None:
            return
        answer = QMessageBox.question(
            self,
            tr("移除图片"),
            f"{tr('将从图片库移除“')}{asset['name']}{tr('”并移入待清理目录。')}"
            + tr("原始文档不受影响，是否继续？"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.project.remove_asset(asset["id"])
        except Exception as exc:
            self.show_error(str(exc))
            return
        self.status_label.setText(f"{tr('状态与日志 · 已移除 ')}{asset['name']}")
        self.refresh_views()

    def cleanup_orphans(self) -> None:
        if self.project is None:
            return
        orphans = self.project.orphan_paths()
        if not orphans:
            self.status_label.setText(tr("状态与日志 · 没有未引用文件"))
            return
        answer = QMessageBox.question(
            self,
            tr("清理未引用文件"),
            (
                f"{tr('发现 ')}{len(orphans)}{tr(' 个不在图片库清单中的文件。')}\n"
                + tr("这些文件将移入待清理目录，可在“更多操作”中还原；")
                + tr("原始文档不受影响。是否继续？")
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        removed = self.project.cleanup_orphans()
        self.status_label.setText(
            f"{tr('状态与日志 · 已将 ')}{len(removed)}{tr(' 个未引用文件移入待清理')}"
        )

    def show_pending_cleanup(self) -> None:
        PendingImageCleanupDialog(self, self).exec()
        self.refresh_views()

    def show_library_health(self) -> None:
        if self.project is None:
            return
        report = self.project.health_report()
        issue_count = sum(
            len(report[key])
            for key in (
                "missing_files",
                "modified_files",
                "orphan_files",
                "pending_cleanup_issues",
            )
        )
        details = [
            f"{tr('图片资产：')}{report['asset_count']}",
            f"{tr('待清理：')}{report['pending_cleanup_count']}",
            f"{tr('缺失文件：')}{len(report['missing_files'])}",
            f"{tr('内容大小变化：')}{len(report['modified_files'])}",
            f"{tr('未登记文件：')}{len(report['orphan_files'])}",
            f"{tr('待清理索引异常：')}{len(report['pending_cleanup_issues'])}",
        ]
        QMessageBox.information(
            self,
            tr("图片库体检"),
            (
                tr("未发现异常。\n\n")
                if not issue_count
                else tr("发现需要处理的项目。\n\n")
            )
            + "\n".join(details),
        )
        self.status_label.setText(
            tr("状态与日志 · 图片库体检完成，")
            + (
                tr("未发现异常")
                if not issue_count
                else f"{tr('发现 ')}{issue_count}{tr(' 项异常')}"
            )
        )

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        paths = [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile() and Path(url.toLocalFile()).is_file()
        ]
        if paths:
            self.add_pending_paths(paths)
            event.acceptProposedAction()

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, tr("图片库"), message)

    def closeEvent(self, event) -> None:  # noqa: N802, ANN001
        if self.ai_thread is not None and self.ai_thread.isRunning():
            self.ai_ignore_result = True
            if self.ai_worker is not None:
                self.ai_worker.cancel()
            event.ignore()
            self.hide()
            self.ai_thread.finished.connect(self.close)
            return
        if self.worker_thread is not None and self.worker_thread.isRunning():
            if self.worker is not None:
                self.worker.cancel()
            event.ignore()
            self.hide()
            self.worker_thread.finished.connect(self.close)
            return
        super().closeEvent(event)


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
