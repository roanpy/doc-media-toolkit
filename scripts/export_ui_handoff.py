#!/usr/bin/env python3

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
for candidate in (REPO_ROOT, SRC_ROOT):
    candidate_str = str(candidate)
    if candidate_str not in sys.path:
        sys.path.insert(0, candidate_str)

from PySide6.QtCore import QSettings, Qt  # noqa: E402
from PySide6.QtGui import QColor, QImage, QPainter  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

PRESET_COLORS = None
PREVIEW_BACKGROUND_THEMES = None
WATERMARK_PRESETS = None
PreviewArtifacts = None
ToolboxWindow = None
SHELL_STRINGS = None
MEDIA_STRINGS = None
WATERMARK_STRINGS = None
PRESET_OPTIONS = None


DESKTOP_DIR = Path.home() / "Desktop" / "PPT-Tools-UI-Handoff"
SCREENSHOT_DIR = DESKTOP_DIR / "screenshots"
DOC_DIR = DESKTOP_DIR / "docs"
DATA_DIR = DESKTOP_DIR / "data"
MOCK_DIR = DATA_DIR / "mock_files"
DEFAULT_SIZE = (1440, 900)


def ensure_dirs() -> None:
    for path in (DESKTOP_DIR, SCREENSHOT_DIR, DOC_DIR, DATA_DIR, MOCK_DIR):
        path.mkdir(parents=True, exist_ok=True)
    shutil.copy2(REPO_ROOT / "docs" / "UI_DESIGN.md", DOC_DIR / "UI_DESIGN.md")
    shutil.copy2(REPO_ROOT / "design-qa.md", DOC_DIR / "design-qa.md")


def app_instance() -> QApplication:
    settings_dir = DATA_DIR / "qt-settings"
    settings_dir.mkdir(parents=True, exist_ok=True)
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        str(settings_dir),
    )
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def load_ui_modules() -> None:
    global PRESET_COLORS
    global PREVIEW_BACKGROUND_THEMES
    global WATERMARK_PRESETS
    global PreviewArtifacts
    global ToolboxWindow
    global SHELL_STRINGS
    global MEDIA_STRINGS
    global WATERMARK_STRINGS
    global PRESET_OPTIONS

    if ToolboxWindow is not None:
        return

    from src.pptx_output_watermark.gui import (
        PRESET_COLORS as _PRESET_COLORS,
        PREVIEW_BACKGROUND_THEMES as _PREVIEW_BACKGROUND_THEMES,
        PreviewArtifacts as _PreviewArtifacts,
        STRINGS as _WATERMARK_STRINGS,
        WATERMARK_PRESETS as _WATERMARK_PRESETS,
    )
    from src.pptx_tools.gui import MainWindow as _ToolboxWindow
    from src.pptx_tools.gui import STRINGS as _SHELL_STRINGS
    from src.pptx_video_compactor_gui import PRESET_OPTIONS as _PRESET_OPTIONS
    from src.pptx_video_compactor_gui import STRINGS as _MEDIA_STRINGS

    PRESET_COLORS = _PRESET_COLORS
    PREVIEW_BACKGROUND_THEMES = _PREVIEW_BACKGROUND_THEMES
    WATERMARK_PRESETS = _WATERMARK_PRESETS
    PreviewArtifacts = _PreviewArtifacts
    ToolboxWindow = _ToolboxWindow
    SHELL_STRINGS = _SHELL_STRINGS
    MEDIA_STRINGS = _MEDIA_STRINGS
    WATERMARK_STRINGS = _WATERMARK_STRINGS
    PRESET_OPTIONS = _PRESET_OPTIONS


def process_events(app: QApplication, rounds: int = 8) -> None:
    for _ in range(rounds):
        app.processEvents()


def make_sample_png() -> Path:
    path = DATA_DIR / "sample_watermark_image.png"
    image = QImage(640, 360, QImage.Format.Format_ARGB32)
    image.fill(QColor("#334155"))
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#f97316"))
    painter.drawRoundedRect(28, 26, 92, 28, 10, 10)
    painter.setBrush(QColor("#e2e8f0"))
    painter.drawRoundedRect(28, 78, 260, 18, 9, 9)
    painter.setBrush(QColor("#64748b"))
    painter.drawRoundedRect(28, 108, 210, 14, 7, 7)
    painter.setBrush(QColor("#2563eb"))
    painter.drawRoundedRect(170, 210, 100, 64, 12, 12)
    painter.setBrush(QColor("#facc15"))
    painter.drawRoundedRect(292, 220, 120, 52, 12, 12)
    painter.setBrush(QColor("#64748b"))
    painter.drawRoundedRect(396, 200, 200, 82, 12, 12)
    painter.end()
    image.save(str(path))
    return path


def make_sample_portrait_png() -> Path:
    path = DATA_DIR / "sample_portrait_page.png"
    image = QImage(640, 900, QImage.Format.Format_ARGB32)
    image.fill(QColor("#334155"))
    painter = QPainter(image)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#e2e8f0"))
    painter.drawRoundedRect(48, 72, 330, 22, 8, 8)
    painter.setBrush(QColor("#94a3b8"))
    painter.drawRoundedRect(48, 112, 260, 16, 7, 7)
    painter.setBrush(QColor("#2563eb"))
    painter.drawRoundedRect(66, 560, 210, 150, 16, 16)
    painter.setBrush(QColor("#64748b"))
    painter.drawRoundedRect(250, 560, 320, 150, 16, 16)
    painter.end()
    image.save(str(path))
    return path


def make_mock_files() -> dict[str, Path]:
    samples = {
        "pptx_alpha": MOCK_DIR / "sample_alpha.pptx",
        "pptx_beta": MOCK_DIR / "sample_beta.pptx",
        "docx": MOCK_DIR / "sample_beta.docx",
        "pdf": MOCK_DIR / "sample_gamma.pdf",
        "png": MOCK_DIR / "equipment_diagram.png",
        "mp4": MOCK_DIR / "production_demo.mp4",
    }
    payloads = {
        "pptx_alpha": b"PPTX sample alpha\n" * 256,
        "pptx_beta": b"PPTX sample beta\n" * 192,
        "docx": b"DOCX sample beta\n" * 96,
        "pdf": b"%PDF-1.7 sample gamma\n" * 72,
        "png": b"PNG sample\n" * 128,
        "mp4": b"MP4 sample\n" * 256,
    }
    for key, path in samples.items():
        path.write_bytes(payloads[key])
    return samples


def set_language(language: str) -> None:
    os.environ["PPTX_TOOLS_LANG"] = language
    os.environ["PPTX_OUTPUT_WATERMARK_LANG"] = language
    os.environ["PPTX_VIDEO_COMPACTOR_LANG"] = language


def choose_text(strings: dict[str, str], *keys: str) -> str:
    return max((strings[key] for key in keys), key=len)


def configure_watermark(
    toolbox: ToolboxWindow, language: str, sample_png: Path
) -> None:
    tool = toolbox.embedded_tools[0].window
    if tool is None:
        return
    text = tool.text
    format_value = "pptx"
    mode_value = "image"
    dpi_label = (
        "image_quality_balanced" if language == "zh" else "image_quality_original"
    )
    preset_key = "dense_review"
    color_label = "Cloud" if language == "en" else "Slate"
    preview_theme = "gray" if language == "zh" else "blue"

    tool.output_format_select.setCurrentIndex(
        tool.output_format_select.findData(format_value)
    )
    tool.output_mode_select.setCurrentIndex(
        tool.output_mode_select.findData(mode_value)
    )
    tool.dpi_select.setCurrentIndex(tool.dpi_select.findText(text[dpi_label]))
    tool.keep_videos_checkbox.setChecked(True)
    tool.update_image_quality_control()
    tool.keep_videos_checkbox.setToolTip(text["image_keep_videos_warning"])

    tool.watermark_checkbox.setChecked(True)
    tool.watermark_text_kind_button.setChecked(True)
    tool.watermark_image_kind_button.setChecked(False)
    tool.watermark_text_input.setText(
        "企业专属，注意保密 / INTERNAL REVIEW ONLY"
        if language == "zh"
        else "Internal review only / Confidential distribution"
    )
    tool.watermark_text_input.setCursorPosition(0)
    tool.preset_select.setCurrentIndex(
        next(
            index
            for index in range(tool.preset_select.count())
            if tool.preset_select.itemData(index)["key"] == preset_key
        )
    )
    tool.color_select.setCurrentIndex(
        tool.color_select.findData(dict(PRESET_COLORS)[color_label])
    )
    tool.font_fix_button.setChecked(True)
    tool._sync_font_fix_button()
    tool.preview_background_select.setCurrentIndex(
        next(
            index
            for index in range(tool.preview_background_select.count())
            if tool.preview_background_select.itemData(index)["key"] == preview_theme
        )
    )
    tool.pick_watermark_image_button.setText(f"PNG: {sample_png.name}")
    tool.pick_watermark_image_button.setToolTip(str(sample_png))
    tool.watermark_image_name_label.set_full_text(sample_png.name)
    tool.watermark_image_name_label.setToolTip(str(sample_png))
    tool.watermark_image_path = sample_png
    tool.on_settings_changed()


def configure_media(toolbox: ToolboxWindow, language: str) -> None:
    tool = toolbox.embedded_tools[1].window
    if tool is None:
        return
    text = tool.text
    video_key = "profile_none" if language == "zh" else "profile_balanced"
    image_key = "image_profile_none" if language == "zh" else "image_profile_balanced"
    tool.target_input.clear()
    tool.profile_select.setCurrentIndex(tool.profile_select.findText(text[video_key]))
    tool.image_profile_select.setCurrentIndex(
        tool.image_profile_select.findText(text[image_key])
    )
    tool.threshold_spinbox.setValue(0.98 if language == "zh" else 0.95)
    tool.on_profile_changed()


def configure_media_populated(
    toolbox: ToolboxWindow,
    language: str,
    mock_files: dict[str, Path],
) -> None:
    configure_media(toolbox, language)
    tool = toolbox.embedded_tools[1].window
    if tool is not None:
        tool.set_files([mock_files["pptx_alpha"], mock_files["mp4"], mock_files["png"]])


def configure_watermark_mixed_queue(
    toolbox: ToolboxWindow,
    language: str,
    sample_png: Path,
    mock_files: dict[str, Path],
) -> None:
    tool = toolbox.embedded_tools[0].window
    if tool is None:
        return
    configure_watermark(toolbox, language, sample_png)
    tool.input_paths = [
        mock_files["pptx_alpha"],
        mock_files["docx"],
        mock_files["pdf"],
    ]
    tool.checked_paths = set(tool.input_paths)
    tool.file_statuses = {path: "pending" for path in tool.input_paths}
    tool.file_outputs = {
        mock_files["pptx_alpha"]: mock_files["pptx_alpha"].with_name(
            "sample_alpha_editable_watermarked.pptx"
        ),
        mock_files["docx"]: mock_files["docx"].with_name("sample_beta_watermarked.pdf"),
        mock_files["pdf"]: mock_files["pdf"].with_name("sample_gamma_watermarked.pdf"),
    }
    tool.output_format_select.setCurrentIndex(
        tool.output_format_select.findData("pptx")
    )
    tool.refresh_file_list()
    tool.file_list.setCurrentRow(1)
    tool.update_idle_status_label()
    tool.output_path_label.set_full_text("输出文件: sample_beta_watermarked.pdf")
    tool.output_path_label.setToolTip(str(tool.file_outputs[mock_files["docx"]]))
    tool.font_check_label.setText(tool.text["font_check_non_pptx"])
    tool.font_check_label.setToolTip(tool.text["font_check_non_pptx"])
    tool.event_log.setPlainText(tool.text["mixed_queue_pptx_body"])
    tool.progress_bar.setValue(0)
    pages = [sample_png] * 5
    tool.current_preview = PreviewArtifacts(
        temp_root=None,
        source_key=("handoff", 0, 0),
        original_paths=pages,
        preview_paths=pages,
        total_pages=24,
    )
    tool.current_preview_page = 2
    tool.update_preview_display()


def configure_watermark_portrait(
    toolbox: ToolboxWindow,
    language: str,
    sample_png: Path,
    portrait_png: Path,
    mock_files: dict[str, Path],
) -> None:
    configure_watermark(toolbox, language, sample_png)
    tool = toolbox.embedded_tools[0].window
    if tool is None:
        return
    tool.input_paths = [mock_files["pdf"]]
    tool.checked_paths = set(tool.input_paths)
    tool.file_statuses = {mock_files["pdf"]: "pending"}
    tool.refresh_file_list()
    tool.file_list.setCurrentRow(0)
    pages = [portrait_png] * 6
    tool.current_preview = PreviewArtifacts(
        temp_root=None,
        source_key=("handoff-portrait", 0, 0),
        original_paths=pages,
        preview_paths=pages,
        total_pages=6,
    )
    tool.current_preview_page = 0
    tool.update_preview_display()


def configure_watermark_running(
    toolbox: ToolboxWindow,
    language: str,
    sample_png: Path,
    mock_files: dict[str, Path],
) -> None:
    tool = toolbox.embedded_tools[0].window
    if tool is None:
        return
    configure_watermark(toolbox, language, sample_png)
    tool.input_paths = [mock_files["pptx_alpha"], mock_files["docx"]]
    tool.checked_paths = set(tool.input_paths)
    tool.file_statuses = {
        mock_files["pptx_alpha"]: "running",
        mock_files["docx"]: "queued",
    }
    tool.file_outputs = {
        mock_files["pptx_alpha"]: mock_files["pptx_alpha"].with_name(
            "sample_alpha_editable_watermarked.pdf"
        ),
        mock_files["docx"]: mock_files["docx"].with_name("sample_beta_watermarked.pdf"),
    }
    tool.refresh_file_list()
    tool.set_running(True)
    tool.current_file_label.setText(
        f"{tool.text['current_processing']}: {mock_files['pptx_alpha'].name}"
    )
    tool.output_path_label.set_full_text(
        "输出文件: sample_alpha_editable_watermarked.pdf"
    )
    tool.output_path_label.setToolTip(str(tool.file_outputs[mock_files["pptx_alpha"]]))
    tool.event_log.setPlainText(
        "任务开始：2 个文件。\n"
        f"正在处理 1/2: {mock_files['pptx_alpha'].name}\n"
        "Converting PPTX to PDF\n"
        "Trying PDF export engine: LibreOffice (macOS)"
    )
    tool.progress_bar.setValue(38)


def configure_media_running(
    toolbox: ToolboxWindow,
    language: str,
    mock_files: dict[str, Path],
) -> None:
    tool = toolbox.embedded_tools[1].window
    if tool is None:
        return
    configure_media(toolbox, language)
    tool.input_paths = [mock_files["pptx_alpha"], mock_files["mp4"]]
    tool.file_statuses = {
        mock_files["pptx_alpha"]: "running",
        mock_files["mp4"]: "queued",
    }
    tool.refresh_file_list()
    tool.file_list.setCurrentRow(0)
    tool.set_running(True)
    tool.current_file_label.setText(
        f"{tool.text['current_processing']}: {mock_files['pptx_alpha'].name}"
    )
    tool.event_log.setPlainText(
        f"任务开始：2 个文件。\n正在处理 1/2: {mock_files['pptx_alpha'].name}\n"
        "Using ffmpeg (bundled)\nGPU: auto -> videotoolbox\n[ 37%] 正在处理视频 2/5"
    )
    tool.progress_bar.setValue(37)


def configure_media_done(
    toolbox: ToolboxWindow,
    language: str,
    mock_files: dict[str, Path],
) -> None:
    tool = toolbox.embedded_tools[1].window
    if tool is None:
        return
    configure_media(toolbox, language)
    tool.input_paths = [mock_files["pptx_alpha"]]
    tool.file_statuses = {mock_files["pptx_alpha"]: "done"}
    output = mock_files["pptx_alpha"].with_name(
        "sample_alpha_compressed_media_high.pptx"
    )
    output.write_bytes(b"compressed sample")
    tool.output_paths = {mock_files["pptx_alpha"]: output}
    tool.refresh_file_list()
    tool.current_file_label.setText("已完成 1 个文件，可继续评估或提档。")
    tool.event_log.setPlainText(
        f"[DONE] {mock_files['pptx_alpha'].with_name('sample_alpha_compressed_media_high.pptx')}"
    )
    tool.progress_bar.setValue(100)
    tool.update_audit_button_state()


def capture_toolbox(
    language: str,
    tab_index: int,
    output_name: str,
    configurator=None,
    size: tuple[int, int] = DEFAULT_SIZE,
    sample_png: Path | None = None,
) -> Path:
    set_language(language)
    app = app_instance()
    window = ToolboxWindow()
    window.resize(*size)
    window.tabs.setCurrentIndex(tab_index)
    if configurator is not None:
        configurator(window, language, sample_png) if sample_png else configurator(
            window, language
        )
    process_events(app, 12)
    QTest.qWait(240)
    process_events(app, 2)
    output = SCREENSHOT_DIR / output_name
    window.grab().save(str(output))
    window.shutdown_embedded_tools()
    window.close()
    process_events(app, 4)
    return output


def configure_empty_library(toolbox: ToolboxWindow, _language: str, index: int) -> None:
    tool = toolbox.embedded_tools[index].window
    tool.project = None
    if index == 2:
        tool.project_label.setText("尚未打开视频库 · 可按主题新建多个独立库")
        tool.log_shelf.setText("状态与日志 · 等待操作")
    else:
        tool.project_path.setText("尚未打开图片库")
        tool.status_label.setText("状态与日志 · 等待开始")
    tool.refresh_views()


def build_field_catalog() -> dict[str, object]:
    load_ui_modules()
    watermark_zh = SHELL_STRINGS["zh"]
    watermark_strings = WATERMARK_STRINGS["zh"]
    media_strings = MEDIA_STRINGS["zh"]

    return {
        "meta": {
            "app": "Doc Media Toolkit",
            "ui_type": "Qt desktop / PySide6",
            "default_size": {"width": DEFAULT_SIZE[0], "height": DEFAULT_SIZE[1]},
            "minimum_review_size": {
                "width": 1280,
                "height": 800,
            },
            "tabs": [
                "文档及媒体水印导出",
                "文档及媒体动态压缩",
                "视频库",
                "图片库",
            ],
        },
        "watermark_page": {
            "subtitle": watermark_zh["watermark_subtitle"],
            "selects": {
                "格式": ["PDF", "PPTX"],
                "形式": [
                    watermark_strings["output_mode_editable"],
                    watermark_strings["output_mode_image"],
                ],
                "质量": [
                    watermark_strings["image_quality_original"],
                    watermark_strings["image_quality_high"],
                    watermark_strings["image_quality_balanced"],
                    watermark_strings["image_quality_low"],
                ],
                "模板": [preset["zh"] for preset in WATERMARK_PRESETS],
                "颜色": [name for name, _ in PRESET_COLORS],
                "背景": [theme["zh"] for theme in PREVIEW_BACKGROUND_THEMES],
            },
            "buttons": {
                "保留视频": [
                    watermark_strings["image_keep_videos_button"],
                    watermark_strings["image_keep_videos_on"],
                    watermark_strings["image_keep_videos_off"],
                ],
                "水印": [
                    watermark_strings["watermark_toggle_on"],
                    watermark_strings["watermark_toggle_off"],
                ],
                "类型": [
                    watermark_strings["watermark_type_text"],
                    watermark_strings["watermark_type_image"],
                ],
                "预览区按钮": [
                    watermark_strings["preview_prev"],
                    watermark_strings["preview_next"],
                    watermark_strings["preview_refresh"],
                    watermark_strings["preview_original"],
                    watermark_strings["preview_output"],
                ],
            },
            "longest_copy": {
                "水印文字": "企业专属，注意保密 / INTERNAL REVIEW ONLY",
                "运行依赖提示": watermark_strings["dependency_missing_libreoffice"],
                "字体检查提示": watermark_strings["font_check_missing"].format(
                    count=9,
                    fonts="HarmonyOS Sans SC, FZLanTingHeiS-DB1-GB, DIN, Arial Narrow, PingFang SC",
                    family="Noto Sans SC",
                ),
                "视频回贴提示": watermark_strings["image_keep_videos_warning"],
            },
            "tooltips_hidden": {
                "视频回贴按钮": watermark_strings["image_keep_videos_warning"],
                "图片化依赖": watermark_strings["dependency_missing_ffmpeg"],
                "非 PPTX 字体补齐": watermark_strings["font_check_non_pptx"],
                "混合队列提醒": watermark_strings["mixed_queue_pptx_body"],
            },
            "state_matrix": {
                "空列表": "按钮禁用，队列提示为未选择文件。",
                "已选文件": "显示队列计数、输出提示和可执行按钮。",
                "混合队列": "PPTX + DOCX/PDF 同时存在时，DOCX/PDF 强制导出 PDF。",
                "运行中": "开始导出按钮切换为停止，文件状态显示 queued/running。",
                "完成": "文件项显示 done，日志保留输出路径。",
                "失败": "文件项显示 failed，日志显示失败原因。",
            },
        },
        "media_page": {
            "subtitle": SHELL_STRINGS["zh"]["video_subtitle"],
            "selects": {
                "视频": [media_strings[f"profile_{key}"] for key in PRESET_OPTIONS],
                "图片": [
                    media_strings[f"image_profile_{key}"] for key in PRESET_OPTIONS
                ],
            },
            "inputs": {
                "文件大小 placeholder": media_strings["target_placeholder"],
                "阈值范围": "0.00 - 1.00",
            },
            "buttons": [
                media_strings["audit_button"],
                media_strings["optimize_button"],
                media_strings["run"],
                media_strings["pick_file"],
                media_strings["remove_file"],
            ],
            "tooltips": {
                "视频阈值": media_strings["video_threshold_tooltip"],
                "图片阈值": media_strings["image_threshold_tooltip"],
            },
            "state_matrix": {
                "空列表": "开始压缩、画质评估、提档优化均禁用。",
                "已选文件": "显示预计输出，开始压缩可点击。",
                "运行中": "开始压缩切为停止，文件状态显示 queued/running。",
                "完成": "画质评估按钮启用，可继续提档优化。",
                "评估后": "只针对 done 文件进行 SSIM 质量检查。",
                "提档后": "只重写低分素材，不生成第三个文件。",
            },
        },
        "video_library_page": {
            "primary_surface": "视频族/版本列表与按需详情抽屉",
            "states": ["全部", "待核对", "无关联", "多版本", "文件异常"],
            "actions": ["播放", "核对关联", "归并视频", "整理视频库", "待清理"],
        },
        "image_library_page": {
            "primary_surface": "图片列表、预览与来源详情",
            "states": ["全部", "重复来源", "相似", "过小", "无来源"],
            "actions": ["编辑信息", "查找相似图", "AI 整理建议", "待清理"],
        },
    }


def write_markdown(screenshots: dict[str, str], catalog: dict[str, object]) -> None:
    content = f"""# Doc Media Toolkit UI Handoff

## 目的
这不是网页前端，而是 `PySide6 / Qt` 桌面界面。这个目录用于把当前交互页交给其他在线 AI 做视觉优化、间距优化、控件尺寸审查和文案排版校对。

## 目录内容
- `screenshots/`：当前界面截图
- `docs/ui-handoff.html`：可直接浏览的静态交接页
- `docs/ui-handoff.md`：文字规格说明
- `docs/UI_DESIGN.md`：唯一有效的完整设计规范
- `docs/design-qa.md`：设计与实现核对结论
- `data/ui-field-catalog.json`：给 AI 或脚本读取的字段清单
- `data/sample_watermark_image.png`：图片水印占位样例

## 当前技术形态
- 容器：`Doc Media Toolkit`
- UI 框架：`PySide6 / Qt`
- 默认窗口尺寸：`{DEFAULT_SIZE[0]} x {DEFAULT_SIZE[1]}`
- 目标平台：`macOS / Windows`
- 主功能页：
  - `文档及媒体水印导出`
  - `文档及媒体动态压缩`
  - `PPTX 视频资产库`
  - `文档图片资产库`

## 交给外部 AI 时要说明
- 这是桌面 Qt GUI，不是 Web 页面。
- 优化目标以 `布局工整、最长字段不遮挡、控件高度统一、最小窗口也能看清` 为主。
- 不要改变媒体处理、匹配、关联、清理和回填业务逻辑。
- 允许优化圆角、控件宽度、留白、标题对齐、分组关系、视觉层级。
- 需要兼顾中英文。
- 桌面审阅基线按 `1280 x 800` 处理，默认工作区为 `1440 x 900`；统一壳本身的最小尺寸更小。

## 必须校验的最长字段
### 文档及媒体水印导出
- 形式：`图片化`
- 图片质量：`平衡`
- 视频回贴状态：`不处理`
- 水印文字：`企业专属，注意保密 / INTERNAL REVIEW ONLY`
- 运行依赖提示：
  `{catalog["watermark_page"]["longest_copy"]["运行依赖提示"]}`
- 字体检查提示：
  `{catalog["watermark_page"]["longest_copy"]["字体检查提示"]}`
- 视频回贴提示：
  `{catalog["watermark_page"]["longest_copy"]["视频回贴提示"]}`

### 文档及媒体动态压缩
- 文件大小 placeholder：`留空按预设`
- 视频预设：`不压缩`
- 图片预设：`不压缩`
- 按钮：`画质评估`、`提档优化`、`开始压缩`
- 阈值 tooltip：需要完整容纳多行说明

### PPTX 视频资产库
- 主工作面：视频族/版本列表
- 状态筛选：`全部`、`待核对`、`无关联`、`多版本`、`文件异常`
- 详情抽屉：单帧封面、规格、引用、路径、哈希和上下文操作

### 文档图片资产库
- 主工作面：图片列表、预览和来源详情
- 状态筛选：`全部`、`重复来源`、`相似`、`过小`、`无来源`
- 低频操作：AI 建议、打开位置、移除、库体检和待清理

## 状态矩阵
### 文档及媒体水印导出
- 空列表：按钮禁用，队列提示为未选择文件
- 已选文件：显示队列数量，可直接开始导出
- 混合队列：`PPTX + DOCX + PDF` 同时存在时，DOCX/PDF 只导 PDF
- 运行中：文件项出现 `queued / running`
- 完成：文件项出现 `done`
- 失败：文件项出现 `failed`

### 文档及媒体动态压缩
- 空列表：开始压缩、画质评估、提档优化均禁用
- 已选文件：显示预计输出，可开始压缩
- 运行中：文件项出现 `queued / running`
- 完成：画质评估启用
- 评估后：只针对完成文件做质量评估
- 提档后：只替换低分素材，不生成第三个文件

### PPTX 视频资产库
- 默认：列表占主区域，详情抽屉关闭
- 选择：右侧覆盖式详情抽屉打开，不挤压关键列
- 待核对：只展示需要人工判断的去重后视频族
- 文件异常：允许隔离，禁止静默删除仍被引用的版本

## 隐藏说明 / Tooltip
### 文档及媒体水印导出
- 视频回贴：
  `{catalog["watermark_page"]["tooltips_hidden"]["视频回贴按钮"]}`
- 图片化依赖：
  `{catalog["watermark_page"]["tooltips_hidden"]["图片化依赖"]}`
- 非 PPTX 字体补齐：
  `{catalog["watermark_page"]["tooltips_hidden"]["非 PPTX 字体补齐"]}`
- 混合队列提醒：
  `{catalog["watermark_page"]["tooltips_hidden"]["混合队列提醒"]}`

### 文档及媒体动态压缩
- 视频阈值说明：
  `{catalog["media_page"]["tooltips"]["视频阈值"]}`
- 图片阈值说明：
  `{catalog["media_page"]["tooltips"]["图片阈值"]}`

## 截图索引
- 文档及媒体水印导出默认态：`{screenshots["watermark_default_zh"]}`
- 文档及媒体水印导出最长字段态：`{screenshots["watermark_max_zh"]}`
- 文档及媒体水印导出混合队列态：`{screenshots["watermark_mixed_zh"]}`
- 文档及媒体水印导出 A4 纵向态：`{screenshots["watermark_portrait_zh"]}`
- 文档及媒体水印导出运行态：`{screenshots["watermark_running_zh"]}`
- 文档及媒体动态压缩默认态：`{screenshots["media_default_zh"]}`
- 文档及媒体动态压缩最长字段态：`{screenshots["media_max_zh"]}`
- 文档及媒体动态压缩运行态：`{screenshots["media_running_zh"]}`
- 文档及媒体动态压缩完成态：`{screenshots["media_done_zh"]}`
- English watermark state：`{screenshots["watermark_max_en"]}`
- English media state：`{screenshots["media_max_en"]}`
- PPTX 视频资产库：`{screenshots["video_library_default_zh"]}`
- 文档图片资产库：`{screenshots["image_library_default_zh"]}`
- PPTX 视频资产库详情抽屉：`video-library-detail-zh.png`
- 视频匹配核对：`dialog-video-match-current.png`
- PPTX 高清回填核对：`dialog-pptx-restore-current.png`
- 视频库整理：`dialog-cleanup-current.png`
- 待清理管理：`dialog-pending-cleanup-current.png`
- 视频库体检：`dialog-library-health-current.png`
"""
    (DOC_DIR / "ui-handoff.md").write_text(content, encoding="utf-8")


def write_html(screenshots: dict[str, str], catalog: dict[str, object]) -> None:
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Doc Media Toolkit UI Handoff</title>
  <style>
    :root {{
      --bg: #0b1017;
      --panel: #121a24;
      --line: #273244;
      --text: #e6edf6;
      --muted: #94a3b8;
      --accent: #f97316;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", sans-serif;
      background: radial-gradient(circle at top, #162233 0%, var(--bg) 52%);
      color: var(--text);
    }}
    .wrap {{ max-width: 1400px; margin: 0 auto; padding: 28px; }}
    .hero, .panel {{
      background: rgba(18, 26, 36, 0.96);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 22px;
      margin-bottom: 18px;
    }}
    h1, h2, h3 {{ margin: 0 0 10px; }}
    h1 {{ font-size: 32px; }}
    h2 {{ font-size: 20px; }}
    p, li {{ color: var(--muted); line-height: 1.6; }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    .shot {{
      background: #0f1720;
      border: 1px solid var(--line);
      border-radius: 16px;
      overflow: hidden;
    }}
    .shot img {{ display: block; width: 100%; height: auto; }}
    .shot .cap {{ padding: 12px 14px; font-size: 14px; color: var(--text); }}
    code {{
      background: #0f1720;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 2px 8px;
      color: #f8fafc;
    }}
    .accent {{ color: var(--accent); }}
    .two-col {{
      display: grid;
      grid-template-columns: 1.1fr 0.9fr;
      gap: 18px;
    }}
    @media (max-width: 980px) {{
      .grid, .two-col {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <div class="accent">DOC MEDIA TOOLKIT / UI HANDOFF</div>
      <h1>Doc Media Toolkit 交接页</h1>
      <p>当前产品是 Qt 桌面 GUI，不是网页。这个页面用于把真实界面、最长字段、文案约束和截图一起交给其他在线 AI 做视觉优化。</p>
    </section>

    <section class="two-col">
      <div class="panel">
        <h2>约束</h2>
        <p>保持四个工作区的业务逻辑不变，优先修正对齐、留白、分组、控件高度、最长文本可见性、中英文适配。</p>
        <p>默认窗口尺寸：<code>{DEFAULT_SIZE[0]} x {DEFAULT_SIZE[1]}</code></p>
        <p>桌面审阅尺寸基线：<code>{catalog["meta"]["minimum_review_size"]["width"]} x {catalog["meta"]["minimum_review_size"]["height"]}</code></p>
        <p>功能页：<code>文档及媒体水印导出</code>、<code>文档及媒体动态压缩</code>、<code>PPTX 视频资产库</code>、<code>文档图片资产库</code></p>
      </div>
      <div class="panel">
        <h2>最长字段重点</h2>
        <p><b>文档及媒体水印导出</b>：图片化、平衡、不处理、视频回贴提示、字体检查长提示、依赖说明长提示。</p>
        <p><b>文档及媒体动态压缩</b>：留空按预设、不压缩、画质评估、提档优化、阈值多行 tooltip。</p>
        <p><b>PPTX 视频资产库</b>：状态筛选、视频族/版本层级、覆盖式详情抽屉、匹配与回填核对。</p>
        <p><b>文档图片资产库</b>：状态筛选、图片预览、来源详情、相似核对和可恢复清理。</p>
      </div>
    </section>

    <section class="panel">
      <h2>截图</h2>
      <div class="grid">
        <div class="shot">
          <img src="../screenshots/{screenshots["watermark_default_zh"]}" alt="watermark default zh">
          <div class="cap">文档及媒体水印导出 / 默认态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["watermark_max_zh"]}" alt="watermark max zh">
          <div class="cap">文档及媒体水印导出 / 最长字段态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["watermark_mixed_zh"]}" alt="watermark mixed zh">
          <div class="cap">文档及媒体水印导出 / 混合队列态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["watermark_running_zh"]}" alt="watermark running zh">
          <div class="cap">文档及媒体水印导出 / 运行态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["media_default_zh"]}" alt="media default zh">
          <div class="cap">文档及媒体动态压缩 / 默认态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["media_max_zh"]}" alt="media max zh">
          <div class="cap">文档及媒体动态压缩 / 最长字段态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["media_running_zh"]}" alt="media running zh">
          <div class="cap">文档及媒体动态压缩 / 运行态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["media_done_zh"]}" alt="media done zh">
          <div class="cap">文档及媒体动态压缩 / 完成态 / 中文</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["watermark_max_en"]}" alt="watermark max en">
          <div class="cap">Document & Media Watermark Export / max text / English</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["media_max_en"]}" alt="media max en">
          <div class="cap">Document & Media Dynamic Compression / max text / English</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["video_library_default_zh"]}" alt="video library">
          <div class="cap">PPTX 视频资产库 / 主列表</div>
        </div>
        <div class="shot">
          <img src="../screenshots/{screenshots["image_library_default_zh"]}" alt="image library">
          <div class="cap">文档图片资产库 / 默认态</div>
        </div>
        <div class="shot">
          <img src="../screenshots/video-library-detail-zh.png" alt="video library detail">
          <div class="cap">PPTX 视频资产库 / 详情抽屉</div>
        </div>
        <div class="shot">
          <img src="../screenshots/dialog-video-match-current.png" alt="video match review">
          <div class="cap">视频匹配 / 人工核对</div>
        </div>
        <div class="shot">
          <img src="../screenshots/dialog-pptx-restore-current.png" alt="pptx restore review">
          <div class="cap">PPTX 高清回填 / 逐项核对</div>
        </div>
      </div>
    </section>

    <section class="two-col">
      <div class="panel">
        <h2>隐藏说明</h2>
        <p><b>文档及媒体水印导出 / 视频回贴：</b><br>{catalog["watermark_page"]["tooltips_hidden"]["视频回贴按钮"]}</p>
        <p><b>文档及媒体水印导出 / 混合队列：</b><br>{catalog["watermark_page"]["tooltips_hidden"]["混合队列提醒"]}</p>
        <p><b>文档及媒体动态压缩 / 视频阈值：</b><br>{catalog["media_page"]["tooltips"]["视频阈值"].replace(chr(10), "<br>")}</p>
        <p><b>文档及媒体动态压缩 / 图片阈值：</b><br>{catalog["media_page"]["tooltips"]["图片阈值"].replace(chr(10), "<br>")}</p>
      </div>
      <div class="panel">
        <h2>状态矩阵</h2>
        <p><b>文档及媒体水印导出：</b><br>{"<br>".join(f"{k}：{v}" for k, v in catalog["watermark_page"]["state_matrix"].items())}</p>
        <p><b>文档及媒体动态压缩：</b><br>{"<br>".join(f"{k}：{v}" for k, v in catalog["media_page"]["state_matrix"].items())}</p>
      </div>
    </section>
  </div>
</body>
</html>
"""
    (DOC_DIR / "ui-handoff.html").write_text(html, encoding="utf-8")


def main() -> None:
    ensure_dirs()
    app_instance()
    load_ui_modules()
    sample_png = make_sample_png()
    portrait_png = make_sample_portrait_png()
    mock_files = make_mock_files()

    screenshots = {
        "watermark_default_zh": capture_toolbox(
            "zh", 0, "watermark-default-zh.png", size=DEFAULT_SIZE
        ).name,
        "watermark_max_zh": capture_toolbox(
            "zh",
            0,
            "watermark-max-zh.png",
            configurator=configure_watermark,
            size=DEFAULT_SIZE,
            sample_png=sample_png,
        ).name,
        "watermark_mixed_zh": capture_toolbox(
            "zh",
            0,
            "watermark-mixed-zh.png",
            configurator=lambda window, lang, png: configure_watermark_mixed_queue(
                window, lang, png, mock_files
            ),
            size=DEFAULT_SIZE,
            sample_png=sample_png,
        ).name,
        "watermark_portrait_zh": capture_toolbox(
            "zh",
            0,
            "watermark-portrait-zh.png",
            configurator=lambda window, lang, png: configure_watermark_portrait(
                window, lang, png, portrait_png, mock_files
            ),
            size=DEFAULT_SIZE,
            sample_png=sample_png,
        ).name,
        "watermark_running_zh": capture_toolbox(
            "zh",
            0,
            "watermark-running-zh.png",
            configurator=lambda window, lang, png: configure_watermark_running(
                window, lang, png, mock_files
            ),
            size=DEFAULT_SIZE,
            sample_png=sample_png,
        ).name,
        "media_default_zh": capture_toolbox(
            "zh", 1, "media-default-zh.png", size=DEFAULT_SIZE
        ).name,
        "media_max_zh": capture_toolbox(
            "zh",
            1,
            "media-max-zh.png",
            configurator=lambda window, lang: configure_media_populated(
                window, lang, mock_files
            ),
            size=DEFAULT_SIZE,
        ).name,
        "media_running_zh": capture_toolbox(
            "zh",
            1,
            "media-running-zh.png",
            configurator=lambda window, lang: configure_media_running(
                window, lang, mock_files
            ),
            size=DEFAULT_SIZE,
        ).name,
        "media_done_zh": capture_toolbox(
            "zh",
            1,
            "media-done-zh.png",
            configurator=lambda window, lang: configure_media_done(
                window, lang, mock_files
            ),
            size=DEFAULT_SIZE,
        ).name,
        "watermark_max_en": capture_toolbox(
            "en",
            0,
            "watermark-max-en.png",
            configurator=configure_watermark,
            size=DEFAULT_SIZE,
            sample_png=sample_png,
        ).name,
        "media_max_en": capture_toolbox(
            "en",
            1,
            "media-max-en.png",
            configurator=lambda window, lang: configure_media_populated(
                window, lang, mock_files
            ),
            size=DEFAULT_SIZE,
        ).name,
        "video_library_default_zh": capture_toolbox(
            "zh",
            2,
            "video-library-default-zh.png",
            configurator=lambda window, lang: configure_empty_library(window, lang, 2),
            size=DEFAULT_SIZE,
        ).name,
        "image_library_default_zh": capture_toolbox(
            "zh",
            3,
            "image-library-default-zh.png",
            configurator=lambda window, lang: configure_empty_library(window, lang, 3),
            size=DEFAULT_SIZE,
        ).name,
    }

    catalog = build_field_catalog()
    (DATA_DIR / "ui-field-catalog.json").write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_markdown(screenshots, catalog)
    write_html(screenshots, catalog)

    print(DESKTOP_DIR)


if __name__ == "__main__":
    main()
