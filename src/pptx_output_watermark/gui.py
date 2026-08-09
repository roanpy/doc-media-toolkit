from __future__ import annotations

import logging
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtCore import (
    QEvent,
    QFileInfo,
    QObject,
    QPoint,
    QRect,
    QSignalBlocker,
    QSize,
    QThread,
    QTimer,
    Qt,
    QUrl,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QDesktopServices,
    QDragEnterEvent,
    QDragMoveEvent,
    QDropEvent,
    QFont,
    QFontDatabase,
    QIcon,
    QKeyEvent,
    QPainter,
    QPen,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractSpinBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFileIconProvider,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QButtonGroup,
    QSizePolicy,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .dependencies import dependency_statuses, missing_dependency_message
from .export_pipeline import (
    default_output_path,
    effective_output_format as effective_export_output_format,
    export_document,
)
from .font_assets import BUNDLED_WATERMARK_FONT_FAMILY, bundled_watermark_font_path
from .models import DEFAULT_WATERMARK_TEXT, ExportOptions, WatermarkOptions
from .pdf_io import open_pdf_reader
from .pdf_rendering import batch_render_pdf_slides
from .presentation_rendering import convert_document_to_pdf
from .process_utils import terminate_active_processes
from .pptx_fonts import FontScanResult, default_source_font_family, scan_missing_fonts
from .pptx_video_support import extract_video_poster_frame
from .runtime_temp import cleanup_stale_runtime_entries, create_runtime_temp_dir
from .watermarking import apply_watermark_to_image
from pptx_tools.app_logging import configure_app_logging
from pptx_tools.language import detect_language as detect_system_language
from pptx_tools.ui_theme import (
    SHARED_DIALOG_QSS,
    SHARED_MAIN_QSS,
    configure_ui_font,
    format_user_file_size,
    install_control_help,
)
from pptx_video_compactor import IMAGE_EXTENSIONS, VIDEO_EXTENSIONS

LOGGER = logging.getLogger("pptx_output_watermark.gui")
PREVIEW_PAGE_LIMIT = 5
AUTO_PREVIEW_MAX_BYTES = 80 * 1024 * 1024
LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"
PRESET_COLORS = [
    ("Slate", "#667788"),
    ("Gray", "#8A97A3"),
    ("Cloud", "#E2E8F0"),
    ("Blue", "#5B728B"),
    ("Amber", "#F59E0B"),
    ("Rose", "#9D6C73"),
]
WATERMARK_PRESETS = [
    {
        "key": "light_balanced",
        "zh": "白底",
        "en": "Light",
        "zh_desc": "适合白色或浅色页面，灰蓝低干扰，不抢正文内容。",
        "en_desc": "For white or light slides. Low-impact slate watermark.",
        "color": "#667788",
        "opacity": 0.18,
        "font_size": 34,
        "spacing": 360,
        "angle": 315,
    },
    {
        "key": "dark_subtle",
        "zh": "深色",
        "en": "Dark",
        "zh_desc": "适合深色页面，浅灰可见但不刺眼。",
        "en_desc": "For dark slides. Light gray remains visible without glare.",
        "color": "#E2E8F0",
        "opacity": 0.18,
        "font_size": 32,
        "spacing": 340,
        "angle": 315,
    },
    {
        "key": "black_warm",
        "zh": "黑底",
        "en": "Black",
        "zh_desc": "适合黑底或高对比页面，暖橙更容易识别。",
        "en_desc": "For black or high-contrast slides. Warm amber is easier to notice.",
        "color": "#F59E0B",
        "opacity": 0.16,
        "font_size": 32,
        "spacing": 360,
        "angle": 315,
    },
    {
        "key": "mixed_low",
        "zh": "混合",
        "en": "Mixed",
        "zh_desc": "适合明暗混合内容，透明度更低、间距更大。",
        "en_desc": "For mixed backgrounds. Lower opacity and wider spacing.",
        "color": "#8A97A3",
        "opacity": 0.12,
        "font_size": 28,
        "spacing": 420,
        "angle": 315,
    },
    {
        "key": "dense_review",
        "zh": "密集",
        "en": "Dense",
        "zh_desc": "适合内部审阅稿，密度更高，防止页面局部无水印。",
        "en_desc": "For internal review. Denser placement reduces blank areas.",
        "color": "#9D6C73",
        "opacity": 0.20,
        "font_size": 30,
        "spacing": 260,
        "angle": 315,
    },
]
DOCUMENT_INPUT_EXTENSIONS = {".pptx", ".docx", ".pdf"}
MEDIA_IMAGE_INPUT_EXTENSIONS = set(IMAGE_EXTENSIONS)
MEDIA_VIDEO_INPUT_EXTENSIONS = set(VIDEO_EXTENSIONS)
WATERMARK_IMAGE_INPUT_EXTENSIONS = tuple(sorted(MEDIA_IMAGE_INPUT_EXTENSIONS))
SUPPORTED_WATERMARK_INPUT_EXTENSIONS = tuple(
    dict.fromkeys(
        [
            *sorted(DOCUMENT_INPUT_EXTENSIONS),
            *sorted(MEDIA_IMAGE_INPUT_EXTENSIONS),
            *sorted(MEDIA_VIDEO_INPUT_EXTENSIONS),
        ]
    )
)


def is_document_input(path: Path) -> bool:
    return path.suffix.lower() in DOCUMENT_INPUT_EXTENSIONS


def is_media_image_input(path: Path) -> bool:
    return path.suffix.lower() in MEDIA_IMAGE_INPUT_EXTENSIONS


def is_media_video_input(path: Path) -> bool:
    return path.suffix.lower() in MEDIA_VIDEO_INPUT_EXTENSIONS


def is_standalone_media_input(path: Path) -> bool:
    return is_media_image_input(path) or is_media_video_input(path)


PREVIEW_BACKGROUND_THEMES = [
    {
        "key": "black",
        "zh": "黑底",
        "en": "Black",
        "bg": "#070B12",
        "line": "#E2E8F0",
        "muted": "#64748B",
        "panel": "#1E293B",
        "panel_soft": "#334155",
    },
    {
        "key": "white",
        "zh": "白底",
        "en": "White",
        "bg": "#F8FAFC",
        "line": "#1E293B",
        "muted": "#64748B",
        "panel": "#CBD5E1",
        "panel_soft": "#E2E8F0",
    },
    {
        "key": "gray",
        "zh": "灰底",
        "en": "Gray",
        "bg": "#334155",
        "line": "#F8FAFC",
        "muted": "#CBD5E1",
        "panel": "#475569",
        "panel_soft": "#64748B",
    },
    {
        "key": "blue",
        "zh": "蓝底",
        "en": "Blue",
        "bg": "#0F2A44",
        "line": "#DBEAFE",
        "muted": "#93C5FD",
        "panel": "#1D4ED8",
        "panel_soft": "#2563EB",
    },
]

STRINGS = {
    "zh": {
        "window_title": "文档及媒体水印导出",
        "eyebrow": "DOC MEDIA TOOLKIT",
        "title": "文档及媒体水印导出",
        "subtitle": "批量处理 DOCX / PDF / PPTX / 图片 / 视频，支持水印导出",
        "queue_title": "文件列表",
        "empty_queue": "拖入文件到这里，或点击添加",
        "empty_queue_hint": "DOCX / PDF / PPTX / 图片 / 视频 · 支持多选",
        "pick_file": "添加文件",
        "remove_file": "移除选中",
        "selected_summary": "已选 {selected}/{total} 个文件",
        "settings_title": "导出设置",
        "output_format_label": "格式",
        "output_format_pdf": "PDF",
        "output_format_pptx": "PPTX",
        "output_format_source": "原格式",
        "output_mode_label": "形式",
        "output_mode_editable": "可编辑",
        "output_mode_image": "图片化",
        "output_mode_not_applicable": "不适用",
        "output_format_document_locked_hint": "当前选中文档固定输出为 PDF。",
        "output_format_media_locked_hint": "当前选中图片/视频固定按原格式或 MP4 输出。",
        "media_mode_locked_hint": "独立图片/视频不使用“形式”设置；按原格式或 MP4 输出。",
        "media_quality_locked_hint": "独立图片会影响 JPG / WEBP 输出质量；PNG 保持原格式。独立视频会按 原 / 高 / 平衡 / 低 四档重编码并加水印。",
        "image_quality_mode_hint": "质量仅用于图片化导出；可编辑导出不使用这里的质量档。",
        "watermark_enable": "添加水印",
        "watermark_switch": "水印",
        "watermark_toggle_on": "水印 开",
        "watermark_toggle_off": "水印 关",
        "watermark_type": "类型",
        "watermark_type_text": "文字",
        "watermark_type_image": "图片",
        "watermark_text": "水印文字",
        "watermark_image": "水印图片",
        "watermark_pick_image": "选择或拖入图片",
        "watermark_no_image": "未选择图片",
        "watermark_image_width": "图片宽",
        "watermark_image_missing_title": "缺少水印图片",
        "watermark_image_missing_body": "图片水印模式需要先选择或拖入图片。",
        "watermark_preset": "模板",
        "watermark_preset_hint": "模板说明",
        "watermark_text_placeholder": "例如：企业专属，注意保密",
        "watermark_color": "颜色",
        "watermark_font_size": "字号",
        "watermark_spacing": "间距",
        "watermark_opacity": "透明度",
        "watermark_angle": "角度",
        "image_dpi": "图片清晰度",
        "image_quality_original": "原",
        "image_quality_high": "高",
        "image_quality_balanced": "平衡",
        "image_quality_low": "低",
        "image_keep_videos": "内嵌视频",
        "image_keep_videos_button": "视频设置",
        "image_keep_videos_on": "加水印并回填",
        "image_keep_videos_off": "不保留",
        "image_keep_videos_hint": "仅用于 图片型 PPTX：关闭时不处理视频；开启后会给视频加水印并尝试按原位置贴回，失败时在输出旁生成 _videos 目录。",
        "image_keep_videos_warning": "注意取消视频与其他元素的组合，否则回填将错位",
        "preview_title": "效果预览",
        "preview_note": "仅渲染前几页看效果，不展示视频交互。",
        "preview_waiting": "选择左侧文件后自动生成预览。",
        "preview_thumbnails_empty": "选择文件后显示页面缩略图",
        "preview_manual": "文件较大，点击刷新预览后生成。",
        "preview_loading": "正在生成预览…",
        "preview_failed": "预览生成失败",
        "preview_original": "原稿",
        "preview_output": "预览",
        "preview_refresh": "刷新",
        "preview_prev": "上页",
        "preview_next": "下页",
        "preview_empty_page": "暂无可显示页面",
        "preview_page": "第 {index}/{rendered} 页预览 · 共 {total} 页",
        "preview_background": "背景",
        "details_title": "状态与日志",
        "advanced_settings": "高级设置",
        "advanced_expand": "展开高级设置",
        "advanced_collapse": "收起高级设置",
        "log_expand": "展开日志",
        "log_collapse": "收起日志",
        "log_hint": "导出日志会显示在这里。",
        "dependency_title": "运行依赖",
        "dependency_ok": "当前模式依赖已满足。",
        "dependency_ok_short": "依赖已满足。",
        "dependency_missing_libreoffice": "缺少 PDF 导出引擎。macOS 上 PPTX 可获取 Keynote、DOCX 可获取 Pages；如版式效果不佳可安装 LibreOffice。Windows 建议安装 Microsoft Office/WPS 或 LibreOffice。",
        "dependency_keynote_fallback": "依赖已满足：使用 Keynote 兜底；如版式效果不佳可安装 LibreOffice。",
        "dependency_pages_fallback": "依赖已满足：使用 Pages 兜底；如版式效果不佳可安装 LibreOffice。",
        "dependency_missing_ffmpeg": "缺少 FFmpeg。图片型 PPTX 回贴视频需要 FFmpeg。",
        "dependency_install_ffmpeg": "安装 FFmpeg",
        "dependency_install_libreoffice": "安装 LibreOffice",
        "dependency_install_keynote": "获取 Keynote",
        "dependency_install_pages": "获取 Pages",
        "dependency_open_settings": "打开权限设置",
        "dependency_permission_keynote": "Keynote 已安装，但当前没有自动化权限。请先在系统设置里允许 Doc Media Toolkit 或 Terminal 控制 Keynote。",
        "dependency_permission_pages": "Pages 已安装，但当前没有自动化权限。请先在系统设置里允许 Doc Media Toolkit 或 Terminal 控制 Pages。",
        "fallback_warning": "警告：PDF/图片导出引擎不可用，已降级生成带水印的可编辑 PPTX。",
        "font_check_title": "字体检查",
        "font_check_waiting": "选择文件后检查原稿字体。",
        "font_check_non_pptx": "源字体补齐仅支持 PPTX；DOCX / PDF 按导出引擎原样处理。",
        "font_check_ok": "未检测到缺失字体。水印字体已内置。",
        "font_check_missing": "检测到 {count} 个可能缺失字体：{fonts}。可开启补齐，用 {family} 临时替换。",
        "font_fix_on": "补齐 开",
        "font_fix_off": "补齐 关",
        "status_ready": "等待开始。添加文件后可直接导出。",
        "status_running": "正在导出，请稍候。",
        "status_stopping": "停止已请求。当前文件完成后停止。",
        "current_processing": "正在导出",
        "status_done": "导出完成。",
        "status_failed": "导出失败。",
        "output_path_label": "输出文件",
        "output_waiting": "选择文件后显示默认输出路径。",
        "log_waiting": "等待开始。",
        "run": "开始导出",
        "stop": "停止",
        "stopping": "等待当前文件完成",
        "help_button": "使用说明",
        "help_title": "使用说明",
        "help_body": (
            "1. 左侧可批量加入 DOCX / PDF / PPTX，也可直接加入 PNG / JPG / WEBP / MP4 / MOV 等独立媒体文件。\n"
            "2. DOCX / PDF 只能导出为 PDF；混合队列里即使全局选择 PPTX，这两类文件也会按你选定的形式导出为 PDF。\n"
            "3. 独立图片按原格式输出；独立视频会在加水印后输出为原容器格式或 MP4。\n"
            "4. 图片清晰度只影响图片化 PDF / 图片化 PPTX；可编辑 PDF / 可编辑 PPTX 会尽量保留原始图片与文档结构。\n"
            "5. 右下预览只展示前几页或媒体首帧，用来确认文字或图片水印效果，不代表视频可播放。\n"
            "6. 图片型 PPTX 默认不保留可播放视频；仅 图片型 PPTX 的“回贴视频”可尝试恢复视频。\n"
            "7. 默认不保留过程产物，输出文件直接放在源文件旁边。"
        ),
        "choose_title": "选择文件",
        "choose_filter": "Supported Files (*.pptx *.docx *.pdf *.png *.jpg *.jpeg *.jpe *.webp *.mp4 *.m4v *.mov *.wmv *.asf *.avi *.mpg *.mpeg *.mpe *.webm *.mkv *.ts *.m2ts *.3gp *.3g2)",
        "missing_file_title": "缺少文件",
        "missing_file_body": "先选择一个或多个文档、图片或视频文件。",
        "no_pending_title": "没有待处理文件",
        "no_pending_body": "列表中的文件都已完成。双击文件可重新标记为待处理。",
        "ok_button": "知道了",
        "queued_marker": "排队",
        "pending_marker": "待处理",
        "running_marker": "处理中",
        "done_marker": "完成",
        "failed_marker": "失败",
        "stopped_marker": "已停止",
        "stage_exporting_editable_pptx": "正在导出编辑版 PPTX",
        "stage_converting_pdf": "正在转换源文档为 PDF",
        "stage_exporting_editable_pdf": "正在生成可编辑 PDF",
        "stage_rendering_images": "正在渲染图片页",
        "stage_building_image_pdf": "正在构建图片型 PDF",
        "stage_building_image_pptx": "正在构建图片型 PPTX",
        "stage_preparing_image_pptx_videos": "正在处理内嵌视频",
        "stage_reinserting_image_pptx_videos": "正在回贴内嵌视频",
        "stage_applying_watermark": "正在应用水印",
        "job_started": "任务开始：{count} 个文件。",
        "job_processing": "正在处理 {index}/{total}: {name}",
        "job_done": "导出完成：{name} -> {output} ({size})",
        "job_failed": "导出失败：{name}: {message}",
        "job_stopped": "任务已停止。未完成文件保持待处理。",
        "done_summary": "成功导出 {success}/{total} 个文件。",
        "queue_summary": "共 {count} 个文件",
        "queue_summary_empty": "未选择文件",
        "queue_summary_mixed_fixed_pdf": "共 {count} 个文件 · 含 DOCX / PDF 固定 PDF",
        "queue_summary_mixed_fixed_output": "共 {count} 个文件 · 不同类型按各自支持格式输出",
        "selection_ready": "已选择 {count} 个文件，可直接开始导出。",
        "selection_fixed_pdf": "已选择 {count} 个文件。当前 DOCX / PDF 固定导出为 PDF，队列中的 PPTX 仍按全局格式导出。",
        "selection_fixed_media": "已选择 {count} 个文件。当前图片/视频按原格式或 MP4 输出，文档 / PPTX 仍按全局格式导出。",
        "output_format_fixed": "{format}（固定）",
        "mixed_queue_pptx_body": "当前选择导出 PPTX，但队列中包含 DOCX 或 PDF。它们无法转成幻灯片，只会按你当前选择的形式导出为 PDF。",
        "close_button": "收起",
        "image_keep_videos_unavailable": "仅在输出格式为 PPTX、形式为图片化时可用。",
        "file_toggle_accessible": "{source} 文件处理开关",
        "file_toggle_tooltip": "{source} 文件。点击切换是否纳入本次处理。",
    },
    "en": {
        "window_title": "Document & Media Watermark Export",
        "eyebrow": "DOC MEDIA TOOLKIT",
        "title": "Document & Media Watermark Export",
        "subtitle": "Batch process DOCX / PDF / PPTX / images / videos with watermark export",
        "queue_title": "Files",
        "empty_queue": "Drop files here or click Add Files",
        "empty_queue_hint": "DOCX / PDF / PPTX / images / videos · multi-select",
        "pick_file": "Add Files",
        "remove_file": "Remove",
        "selected_summary": "{selected}/{total} selected",
        "settings_title": "Export Settings",
        "output_format_label": "Format",
        "output_format_pdf": "PDF",
        "output_format_pptx": "PPTX",
        "output_format_source": "Source",
        "output_mode_label": "Mode",
        "output_mode_editable": "Editable",
        "output_mode_image": "Image",
        "output_mode_not_applicable": "N/A",
        "output_format_document_locked_hint": "The current document is fixed to PDF output.",
        "output_format_media_locked_hint": "The current image/video is fixed to source format or MP4 output.",
        "media_mode_locked_hint": "Standalone images and videos do not use the mode setting; they export in source format or MP4.",
        "media_quality_locked_hint": "Standalone images use this for JPG / WEBP output quality while PNG keeps its source format. Standalone videos re-encode with watermark using Original / High / Balanced / Low.",
        "image_quality_mode_hint": "Quality is only used for image-based export; editable export does not use this quality preset.",
        "watermark_enable": "Add watermark",
        "watermark_switch": "Watermark",
        "watermark_toggle_on": "Watermark On",
        "watermark_toggle_off": "Watermark Off",
        "watermark_type": "Type",
        "watermark_type_text": "Text",
        "watermark_type_image": "Image",
        "watermark_text": "Watermark text",
        "watermark_image": "Watermark image",
        "watermark_pick_image": "Choose or drop image",
        "watermark_no_image": "No image selected",
        "watermark_image_width": "Image width",
        "watermark_image_missing_title": "Missing watermark image",
        "watermark_image_missing_body": "Image watermark mode needs an image first.",
        "watermark_preset": "Preset",
        "watermark_preset_hint": "Preset note",
        "watermark_text_placeholder": "Example: Company Confidential",
        "watermark_color": "Color",
        "watermark_font_size": "Font size",
        "watermark_spacing": "Spacing",
        "watermark_opacity": "Opacity",
        "watermark_angle": "Angle",
        "image_dpi": "Image quality",
        "image_quality_original": "Original",
        "image_quality_high": "High",
        "image_quality_balanced": "Balanced",
        "image_quality_low": "Low",
        "image_keep_videos": "Video keep",
        "image_keep_videos_button": "Reinsert",
        "image_keep_videos_on": "Reinsert On",
        "image_keep_videos_off": "Skip",
        "image_keep_videos_hint": "Image-based PPTX only: off means videos are not processed; on watermarks videos and tries to place them back. Failed reinsertion exports a _videos folder beside the output.",
        "image_keep_videos_warning": "Ungroup videos from other elements first, otherwise reinsertion may be misaligned.",
        "preview_title": "Preview",
        "preview_note": "Only the first few pages are rendered for visual review; video playback is not shown.",
        "preview_waiting": "Select a file on the left to generate a preview.",
        "preview_thumbnails_empty": "Select a file to show page thumbnails",
        "preview_manual": "Large file. Click Refresh to generate the preview.",
        "preview_loading": "Generating preview…",
        "preview_failed": "Preview failed",
        "preview_original": "Original",
        "preview_output": "Preview",
        "preview_refresh": "Refresh",
        "preview_prev": "Prev",
        "preview_next": "Next",
        "preview_empty_page": "No preview page available",
        "preview_page": "Preview page {index}/{rendered} · total slides {total}",
        "preview_background": "Bg",
        "details_title": "Status & Log",
        "advanced_settings": "Advanced",
        "advanced_expand": "Show advanced settings",
        "advanced_collapse": "Hide advanced settings",
        "log_expand": "Show log",
        "log_collapse": "Hide log",
        "log_hint": "Export logs will appear here.",
        "dependency_title": "Runtime dependencies",
        "dependency_ok": "All dependencies required by the current mode are available.",
        "dependency_ok_short": "Dependencies are available.",
        "dependency_missing_libreoffice": "PDF export engine is missing. On macOS, get Keynote for PPTX or Pages for DOCX; use LibreOffice if output fidelity is not acceptable. On Windows, install Microsoft Office/WPS or LibreOffice.",
        "dependency_keynote_fallback": "Dependencies are available: using Keynote fallback. Install LibreOffice if output fidelity is not acceptable.",
        "dependency_pages_fallback": "Dependencies are available: using Pages fallback. Install LibreOffice if output fidelity is not acceptable.",
        "dependency_missing_ffmpeg": "FFmpeg is missing. Image-based PPTX video reinsertion needs FFmpeg.",
        "dependency_install_ffmpeg": "Install FFmpeg",
        "dependency_install_libreoffice": "Install LibreOffice",
        "dependency_install_keynote": "Get Keynote",
        "dependency_install_pages": "Get Pages",
        "dependency_open_settings": "Open Automation Settings",
        "dependency_permission_keynote": "Keynote is installed, but Automation permission is missing. Allow Doc Media Toolkit or Terminal to control Keynote in System Settings first.",
        "dependency_permission_pages": "Pages is installed, but Automation permission is missing. Allow Doc Media Toolkit or Terminal to control Pages in System Settings first.",
        "fallback_warning": "Warning: PDF/image export engine is unavailable. Falling back to watermarked editable PPTX.",
        "font_check_title": "Font check",
        "font_check_waiting": "Choose a file to inspect source fonts.",
        "font_check_non_pptx": "Source font fallback only applies to PPTX. DOCX / PDF use the export engine as-is.",
        "font_check_ok": "No missing source fonts detected. The watermark font is bundled.",
        "font_check_missing": "{count} possibly missing font(s): {fonts}. Enable fallback to temporarily replace them with {family}.",
        "font_fix_on": "Fallback On",
        "font_fix_off": "Fallback Off",
        "status_ready": "Waiting to start. Add files to export.",
        "status_running": "Exporting, please wait.",
        "status_stopping": "Stop requested. The app will stop after the current file.",
        "current_processing": "Exporting",
        "status_done": "Export complete.",
        "status_failed": "Export failed.",
        "output_path_label": "Output file",
        "output_waiting": "Choose a file to show the default output path.",
        "log_waiting": "Waiting to start.",
        "run": "Export",
        "stop": "Stop",
        "stopping": "Stopping after current file",
        "help_button": "Help",
        "help_title": "How to use",
        "help_body": (
            "1. Add DOCX, PDF, PPTX, or standalone image/video files on the left.\n"
            "2. DOCX and PDF can only export to PDF. In mixed queues, they still export as PDF even if the global format is PPTX.\n"
            "3. Standalone images export in their original format. Standalone videos export as the original container when possible, otherwise MP4.\n"
            "4. Image quality only affects image-based PDF / image-based PPTX. Editable PDF / editable PPTX keep original document structure and source images where possible.\n"
            "5. The preview only renders the first few pages or the first media frame to verify watermark placement.\n"
            "6. Playable videos are removed by default in image mode; only image-based PPTX video reinsertion attempts to restore them.\n"
            "7. Intermediate artifacts are cleaned by default and outputs are saved beside the source file."
        ),
        "choose_title": "Choose Files",
        "choose_filter": "Supported Files (*.pptx *.docx *.pdf *.png *.jpg *.jpeg *.jpe *.webp *.mp4 *.m4v *.mov *.wmv *.asf *.avi *.mpg *.mpeg *.mpe *.webm *.mkv *.ts *.m2ts *.3gp *.3g2)",
        "missing_file_title": "Missing file",
        "missing_file_body": "Choose one or more document, image, or video files first.",
        "no_pending_title": "Nothing to export",
        "no_pending_body": "All files are already complete. Double-click a file to mark it pending again.",
        "ok_button": "OK",
        "queued_marker": "Queued",
        "pending_marker": "Pending",
        "running_marker": "Processing",
        "done_marker": "Done",
        "failed_marker": "Failed",
        "stopped_marker": "Stopped",
        "stage_exporting_editable_pptx": "Building editable PPTX",
        "stage_converting_pdf": "Converting source document to PDF",
        "stage_exporting_editable_pdf": "Building editable PDF",
        "stage_rendering_images": "Rendering page images",
        "stage_building_image_pdf": "Building image PDF",
        "stage_building_image_pptx": "Building image PPTX",
        "stage_preparing_image_pptx_videos": "Preparing embedded videos",
        "stage_reinserting_image_pptx_videos": "Reinserting embedded videos",
        "stage_applying_watermark": "Applying watermark",
        "job_started": "Job started: {count} file(s).",
        "job_processing": "Processing {index}/{total}: {name}",
        "job_done": "Done: {name} -> {output} ({size})",
        "job_failed": "Failed: {name}: {message}",
        "job_stopped": "Job stopped. Unfinished files remain pending.",
        "done_summary": "Exported {success}/{total} file(s) successfully.",
        "queue_summary": "{count} file(s)",
        "queue_summary_empty": "No files selected",
        "queue_summary_mixed_fixed_pdf": "{count} file(s) · includes DOCX/PDF fixed to PDF",
        "queue_summary_mixed_fixed_output": "{count} file(s) · each type uses its supported output",
        "selection_ready": "{count} file(s) selected. Ready to export.",
        "selection_fixed_pdf": "{count} file(s) selected. The current DOCX/PDF file is fixed to PDF, while queue PPTX files still follow the global target format.",
        "selection_fixed_media": "{count} file(s) selected. The current image/video exports in source format or MP4, while documents/PPTX still follow the global target format.",
        "output_format_fixed": "{format} (fixed)",
        "mixed_queue_pptx_body": "PPTX output is selected, but the queue also contains DOCX or PDF files. They cannot become slides and will export as PDF using the current mode.",
        "close_button": "Collapse",
        "image_keep_videos_unavailable": "Available only for image-based PPTX output.",
        "file_toggle_accessible": "{source} file processing toggle",
        "file_toggle_tooltip": "{source} file. Click to include or exclude it from this export.",
    },
}


def detect_language() -> str:
    return detect_system_language("PPTX_OUTPUT_WATERMARK_LANG")


def format_file_size(size_bytes: int) -> str:
    return format_user_file_size(size_bytes)


def resource_root() -> Path:
    if getattr(sys, "frozen", False):
        bundle_root = Path(
            getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent)
        )
        return bundle_root
    source_root = Path(__file__).resolve().parents[2]
    return source_root if (source_root / "assets").is_dir() else Path(sys.prefix)


def resource_path(*parts: str) -> Path:
    return resource_root().joinpath(*parts)


def preview_cache_key(path: Path) -> tuple[str, int, int]:
    stat = path.stat()
    return (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))


@dataclass(slots=True)
class GuiExportSettings:
    output_format: str
    output_mode: str
    dpi: int
    jpeg_quality: int
    video_quality_profile: str
    preserve_videos_in_image_pptx: bool
    video_encoder: str
    watermark_enabled: bool
    watermark_kind: str
    watermark_text: str
    watermark_image_path: Path | None
    watermark_image_width: int
    watermark_color: str
    watermark_opacity: float
    watermark_font_size: int
    watermark_spacing: int
    watermark_angle: float
    watermark_bold: bool = True
    watermark_margin: int = 90
    replace_source_fonts: bool = False
    source_font_family: str = field(default_factory=default_source_font_family)
    source_font_names: tuple[str, ...] = ()
    source_font_names_by_path: dict[str, tuple[str, ...]] | None = None

    def watermark(self) -> WatermarkOptions:
        return WatermarkOptions(
            enabled=self.watermark_enabled,
            kind=self.watermark_kind,
            text=self.watermark_text,
            image_path=self.watermark_image_path,
            image_width=self.watermark_image_width,
            angle=self.watermark_angle,
            color=self.watermark_color,
            opacity=self.watermark_opacity,
            font_size=self.watermark_font_size,
            spacing=self.watermark_spacing,
            margin=self.watermark_margin,
            bold=self.watermark_bold,
        )

    def export_options(self, input_path: Path) -> ExportOptions:
        source_font_names = self.source_font_names
        if self.source_font_names_by_path is not None:
            source_font_names = self.source_font_names_by_path.get(
                str(input_path.expanduser().resolve()),
                (),
            )
        output_format = self.effective_output_format_for_path(input_path)
        return ExportOptions(
            input_path=input_path,
            output_format=output_format,
            output_mode=self.output_mode,
            output_path=None,
            preserve_videos_in_image_pptx=self.preserve_videos_in_image_pptx,
            video_encoder=self.video_encoder,
            video_quality_profile=self.video_quality_profile,
            dpi=self.dpi,
            jpeg_quality=self.jpeg_quality,
            keep_artifacts=False,
            replace_source_fonts=self.replace_source_fonts
            and (bool(source_font_names) or self.source_font_names_by_path is None),
            source_font_family=self.source_font_family,
            source_font_names=source_font_names,
            watermark=self.watermark(),
        )

    def effective_output_format_for_path(self, input_path: Path) -> str:
        return effective_export_output_format(input_path, self.output_format)


@dataclass(slots=True)
class PreviewArtifacts:
    temp_root: Path | None
    source_key: tuple[str, int, int]
    original_paths: list[Path]
    preview_paths: list[Path]
    total_pages: int


@dataclass(slots=True)
class PreviewSource:
    key: tuple[str, int, int]
    temp_root: Path
    original_paths: list[Path]
    total_pages: int


class FileListWidget(QListWidget):
    filesDropped = Signal(list)
    deletePressed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._empty_message = ""
        self._empty_hint = ""
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    def set_empty_state(self, message: str, hint: str = "") -> None:
        self._empty_message = message
        self._empty_hint = hint
        self.viewport().update()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(
                url.isLocalFile()
                and url.toLocalFile()
                .lower()
                .endswith(SUPPORTED_WATERMARK_INPUT_EXTENSIONS)
                for url in urls
            ):
                event.acceptProposedAction()
                return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:  # noqa: N802
        self.dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        paths: list[Path] = []
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            path = Path(url.toLocalFile())
            if path.suffix.lower() in SUPPORTED_WATERMARK_INPUT_EXTENSIONS:
                paths.append(path)
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()
            return
        event.ignore()

    def keyPressEvent(self, event: QKeyEvent) -> None:  # noqa: N802
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.deletePressed.emit()
            return
        super().keyPressEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802, ANN001
        super().paintEvent(event)
        if self.count() != 0 or not self._empty_message:
            return

        viewport = self.viewport()
        rect = viewport.rect().adjusted(22, 18, -22, -18)
        if rect.width() <= 0 or rect.height() <= 0:
            return

        card_width = min(420, rect.width())
        card_height = min(116, rect.height())
        card_rect = QRect(0, 0, card_width, card_height)
        card_rect.moveCenter(rect.center())

        painter = QPainter(viewport)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#334155"), 1, Qt.PenStyle.DashLine))
        painter.setBrush(QColor("#0d1620"))
        painter.drawRoundedRect(card_rect, 14, 14)

        title_rect = card_rect.adjusted(14, 22, -14, -58)
        title_font = painter.font()
        title_font.setPixelSize(12)
        title_font.setBold(True)
        painter.setFont(title_font)
        painter.setPen(QColor("#cbd5e1"))
        painter.drawText(
            title_rect,
            int(
                Qt.AlignmentFlag.AlignHCenter
                | Qt.AlignmentFlag.AlignVCenter
                | Qt.TextFlag.TextWordWrap
            ),
            self._empty_message,
        )

        if self._empty_hint:
            hint_rect = card_rect.adjusted(12, 66, -12, -12)
            hint_font = painter.font()
            hint_font.setPixelSize(11)
            hint_font.setBold(False)
            painter.setFont(hint_font)
            painter.setPen(QColor("#64748b"))
            painter.drawText(
                hint_rect,
                int(
                    Qt.AlignmentFlag.AlignHCenter
                    | Qt.AlignmentFlag.AlignTop
                    | Qt.TextFlag.TextWordWrap
                ),
                self._empty_hint,
            )


class WatermarkImageDropButton(QPushButton):
    fileDropped = Signal(object)

    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if not url.isLocalFile():
                    continue
                path = Path(url.toLocalFile())
                if path.suffix.lower() in WATERMARK_IMAGE_INPUT_EXTENSIONS:
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:  # noqa: N802
        self.dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            path = Path(url.toLocalFile())
            if path.suffix.lower() in WATERMARK_IMAGE_INPUT_EXTENSIONS:
                self.fileDropped.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()


class ElidedLabel(QLabel):
    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.full_text = text
        QLabel.setText(self, text)

    def setText(self, text: str) -> None:  # noqa: N802
        self.set_full_text(text)

    def set_full_text(self, text: str) -> None:
        self.full_text = text
        self.update_elide()

    def resizeEvent(self, event) -> None:  # noqa: N802, ANN001
        super().resizeEvent(event)
        self.update_elide()

    def update_elide(self) -> None:
        width = max(20, self.width())
        QLabel.setText(
            self,
            self.fontMetrics().elidedText(
                self.full_text, Qt.TextElideMode.ElideMiddle, width
            ),
        )


class CleanComboBox(QComboBox):
    def paintEvent(self, event) -> None:  # noqa: N802, ANN001
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        color = QColor("#f8fafc" if self.isEnabled() else "#94a3b8")
        pen = QPen(color, 2)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        center_x = self.width() - 15
        center_y = self.height() // 2 + 1
        painter.drawLine(center_x - 5, center_y - 4, center_x, center_y + 1)
        painter.drawLine(center_x, center_y + 1, center_x + 5, center_y - 4)


class PreviewCanvas(QLabel):
    DEFAULT_ASPECT_RATIO = 16.0 / 9.0
    FRAME_PADDING = 2
    MIN_PREVIEW_HEIGHT = 220
    PREFERRED_WIDTH = 480

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__("", parent)
        self._preview_pixmap = QPixmap()
        self._message = text
        self._sample_watermark: WatermarkOptions | None = None
        self._sample_theme: dict[str, str] = PREVIEW_BACKGROUND_THEMES[0]
        self._watermark_font_family = BUNDLED_WATERMARK_FONT_FAMILY
        self.setAlignment(Qt.AlignCenter)

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def current_aspect_ratio(self) -> float:
        if not self._preview_pixmap.isNull():
            return self._preview_pixmap.width() / max(
                1.0, self._preview_pixmap.height()
            )
        return self.DEFAULT_ASPECT_RATIO

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        content_width = max(1, width - self.FRAME_PADDING)
        return (
            int(round(content_width / self.current_aspect_ratio())) + self.FRAME_PADDING
        )

    def sizeHint(self) -> QSize:  # noqa: N802
        width = max(self.PREFERRED_WIDTH, super().sizeHint().width())
        return QSize(width, self.heightForWidth(width))

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        width = 390
        return QSize(width, self.heightForWidth(width))

    def set_watermark_font_family(self, family: str) -> None:
        self._watermark_font_family = family
        self.update()

    def _fit_rect_for_aspect(self, rect: QRect, aspect_ratio: float) -> QRect:
        if rect.width() <= 0 or rect.height() <= 0:
            return QRect()
        if aspect_ratio <= 0:
            aspect_ratio = 16.0 / 9.0

        target_width = rect.width()
        target_height = int(round(target_width / aspect_ratio))
        if target_height > rect.height():
            target_height = rect.height()
            target_width = int(round(target_height * aspect_ratio))

        x = rect.x() + (rect.width() - target_width) // 2
        y = rect.y()
        return QRect(x, y, max(1, target_width), max(1, target_height))

    def set_preview_pixmap(self, pixmap: QPixmap) -> None:
        self._preview_pixmap = pixmap
        self._message = ""
        self._sample_watermark = None
        self.updateGeometry()
        self.update()

    def clear_preview(
        self,
        text: str,
        sample_watermark: WatermarkOptions | None = None,
        sample_theme: dict[str, str] | None = None,
    ) -> None:
        self._preview_pixmap = QPixmap()
        self._message = text
        self._sample_watermark = sample_watermark
        self._sample_theme = sample_theme or PREVIEW_BACKGROUND_THEMES[0]
        self.updateGeometry()
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802, ANN001
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        outer_rect = self.contentsRect().adjusted(2, 2, -3, -3)
        if outer_rect.width() <= 0 or outer_rect.height() <= 0:
            painter.end()
            return
        rect = outer_rect.adjusted(2, 2, -2, -2)

        if not self._preview_pixmap.isNull():
            aspect_ratio = self._preview_pixmap.width() / max(
                1, self._preview_pixmap.height()
            )
        else:
            aspect_ratio = 16.0 / 9.0
        page_rect = self._fit_rect_for_aspect(rect, aspect_ratio)

        if not self._preview_pixmap.isNull():
            painter.drawPixmap(
                page_rect.adjusted(1, 1, -1, -1),
                self._preview_pixmap,
                self._preview_pixmap.rect(),
            )
        elif self._sample_watermark is not None:
            self._draw_sample_page(painter, page_rect, self._sample_theme)
            if self._sample_watermark.enabled:
                self._draw_sample_watermark(painter, page_rect, self._sample_watermark)
            painter.setPen(QColor("#64748b"))
            painter.drawText(
                page_rect.adjusted(18, 18, -18, -18),
                Qt.AlignBottom | Qt.AlignHCenter | Qt.TextWordWrap,
                self._message,
            )
        else:
            painter.fillRect(page_rect, QColor("#0b1017"))
            painter.setPen(QColor("#64748b"))
            painter.drawText(
                page_rect.adjusted(18, 18, -18, -18),
                Qt.AlignCenter | Qt.TextWordWrap,
                self._message,
            )

        painter.end()

    def _draw_sample_page(
        self, painter: QPainter, page_rect: QRect, theme: dict[str, str]
    ) -> None:
        painter.fillRect(page_rect, QColor(theme["bg"]))

        pad = max(18, int(page_rect.width() * 0.035))
        title_h = max(12, int(page_rect.height() * 0.045))
        line_h = max(7, int(page_rect.height() * 0.023))
        block_h = max(26, int(page_rect.height() * 0.16))

        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#f97316"))
        painter.drawRoundedRect(
            page_rect.x() + pad, page_rect.y() + pad, title_h * 3, title_h, 4, 4
        )
        painter.setBrush(QColor(theme["line"]))
        painter.drawRoundedRect(
            page_rect.x() + pad,
            page_rect.y() + pad + title_h + 12,
            page_rect.width() // 3,
            line_h,
            3,
            3,
        )
        painter.setBrush(QColor(theme["muted"]))
        painter.drawRoundedRect(
            page_rect.x() + pad,
            page_rect.y() + pad + title_h + 28,
            page_rect.width() // 4,
            line_h,
            3,
            3,
        )

        chart_y = page_rect.y() + page_rect.height() // 2 - block_h // 2
        chart_w = max(40, page_rect.width() // 7)
        for idx, color in enumerate(("#2563eb", "#facc15", "#ef4444")):
            rect = QRect(
                page_rect.x()
                + page_rect.width() // 2
                - chart_w * 2
                + idx * int(chart_w * 1.35),
                chart_y + idx * 8,
                chart_w,
                block_h - idx * 10,
            )
            painter.setBrush(QColor(color))
            painter.drawRoundedRect(rect, 10, 10)

        right_x = page_rect.x() + page_rect.width() // 2 + pad
        right_width = page_rect.width() // 2
        painter.setBrush(QColor(theme["panel"]))
        painter.drawRoundedRect(
            right_x, page_rect.y() + pad, right_width // 3, line_h, 3, 3
        )
        painter.setBrush(QColor(theme["muted"]))
        painter.drawRoundedRect(
            right_x, page_rect.y() + pad + 18, right_width // 4, line_h, 3, 3
        )
        painter.setBrush(QColor(theme["panel_soft"]))
        painter.drawRoundedRect(right_x, chart_y, right_width - pad * 2, block_h, 8, 8)

    def _draw_sample_watermark(
        self, painter: QPainter, page_rect: QRect, options: WatermarkOptions
    ) -> None:
        if options.kind == "image":
            self._draw_sample_image_watermark(painter, page_rect, options)
            return

        text = options.text.strip() or DEFAULT_WATERMARK_TEXT
        scale = max(0.28, page_rect.width() / 1600.0)
        font = QFont()
        families = [
            self._watermark_font_family,
            BUNDLED_WATERMARK_FONT_FAMILY,
            "PingFang SC",
            "Noto Sans CJK SC",
            "Microsoft YaHei",
            "Arial Unicode MS",
            "Arial",
        ]
        try:
            font.setFamilies(families)
        except AttributeError:
            font.setFamily(self._watermark_font_family)
        font.setPixelSize(max(11, int(options.font_size * scale)))
        font.setBold(options.bold)
        painter.save()
        painter.setClipRect(page_rect)
        painter.setFont(font)
        color = QColor(options.color)
        color.setAlphaF(max(0.0, min(1.0, options.opacity)))
        painter.setPen(color)

        metrics = painter.fontMetrics()
        text_width = max(1, metrics.horizontalAdvance(text))
        text_height = max(1, metrics.height())
        step = max(110, int(text_width + options.spacing * scale))
        row_step = max(88, int(text_height + options.spacing * scale * 0.72))
        start_x = page_rect.x() - step
        end_x = page_rect.right() + step
        start_y = page_rect.y() - row_step
        end_y = page_rect.bottom() + row_step

        for y in range(start_y, end_y + 1, row_step):
            for x in range(start_x, end_x + 1, step):
                painter.save()
                painter.translate(x, y)
                painter.rotate(options.angle)
                painter.drawText(-text_width // 2, text_height // 2, text)
                painter.restore()
        painter.restore()

    def _draw_sample_image_watermark(
        self, painter: QPainter, page_rect: QRect, options: WatermarkOptions
    ) -> None:
        if options.image_path is None:
            return
        pixmap = QPixmap(str(options.image_path))
        if pixmap.isNull():
            return

        scale = max(0.28, page_rect.width() / 1600.0)
        image_width = max(12, int(options.image_width * scale))
        image_height = max(
            1, int(round(image_width * pixmap.height() / max(1, pixmap.width())))
        )
        spacing = max(60, int(options.spacing * scale))
        step_x = max(90, image_width + spacing)
        step_y = max(70, image_height + int(spacing * 0.72))

        painter.save()
        painter.setClipRect(page_rect)
        painter.setOpacity(max(0.0, min(1.0, options.opacity)))
        for y in range(page_rect.y() - step_y, page_rect.bottom() + step_y + 1, step_y):
            for x in range(
                page_rect.x() - step_x, page_rect.right() + step_x + 1, step_x
            ):
                painter.save()
                painter.translate(x, y)
                painter.rotate(options.angle)
                painter.drawPixmap(
                    -image_width // 2,
                    -image_height // 2,
                    image_width,
                    image_height,
                    pixmap,
                )
                painter.restore()
        painter.restore()


class StyledDialog(QDialog):
    def __init__(self, parent: QWidget, title: str, message: str) -> None:
        super().__init__(parent)
        self.setModal(True)
        self.setWindowTitle(title)
        self.setMinimumWidth(440)
        self.setObjectName("styledDialog")

        # Start invisible to hide macOS window cascading jump
        self.setWindowOpacity(0.0)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        title_label = QLabel(title)
        title_label.setObjectName("dialogTitle")
        layout.addWidget(title_label)

        body_label = QLabel(message)
        body_label.setObjectName("dialogBody")
        body_label.setWordWrap(True)
        layout.addWidget(body_label)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button = QPushButton("OK")
        button.setObjectName("dialogPrimaryButton")
        button.setFixedHeight(40)
        button.clicked.connect(self.accept)
        button_row.addWidget(button)
        layout.addLayout(button_row)

        self.setStyleSheet(
            """
            QDialog#styledDialog { background: #0b1017; }
            QLabel#dialogTitle { color: #f8fafc; font-size: 16px; font-weight: 600; }
            QLabel#dialogBody { color: #cbd5e1; font-size: 11px; line-height: 1.35; }
            QPushButton {
                background: #f97316;
                color: #ffffff;
                border: 1px solid #fb923c;
                border-radius: 10px;
                padding: 6px 14px;
                min-width: 90px;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton:hover { background: #ea580c; }
            """
            + SHARED_DIALOG_QSS
        )

        # We use QTimer in showEvent instead of __init__ to beat macOS window cascading.

    def showEvent(self, event) -> None:  # noqa: N802, ANN001
        super().showEvent(event)
        from PySide6.QtCore import QTimer

        QTimer.singleShot(0, self._force_center)
        QTimer.singleShot(16, self._force_center)

    def _force_center(self) -> None:
        self.adjustSize()
        parent = self.parentWidget()
        anchor = None
        if parent is not None and parent.isVisible():
            anchor = (
                parent.window()
                if parent.window() and parent.window().isVisible()
                else parent
            )
        elif QApplication.activeWindow() is not None:
            anchor = QApplication.activeWindow()

        if anchor is not None:
            if anchor.isWindow():
                parent_rect = anchor.frameGeometry()
            else:
                parent_rect = QRect(anchor.mapToGlobal(QPoint(0, 0)), anchor.size())
        else:
            screen = self.screen() or QApplication.primaryScreen()
            if screen is None:
                self.setWindowOpacity(1.0)
                return
            parent_rect = screen.availableGeometry()

        target_x = parent_rect.x() + (parent_rect.width() - self.width()) // 2
        target_y = parent_rect.y() + (parent_rect.height() - self.height()) // 2
        self.move(target_x, target_y)

        # Make visible after moving to center
        self.setWindowOpacity(1.0)


class ExportWorker(QObject):
    finished = Signal(list, list, bool)
    progress = Signal(int, str)
    log = Signal(str)
    fileStarted = Signal(object)
    fileCompleted = Signal(object, object, int)
    fileFailed = Signal(object, str)

    def __init__(
        self, input_paths: list[Path], settings: GuiExportSettings, text: dict[str, str]
    ) -> None:
        super().__init__()
        self.input_paths = input_paths
        self.settings = settings
        self.text = text
        self.cancel_requested = False
        self.current_file_index = 0

    def cancel(self) -> None:
        self.cancel_requested = True
        terminate_active_processes()

    def run(self) -> None:
        results: list[tuple[Path, Path, int]] = []
        failures: list[tuple[Path, str]] = []
        total = len(self.input_paths)
        stopped = False
        for file_index, input_path in enumerate(self.input_paths, start=1):
            if self.cancel_requested:
                stopped = True
                break
            self.current_file_index = file_index
            self.fileStarted.emit(input_path)
            self.log.emit(
                self.text["job_processing"].format(
                    index=file_index, total=total, name=input_path.name
                )
            )
            self.progress.emit(
                int(((file_index - 1) / max(1, total)) * 100), input_path.name
            )
            try:
                options = self.settings.export_options(input_path)
                try:
                    output_path = export_document(options, logger=self._stage_logger)
                except RuntimeError as inner_exc:
                    if self.cancel_requested:
                        stopped = True
                        break
                    err_str = str(inner_exc).lower()
                    is_pdf_fail = (
                        "libreoffice" in err_str
                        or "powerpoint" in err_str
                        or "missing required runtime" in err_str
                    )
                    if options.output_format != "pptx" and is_pdf_fail:
                        self.log.emit(
                            self.text.get(
                                "fallback_warning",
                                "Warning: PDF/Image engine failed. Falling back to watermarked PPTX...",
                            )
                        )
                        fallback_options = self.settings.export_options(input_path)
                        fallback_options.output_format = "pptx"
                        fallback_options.output_mode = "editable"
                        output_path = export_document(
                            fallback_options, logger=self._stage_logger
                        )
                    else:
                        raise inner_exc

                size_bytes = output_path.stat().st_size if output_path.exists() else 0
                results.append((input_path, output_path, size_bytes))
                self.fileCompleted.emit(input_path, output_path, size_bytes)
                self.log.emit(
                    self.text["job_done"].format(
                        name=input_path.name,
                        output=output_path,
                        size=format_file_size(size_bytes),
                    )
                )
            except Exception as exc:
                if self.cancel_requested:
                    stopped = True
                    break
                failures.append((input_path, str(exc)))
                self.fileFailed.emit(input_path, str(exc))
                self.log.emit(
                    self.text["job_failed"].format(
                        name=input_path.name, message=str(exc)
                    )
                )
        if self.cancel_requested and len(results) + len(failures) < total:
            stopped = True
        self.finished.emit(results, failures, stopped)

    def _stage_logger(self, message: str) -> None:
        percent = self._percent_for_stage(message)
        if percent is not None:
            self.progress.emit(percent, self._localize_stage(message))
        self.log.emit(message)

    def _percent_for_stage(self, message: str) -> int | None:
        total = max(1, len(self.input_paths))
        file_offset = self.current_file_index - 1
        if message.startswith("Converting ") and message.endswith(" to PDF"):
            stage_percent = 40
        else:
            stage_percent = {
                "Exporting editable PPTX": 20,
                "Exporting editable PDF": 65,
                "Rendering PDF pages to images": 75,
                "Building image-based PDF": 92,
                "Building image-based PPTX": 92,
                "Preparing embedded videos for image-based PPTX": 95,
                "Reinserting embedded videos into image-based PPTX": 97,
            }.get(message)
        if stage_percent is None:
            return None
        return min(99, int(((file_offset + stage_percent / 100.0) / total) * 100))

    def _localize_stage(self, message: str) -> str:
        if message.startswith("Converting ") and message.endswith(" to PDF"):
            return self.text["stage_converting_pdf"]
        mapping = {
            "Exporting editable PPTX": self.text["stage_exporting_editable_pptx"],
            "Exporting editable PDF": self.text["stage_exporting_editable_pdf"],
            "Rendering PDF pages to images": self.text["stage_rendering_images"],
            "Building image-based PDF": self.text["stage_building_image_pdf"],
            "Building image-based PPTX": self.text["stage_building_image_pptx"],
            "Preparing embedded videos for image-based PPTX": self.text[
                "stage_preparing_image_pptx_videos"
            ],
            "Reinserting embedded videos into image-based PPTX": self.text[
                "stage_reinserting_image_pptx_videos"
            ],
        }
        return mapping.get(message, message)


class PreviewWorker(QObject):
    finished = Signal(int, object, object, object)
    failed = Signal(int, object, str)

    def __init__(
        self,
        request_id: int,
        input_path: Path,
        settings: GuiExportSettings,
        source_key: tuple[str, int, int],
        source: PreviewSource | None,
    ) -> None:
        super().__init__()
        self.request_id = request_id
        self.input_path = input_path
        self.settings = settings
        self.source_key = source_key
        self.source = source

    def run(self) -> None:
        created_source: PreviewSource | None = None
        preview_root: Path | None = None
        temp_pdf: Path | None = None
        owns_temp_pdf = False
        try:
            source = self.source
            if source is None:
                source_root = create_runtime_temp_dir(
                    "pptx_output_watermark_preview_source_",
                    purpose="preview_source_pages",
                )
                input_suffix = self.input_path.suffix.lower()
                if input_suffix in IMAGE_EXTENSIONS:
                    preview_image = source_root / f"original{input_suffix}"
                    shutil.copy2(self.input_path, preview_image)
                    source = PreviewSource(
                        key=self.source_key,
                        temp_root=source_root,
                        original_paths=[preview_image],
                        total_pages=1,
                    )
                    created_source = source
                elif input_suffix in VIDEO_EXTENSIONS:
                    preview_image = source_root / "poster.jpg"
                    extract_video_poster_frame(self.input_path, preview_image)
                    source = PreviewSource(
                        key=self.source_key,
                        temp_root=source_root,
                        original_paths=[preview_image],
                        total_pages=1,
                    )
                    created_source = source
                elif input_suffix == ".pdf":
                    temp_pdf = self.input_path
                else:
                    temp_pdf = convert_document_to_pdf(
                        self.input_path, timeout_seconds=900
                    )
                    owns_temp_pdf = True
                if source is None:
                    total_pages = len(open_pdf_reader(str(temp_pdf)).pages)
                    render_count = min(total_pages, PREVIEW_PAGE_LIMIT)
                    rendered_dir = source_root / "original"
                    rendered = batch_render_pdf_slides(
                        temp_pdf,
                        num_slides=render_count,
                        output_dir=rendered_dir,
                        dpi=96,
                        jpeg_quality=78,
                        max_edge=1600,
                        max_pixels=1_800_000,
                    )
                    source = PreviewSource(
                        key=self.source_key,
                        temp_root=source_root,
                        original_paths=[rendered[idx] for idx in sorted(rendered)],
                        total_pages=total_pages,
                    )
                    created_source = source

            if self.settings.watermark_enabled:
                preview_root = create_runtime_temp_dir(
                    "pptx_output_watermark_preview_overlay_",
                    purpose="preview_overlay_pages",
                )
                preview_paths = []
                for image_path in source.original_paths:
                    output_path = preview_root / image_path.name
                    preview_paths.append(
                        apply_watermark_to_image(
                            image_path, output_path, self.settings.watermark()
                        )
                    )
            else:
                preview_paths = list(source.original_paths)
            artifacts = PreviewArtifacts(
                temp_root=preview_root,
                source_key=source.key,
                original_paths=source.original_paths,
                preview_paths=preview_paths,
                total_pages=source.total_pages,
            )
            self.finished.emit(self.request_id, self.input_path, artifacts, source)
            created_source = None
        except Exception as exc:
            if preview_root is not None:
                shutil.rmtree(preview_root, ignore_errors=True)
            if created_source is not None:
                shutil.rmtree(created_source.temp_root, ignore_errors=True)
            self.failed.emit(self.request_id, self.input_path, str(exc))
        finally:
            if temp_pdf is not None and owns_temp_pdf:
                try:
                    temp_pdf.unlink()
                except FileNotFoundError:
                    pass
                except Exception:
                    pass
                try:
                    temp_pdf.parent.rmdir()
                except OSError:
                    pass


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        app = QApplication.instance()
        if app is not None:
            configure_ui_font(app)
        super().__init__()
        cleanup_stale_runtime_entries()
        self.language = detect_language()
        self.text = STRINGS[self.language]
        self.setWindowTitle(self.text["window_title"])
        self.setMinimumSize(880, 560)
        self.resize(960, 620)

        icon_path = resource_path("assets", "app_icon.png")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.watermark_font_family = self.load_bundled_watermark_font()
        self.input_paths: list[Path] = []
        self.checked_paths: set[Path] = set()
        self.file_statuses: dict[Path, str] = {}
        self.file_outputs: dict[Path, Path] = {}
        self.is_running = False
        self.worker_thread: QThread | None = None
        self.worker: ExportWorker | None = None
        self.preview_thread: QThread | None = None
        self.preview_worker: PreviewWorker | None = None
        self.preview_request_id = 0
        self.preview_dirty = False
        self.preview_source_cache: dict[tuple[str, int, int], PreviewSource] = {}
        self.preview_source_cleanup_pending: set[Path] = set()
        self.font_report_cache: dict[tuple[str, int, int], FontScanResult] = {}
        self.current_preview: PreviewArtifacts | None = None
        self.current_preview_page = 0
        self.preview_mode = "preview"
        self.watermark_image_path: Path | None = None
        self.current_font_report: FontScanResult | None = None
        self.current_missing_fonts: tuple[str, ...] = ()
        self.available_font_families = self.collect_available_font_families()
        self.pdf_engine_download_url = LIBREOFFICE_DOWNLOAD_URL

        central = QWidget()
        central.setObjectName("central")
        self.content_widget = central
        central.installEventFilter(self)
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(14, 10, 14, 10)
        root.setSpacing(8)

        header = QFrame()
        header.setObjectName("headerCard")
        header.setFixedHeight(76)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 7, 16, 7)
        header_layout.setSpacing(10)

        title_stack_widget = QWidget(header)
        title_stack = QVBoxLayout(title_stack_widget)
        title_stack.setContentsMargins(0, 0, 0, 0)
        title_stack.setSpacing(3)

        eyebrow = QLabel(self.text["eyebrow"])
        eyebrow.setObjectName("eyebrow")
        eyebrow.setFixedHeight(16)
        title_stack.addWidget(eyebrow)

        title = QLabel(self.text["title"])
        title.setObjectName("title")
        title.setFixedHeight(26)
        title_stack.addWidget(title)

        subtitle = QLabel(self.text["subtitle"])
        subtitle.setObjectName("subtitle")
        subtitle.setFixedHeight(16)
        title_stack.addWidget(subtitle)
        title_stack.addStretch(1)
        header_layout.addWidget(title_stack_widget, 1)

        self.help_button = QPushButton("?")
        self.help_button.setObjectName("helpIconButton")
        self.help_button.setToolTip(self.text["help_button"])
        self.help_button.setFixedSize(30, 30)
        self.help_button.clicked.connect(self.show_help)
        header_layout.addWidget(self.help_button, 0, Qt.AlignRight | Qt.AlignVCenter)
        root.addWidget(header)

        body_row = QHBoxLayout()
        body_row.setSpacing(10)
        left_column = QVBoxLayout()
        left_column.setSpacing(8)
        left_column.setContentsMargins(0, 0, 0, 0)
        center_column = QVBoxLayout()
        center_column.setSpacing(8)
        center_column.setContentsMargins(0, 0, 0, 0)
        right_column = QVBoxLayout()
        right_column.setSpacing(12)
        right_column.setContentsMargins(0, 0, 0, 0)

        left_card = QFrame()
        left_card.setObjectName("leftCard")
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(8)

        header_row_height = 30
        control_row_height = 30

        queue_header = QHBoxLayout()
        queue_header.setSpacing(10)
        queue_title = QLabel(self.text["queue_title"])
        queue_title.setObjectName("sectionTitle")
        queue_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        queue_title.setFixedHeight(header_row_height)
        queue_header.addWidget(queue_title, 1)
        self.queue_count_label = QLabel(self.text["queue_summary_empty"])
        self.queue_count_label.setObjectName("estimatePill")
        self.queue_count_label.setAlignment(Qt.AlignCenter)
        self.queue_count_label.setFixedHeight(header_row_height)
        self.queue_count_label.setMinimumWidth(88 if self.language == "zh" else 104)
        queue_header.addWidget(self.queue_count_label)
        left_layout.addLayout(queue_header)

        queue_buttons = QHBoxLayout()
        queue_buttons.setSpacing(7)
        self.pick_button = QPushButton(self.text["pick_file"])
        self.pick_button.setObjectName("secondaryButton")
        self.pick_button.setMinimumHeight(control_row_height)
        self.pick_button.clicked.connect(self.choose_file)
        self.remove_button = QPushButton(self.text["remove_file"])
        self.remove_button.setObjectName("secondaryButton")
        self.remove_button.setMinimumHeight(control_row_height)
        self.remove_button.clicked.connect(self.remove_selected_files)
        queue_buttons.addWidget(self.pick_button, 1)
        queue_buttons.addWidget(self.remove_button, 1)
        left_layout.addLayout(queue_buttons)

        self.file_list = FileListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.set_empty_state(
            self.text["empty_queue"], self.text["empty_queue_hint"]
        )
        self.file_list.filesDropped.connect(self.set_files)
        self.file_list.deletePressed.connect(self.remove_selected_files)
        self.file_list.itemSelectionChanged.connect(self.on_selection_changed)
        self.file_list.itemDoubleClicked.connect(self.mark_item_pending)
        self.file_list.setMinimumHeight(204)
        left_layout.addWidget(self.file_list, 1)

        self.selection_summary_label = QLabel()
        self.selection_summary_label.setObjectName("selectionSummary")
        self.selection_summary_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        left_layout.addWidget(self.selection_summary_label)
        left_column.addWidget(left_card, 1)

        right_card = QFrame()
        right_card.setObjectName("rightCard")
        settings_layout = QVBoxLayout(right_card)
        settings_layout.setContentsMargins(14, 12, 14, 14)
        settings_layout.setSpacing(8)

        top_control_height = control_row_height

        settings_header = QHBoxLayout()
        settings_header.setContentsMargins(0, 0, 0, 0)
        settings_header.setSpacing(10)
        settings_header.setAlignment(Qt.AlignVCenter)
        settings_title = QLabel(self.text["settings_title"])
        settings_title.setObjectName("sectionTitle")
        settings_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        settings_title.setFixedHeight(header_row_height)
        settings_header.addWidget(settings_title, 1)

        format_label = QLabel(self.text["output_format_label"])
        format_label.setObjectName("fieldLabel")
        format_label.setContentsMargins(0, 0, 0, 0)
        self.output_format_select = CleanComboBox()
        self.output_format_select.setView(QListView())
        self.output_format_select.view().setObjectName("comboPopup")
        self.output_format_select.addItem(self.text["output_format_pdf"], "pdf")
        self.output_format_select.addItem(self.text["output_format_pptx"], "pptx")
        self.output_format_select.addItem(self.text["output_format_source"], "source")
        self.output_format_select.setMinimumHeight(top_control_height)
        self.output_format_select.setMinimumContentsLength(2)
        self.output_format_select.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.output_format_select.currentIndexChanged.connect(self.on_settings_changed)

        mode_label = QLabel(self.text["output_mode_label"])
        mode_label.setObjectName("fieldLabel")
        mode_label.setContentsMargins(0, 0, 0, 0)
        self.output_mode_select = CleanComboBox()
        self.output_mode_select.setView(QListView())
        self.output_mode_select.view().setObjectName("comboPopup")
        self.output_mode_select.addItem(self.text["output_mode_editable"], "editable")
        self.output_mode_select.addItem(self.text["output_mode_image"], "image")
        self.output_mode_select.addItem(self.text["output_mode_not_applicable"], "na")
        self.output_mode_select.setMinimumHeight(top_control_height)
        self.output_mode_select.setMinimumContentsLength(2)
        self.output_mode_select.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.output_mode_select.currentIndexChanged.connect(self.on_settings_changed)
        self._primary_context_lock = False
        self._preferred_output_format = "pdf"
        self._preferred_output_mode = "editable"

        dpi_label = QLabel(self.text["image_dpi"])
        dpi_label.setObjectName("fieldLabel")
        dpi_label.setContentsMargins(0, 0, 0, 0)
        self.dpi_select = CleanComboBox()
        self.dpi_select.setView(QListView())
        self.dpi_select.view().setObjectName("comboPopup")
        quality_options = (
            ("image_quality_original", 300, 94),
            ("image_quality_high", 240, 90),
            ("image_quality_balanced", 200, 85),
            ("image_quality_low", 160, 78),
        )
        for label_key, dpi_value, jpeg_quality in quality_options:
            self.dpi_select.addItem(
                self.text[label_key],
                (dpi_value, jpeg_quality),
            )
        self.dpi_select.setCurrentIndex(1)
        self.dpi_select.setMinimumHeight(top_control_height)
        self.dpi_select.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.dpi_select.currentIndexChanged.connect(self.on_settings_changed)

        keep_videos_label = QLabel(self.text["image_keep_videos"])
        keep_videos_label.setObjectName("fieldLabel")
        keep_videos_label.setContentsMargins(0, 0, 0, 0)
        self.keep_videos_checkbox = QPushButton(self.text["image_keep_videos_button"])
        self.keep_videos_checkbox.setObjectName("toggleButton")
        self.keep_videos_checkbox.setCheckable(True)
        self.keep_videos_checkbox.setChecked(False)
        self.keep_videos_checkbox.setMinimumHeight(top_control_height)
        self.keep_videos_checkbox.setFixedWidth(92 if self.language == "zh" else 108)
        self.keep_videos_checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        self.keep_videos_checkbox.clicked.connect(self.on_settings_changed)

        self.run_button = QPushButton(self.text["run"])
        self.run_button.setObjectName("primaryButton")
        self.run_button.setProperty("mode", "run")
        self.run_button.setMinimumHeight(top_control_height)
        self.run_button.setFixedWidth(124 if self.language == "zh" else 138)
        self.run_button.setMinimumWidth(124 if self.language == "zh" else 138)
        self.run_button.setDisabled(True)
        self.run_button.clicked.connect(self.run_or_stop)
        settings_header.addWidget(self.run_button, 0, Qt.AlignRight | Qt.AlignVCenter)
        settings_layout.addLayout(settings_header)

        top_controls = QGridLayout()
        top_controls.setContentsMargins(0, 0, 0, 0)
        top_controls.setHorizontalSpacing(10)
        top_controls.setVerticalSpacing(8)

        format_group = QWidget()
        format_group.setObjectName("controlGroup")
        format_group.setProperty("plain", True)
        format_group_layout = QHBoxLayout(format_group)
        format_group_layout.setContentsMargins(0, 0, 0, 0)
        format_group_layout.setSpacing(4)
        format_label.setFixedWidth(28 if self.language == "zh" else 44)
        format_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.output_format_label_widget = format_label
        format_group_layout.addWidget(format_label)
        format_group_layout.addWidget(self.output_format_select, 1)
        format_group.setMinimumWidth(0)
        format_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        mode_group = QWidget()
        mode_group.setObjectName("controlGroup")
        mode_group.setProperty("plain", True)
        mode_group_layout = QHBoxLayout(mode_group)
        mode_group_layout.setContentsMargins(0, 0, 0, 0)
        mode_group_layout.setSpacing(4)
        mode_label.setFixedWidth(28 if self.language == "zh" else 40)
        mode_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.output_mode_label_widget = mode_label
        mode_group_layout.addWidget(mode_label)
        mode_group_layout.addWidget(self.output_mode_select, 1)
        mode_group.setMinimumWidth(0)
        mode_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        dpi_group = QWidget()
        dpi_group.setObjectName("controlGroup")
        dpi_group.setProperty("plain", True)
        dpi_group_layout = QHBoxLayout(dpi_group)
        dpi_group_layout.setContentsMargins(0, 0, 0, 0)
        dpi_group_layout.setSpacing(4)
        dpi_label.setText("质量" if self.language == "zh" else "Quality")
        dpi_label.setFixedWidth(28 if self.language == "zh" else 54)
        dpi_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.output_quality_label_widget = dpi_label
        dpi_group_layout.addWidget(dpi_label)
        dpi_group_layout.addWidget(self.dpi_select, 1)
        dpi_group.setMinimumWidth(116 if self.language == "zh" else 166)
        dpi_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        keep_videos_label.setToolTip(self.text["image_keep_videos_warning"])
        keep_videos_group = QWidget()
        keep_videos_group.setObjectName("controlGroup")
        keep_videos_group.setProperty("plain", True)
        keep_videos_group_layout = QHBoxLayout(keep_videos_group)
        keep_videos_group_layout.setContentsMargins(0, 0, 6, 0)
        keep_videos_group_layout.setSpacing(4)
        keep_videos_label.setText(
            "内嵌视频" if self.language == "zh" else "Embedded video"
        )
        keep_videos_label.setFixedWidth(56 if self.language == "zh" else 96)
        keep_videos_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.keep_videos_label_widget = keep_videos_label
        keep_videos_group_layout.addWidget(keep_videos_label)
        keep_videos_group_layout.addWidget(self.keep_videos_checkbox)
        self.keep_videos_group = keep_videos_group
        keep_videos_group.setMinimumWidth(128 if self.language == "zh" else 176)
        keep_videos_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        format_group.setMinimumWidth(139 if self.language == "zh" else 139)
        mode_group.setMinimumWidth(136 if self.language == "zh" else 136)
        dpi_group.setMinimumWidth(120 if self.language == "zh" else 146)
        keep_videos_group.setMinimumWidth(166 if self.language == "zh" else 218)

        top_controls.addWidget(format_group, 0, 0)
        top_controls.addWidget(mode_group, 0, 1)
        top_controls.addWidget(dpi_group, 1, 0)
        top_controls.addWidget(keep_videos_group, 1, 1)
        top_controls.setColumnStretch(0, 1)
        top_controls.setColumnStretch(1, 1)
        self.top_controls = top_controls
        settings_layout.addLayout(top_controls)

        self.watermark_checkbox = QPushButton(self.text["watermark_toggle_on"])
        self.watermark_checkbox.setObjectName("toggleButton")
        self.watermark_checkbox.setCheckable(True)
        self.watermark_checkbox.setChecked(True)
        self.watermark_checkbox.setMinimumHeight(30)
        self.watermark_checkbox.setFixedWidth(62 if self.language == "zh" else 116)
        self.watermark_checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        self.watermark_checkbox.clicked.connect(self.on_settings_changed)
        switch_label = QLabel(self.text["watermark_switch"])
        switch_label.setObjectName("fieldLabel")
        switch_label.setContentsMargins(8, 0, 0, 0)

        kind_label = QLabel(self.text["watermark_type"])
        kind_label.setObjectName("fieldLabel")
        kind_label.setContentsMargins(8, 0, 0, 0)
        self.watermark_kind_group = QButtonGroup(self)
        self.watermark_kind_group.setExclusive(True)
        self.watermark_text_kind_button = QPushButton(self.text["watermark_type_text"])
        self.watermark_text_kind_button.setObjectName("segmentButton")
        self.watermark_text_kind_button.setProperty("side", "left")
        self.watermark_text_kind_button.setCheckable(True)
        self.watermark_text_kind_button.setChecked(True)
        self.watermark_text_kind_button.setFixedWidth(
            50 if self.language == "zh" else 58
        )
        self.watermark_image_kind_button = QPushButton(
            self.text["watermark_type_image"]
        )
        self.watermark_image_kind_button.setObjectName("segmentButton")
        self.watermark_image_kind_button.setProperty("side", "right")
        self.watermark_image_kind_button.setCheckable(True)
        self.watermark_image_kind_button.setFixedWidth(
            50 if self.language == "zh" else 58
        )
        self.watermark_kind_group.addButton(self.watermark_text_kind_button)
        self.watermark_kind_group.addButton(self.watermark_image_kind_button)
        self.watermark_text_kind_button.clicked.connect(self.on_settings_changed)
        self.watermark_image_kind_button.clicked.connect(self.on_settings_changed)

        color_label = QLabel(self.text["watermark_color"])
        color_label.setObjectName("fieldLabel")
        color_label.setContentsMargins(8, 0, 0, 0)
        self.color_select = CleanComboBox()
        self.color_select.setObjectName("colorSelect")
        self.color_select.setView(QListView())
        self.color_select.view().setObjectName("comboPopup")
        for name, value in PRESET_COLORS:
            display_name = (
                {
                    "Slate": "灰蓝",
                    "Gray": "中灰",
                    "Cloud": "浅灰",
                    "Blue": "蓝灰",
                    "Amber": "琥珀",
                    "Rose": "玫瑰灰",
                }.get(name, name)
                if self.language == "zh"
                else name
            )
            swatch = QPixmap(14, 14)
            swatch.fill(QColor(value))
            self.color_select.addItem(QIcon(swatch), display_name, value)
            self.color_select.setItemData(
                self.color_select.count() - 1,
                f"{display_name} · {value}",
                Qt.ItemDataRole.ToolTipRole,
            )
        self.color_select.setMinimumHeight(30)
        self.color_select.setMinimumWidth(120)
        self.color_select.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.color_select.currentIndexChanged.connect(self.on_settings_changed)

        opacity_label = QLabel(self.text["watermark_opacity"])
        opacity_label.setObjectName("fieldLabel")
        opacity_label.setContentsMargins(8, 0, 0, 0)
        self.opacity_input = QDoubleSpinBox()
        self.opacity_input.setRange(0.05, 0.60)
        self.opacity_input.setSingleStep(0.01)
        self.opacity_input.setDecimals(2)
        self.opacity_input.setValue(0.18)
        self.opacity_input.setMinimumHeight(30)
        self.opacity_input.setMinimumWidth(0)
        self.opacity_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.opacity_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.opacity_input.valueChanged.connect(self.on_settings_changed)

        text_label = QLabel(self.text["watermark_text"])
        text_label.setObjectName("fieldLabel")
        text_label.setContentsMargins(8, 0, 0, 0)
        self.watermark_text_input = QLineEdit()
        self.watermark_text_input.setPlaceholderText(
            self.text["watermark_text_placeholder"]
        )
        self.watermark_text_input.setText(
            DEFAULT_WATERMARK_TEXT if self.language == "zh" else "Company Confidential"
        )
        self.watermark_text_input.setMinimumHeight(30)
        self.watermark_text_input.setMinimumWidth(0)
        self.watermark_text_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.watermark_text_input.textChanged.connect(self.on_settings_changed)

        image_label = QLabel(self.text["watermark_image"])
        image_label.setObjectName("fieldLabel")
        image_label.setContentsMargins(8, 0, 0, 0)
        self.pick_watermark_image_button = WatermarkImageDropButton(
            self.text["watermark_pick_image"]
        )
        self.pick_watermark_image_button.setObjectName("secondaryButton")
        self.pick_watermark_image_button.setMinimumHeight(30)
        self.pick_watermark_image_button.setMinimumWidth(132)
        self.pick_watermark_image_button.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.pick_watermark_image_button.clicked.connect(self.choose_watermark_image)
        self.pick_watermark_image_button.fileDropped.connect(
            self.apply_watermark_image_path
        )
        self.watermark_image_name_label = ElidedLabel(self.text["watermark_no_image"])
        self.watermark_image_name_label.setObjectName("compactMeta")
        self.watermark_image_name_label.setMinimumHeight(30)

        preset_label = QLabel(self.text["watermark_preset"])
        preset_label.setObjectName("fieldLabel")
        preset_label.setContentsMargins(8, 0, 0, 0)
        self.preset_select = CleanComboBox()
        self.preset_select.setView(QListView())
        self.preset_select.view().setObjectName("comboPopup")
        for preset in WATERMARK_PRESETS:
            self.preset_select.addItem(preset[self.language], preset)
        self.preset_select.setMinimumHeight(30)
        self.preset_select.setFixedWidth(108 if self.language == "zh" else 126)

        image_width_label = QLabel(self.text["watermark_image_width"])
        image_width_label.setObjectName("fieldLabel")
        image_width_label.setContentsMargins(8, 0, 0, 0)
        self.image_width_input = QSpinBox()
        self.image_width_input.setRange(24, 800)
        self.image_width_input.setSingleStep(10)
        self.image_width_input.setValue(180)
        self.image_width_input.setMinimumHeight(30)
        self.image_width_input.setMinimumWidth(0)
        self.image_width_input.setButtonSymbols(
            QAbstractSpinBox.ButtonSymbols.NoButtons
        )
        self.image_width_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.image_width_input.valueChanged.connect(self.on_settings_changed)

        font_size_label = QLabel(self.text["watermark_font_size"])
        font_size_label.setObjectName("fieldLabel")
        font_size_label.setContentsMargins(8, 0, 0, 0)
        self.font_size_input = QSpinBox()
        self.font_size_input.setRange(12, 96)
        self.font_size_input.setValue(34)
        self.font_size_input.setMinimumHeight(30)
        self.font_size_input.setMinimumWidth(0)
        self.font_size_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.font_size_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.font_size_input.valueChanged.connect(self.on_settings_changed)

        spacing_label = QLabel(self.text["watermark_spacing"])
        spacing_label.setObjectName("fieldLabel")
        spacing_label.setContentsMargins(8, 0, 0, 0)
        self.spacing_input = QSpinBox()
        self.spacing_input.setRange(120, 720)
        self.spacing_input.setSingleStep(20)
        self.spacing_input.setValue(360)
        self.spacing_input.setMinimumHeight(30)
        self.spacing_input.setMinimumWidth(0)
        self.spacing_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.spacing_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.spacing_input.valueChanged.connect(self.on_settings_changed)

        angle_label = QLabel(self.text["watermark_angle"])
        angle_label.setObjectName("fieldLabel")
        angle_label.setContentsMargins(8, 0, 0, 0)
        self.angle_input = QSpinBox()
        self.angle_input.setRange(0, 359)
        self.angle_input.setValue(315)
        self.angle_input.setMinimumHeight(30)
        self.angle_input.setMinimumWidth(0)
        self.angle_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.angle_input.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.angle_input.valueChanged.connect(self.on_settings_changed)

        watermark_row = QGridLayout()
        watermark_row.setHorizontalSpacing(10)
        watermark_row.setVerticalSpacing(8)

        switch_group = QWidget()
        switch_group.setObjectName("controlGroup")
        switch_group.setProperty("plain", True)
        switch_group_layout = QVBoxLayout(switch_group)
        switch_group_layout.setContentsMargins(0, 0, 0, 0)
        switch_group_layout.setSpacing(4)
        switch_group_layout.addWidget(switch_label)
        switch_group_layout.addWidget(self.watermark_checkbox)
        switch_group.setFixedWidth(68 if self.language == "zh" else 116)
        watermark_row.addWidget(switch_group, 0, 0)

        kind_group = QWidget()
        kind_group.setObjectName("controlGroup")
        kind_group.setProperty("plain", True)
        kind_group_layout = QVBoxLayout(kind_group)
        kind_group_layout.setContentsMargins(0, 0, 0, 0)
        kind_group_layout.setSpacing(4)
        kind_buttons = QHBoxLayout()
        kind_buttons.setContentsMargins(0, 0, 0, 0)
        kind_buttons.setSpacing(0)
        kind_buttons.addWidget(self.watermark_text_kind_button)
        kind_buttons.addWidget(self.watermark_image_kind_button)
        kind_group_layout.addWidget(kind_label)
        kind_group_layout.addLayout(kind_buttons)
        kind_group.setFixedWidth(106 if self.language == "zh" else 122)
        watermark_row.addWidget(kind_group, 0, 1)

        text_group = QWidget()
        text_group.setObjectName("controlGroup")
        text_group.setProperty("plain", True)
        text_group_layout = QVBoxLayout(text_group)
        text_group_layout.setContentsMargins(0, 0, 0, 0)
        text_group_layout.setSpacing(4)
        text_group_layout.addWidget(text_label)
        text_group_layout.addWidget(self.watermark_text_input)
        self.watermark_text_group = text_group
        watermark_row.addWidget(text_group, 1, 0, 1, 3)

        image_group = QWidget()
        image_group.setObjectName("controlGroup")
        image_group.setProperty("plain", True)
        image_group_layout = QVBoxLayout(image_group)
        image_group_layout.setContentsMargins(0, 0, 0, 0)
        image_group_layout.setSpacing(4)
        image_group_layout.addWidget(image_label)
        image_group_layout.addWidget(self.pick_watermark_image_button)
        self.watermark_image_group = image_group
        watermark_row.addWidget(image_group, 1, 0, 1, 3)

        preset_group = QWidget()
        preset_group.setObjectName("controlGroup")
        preset_group.setProperty("plain", True)
        preset_group_layout = QVBoxLayout(preset_group)
        preset_group_layout.setContentsMargins(0, 0, 0, 0)
        preset_group_layout.setSpacing(4)
        preset_group_layout.addWidget(preset_label)
        preset_group_layout.addWidget(self.preset_select)
        self.watermark_preset_group = preset_group
        watermark_row.addWidget(preset_group, 0, 2)

        image_width_group = QWidget()
        image_width_group.setObjectName("controlGroup")
        image_width_group.setProperty("plain", True)
        image_width_group_layout = QVBoxLayout(image_width_group)
        image_width_group_layout.setContentsMargins(0, 0, 0, 0)
        image_width_group_layout.setSpacing(4)
        image_width_group_layout.addWidget(image_width_label)
        image_width_group_layout.addWidget(self.image_width_input)
        self.watermark_image_width_group = image_width_group
        watermark_row.addWidget(image_width_group, 0, 2)
        preset_group.setFixedWidth(108 if self.language == "zh" else 126)
        image_width_group.setFixedWidth(108 if self.language == "zh" else 126)
        watermark_row.setColumnStretch(0, 0)
        watermark_row.setColumnStretch(1, 0)
        watermark_row.setColumnStretch(2, 1)
        settings_layout.addLayout(watermark_row)

        tuning_grid = QGridLayout()
        tuning_grid.setHorizontalSpacing(6)
        tuning_grid.setVerticalSpacing(6)

        color_group = QWidget()
        color_group_layout = QVBoxLayout(color_group)
        color_group_layout.setContentsMargins(0, 0, 0, 0)
        color_group_layout.setSpacing(4)
        color_group_layout.addWidget(color_label)
        color_group_layout.addWidget(self.color_select)
        color_group.setMinimumWidth(88)

        opacity_group = QWidget()
        opacity_group_layout = QVBoxLayout(opacity_group)
        opacity_group_layout.setContentsMargins(0, 0, 0, 0)
        opacity_group_layout.setSpacing(4)
        opacity_group_layout.addWidget(opacity_label)
        opacity_group_layout.addWidget(self.opacity_input)

        font_group = QWidget()
        font_group_layout = QVBoxLayout(font_group)
        font_group_layout.setContentsMargins(0, 0, 0, 0)
        font_group_layout.setSpacing(4)
        font_group_layout.addWidget(font_size_label)
        font_group_layout.addWidget(self.font_size_input)

        spacing_group = QWidget()
        spacing_group_layout = QVBoxLayout(spacing_group)
        spacing_group_layout.setContentsMargins(0, 0, 0, 0)
        spacing_group_layout.setSpacing(4)
        spacing_group_layout.addWidget(spacing_label)
        spacing_group_layout.addWidget(self.spacing_input)

        angle_group = QWidget()
        angle_group_layout = QVBoxLayout(angle_group)
        angle_group_layout.setContentsMargins(0, 0, 0, 0)
        angle_group_layout.setSpacing(4)
        angle_group_layout.addWidget(angle_label)
        angle_group_layout.addWidget(self.angle_input)

        tuning_grid.addWidget(color_group, 0, 0, 1, 2)
        tuning_grid.addWidget(opacity_group, 0, 2)
        tuning_grid.addWidget(font_group, 0, 3)
        tuning_grid.addWidget(spacing_group, 1, 0, 1, 2)
        tuning_grid.addWidget(angle_group, 1, 2, 1, 2)
        for idx in range(4):
            tuning_grid.setColumnStretch(idx, 1)
        settings_layout.addLayout(tuning_grid)
        self.preset_select.currentIndexChanged.connect(
            self.apply_watermark_preset_from_select
        )

        side_label_width = 72 if self.language == "zh" else 118
        preset_hint_row = QHBoxLayout()
        preset_hint_row.setSpacing(8)
        preset_hint_title = QLabel(self.text["watermark_preset_hint"])
        preset_hint_title.setObjectName("fieldLabel")
        preset_hint_title.setFixedWidth(side_label_width)
        preset_hint_row.addWidget(preset_hint_title, 0, Qt.AlignVCenter)
        preset_hint_content = QWidget()
        preset_hint_content_layout = QVBoxLayout(preset_hint_content)
        preset_hint_content_layout.setContentsMargins(0, 0, 0, 0)
        preset_hint_content_layout.setSpacing(6)
        self.preset_hint_label = QLabel()
        self.preset_hint_label.setObjectName("outputHint")
        self.preset_hint_label.setWordWrap(True)
        self.preset_hint_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.preset_hint_label.setMinimumHeight(48)
        self.preset_hint_label.setMaximumHeight(58)
        self.preset_hint_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        preset_hint_content_layout.addWidget(self.preset_hint_label)

        self.advanced_button = QPushButton(self.text["advanced_settings"])
        self.advanced_button.setObjectName("advancedDisclosure")
        self.advanced_button.setCheckable(True)
        self.advanced_button.setAccessibleName(self.text["advanced_settings"])
        self.advanced_button.setMinimumHeight(30)
        self.advanced_button.clicked.connect(self.toggle_advanced_settings)
        preset_hint_row.addWidget(preset_hint_content, 1)
        settings_layout.addLayout(preset_hint_row)
        settings_layout.addWidget(self.advanced_button)

        self.advanced_panel = QWidget()
        self.advanced_panel.setObjectName("advancedPanel")
        advanced_layout = QVBoxLayout(self.advanced_panel)
        advanced_layout.setContentsMargins(0, 0, 0, 0)
        advanced_layout.setSpacing(8)

        font_check_row = QHBoxLayout()
        font_check_row.setSpacing(8)
        font_check_header = QLabel(self.text["font_check_title"])
        font_check_header.setObjectName("fieldLabel")
        font_check_header.setFixedWidth(side_label_width)
        font_check_row.addWidget(font_check_header, 0, Qt.AlignVCenter)
        self.font_check_label = ElidedLabel(self.text["font_check_waiting"])
        self.font_check_label.setObjectName("outputHint")
        self.font_check_label.setMinimumHeight(32)
        self.font_check_label.setMaximumHeight(32)
        self.font_check_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed
        )
        font_check_row.addWidget(self.font_check_label, 1)
        self.font_fix_button = QPushButton(self.text["font_fix_off"])
        self.font_fix_button.setObjectName("toggleButton")
        self.font_fix_button.setCheckable(True)
        self.font_fix_button.setMinimumHeight(32)
        self.font_fix_button.setFixedWidth(86 if self.language == "zh" else 106)
        self.font_fix_button.clicked.connect(self.on_settings_changed)
        font_check_row.addWidget(self.font_fix_button, 0, Qt.AlignVCenter)
        advanced_layout.addLayout(font_check_row)

        dependency_row = QHBoxLayout()
        dependency_row.setSpacing(8)
        dependency_header = QLabel(self.text["dependency_title"])
        dependency_header.setObjectName("fieldLabel")
        dependency_header.setFixedWidth(side_label_width)
        dependency_row.addWidget(dependency_header, 0, Qt.AlignVCenter)
        self.dependency_hint_label = ElidedLabel(self.text["dependency_ok"])
        self.dependency_hint_label.setObjectName("outputHint")
        self.dependency_hint_label.setMinimumHeight(32)
        self.dependency_hint_label.setMaximumHeight(32)
        self.dependency_hint_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed
        )
        dependency_row.addWidget(self.dependency_hint_label, 1)
        self.install_libreoffice_button = QPushButton(
            self.text["dependency_install_libreoffice"]
        )
        self.install_libreoffice_button.setObjectName("secondaryButton")
        self.install_libreoffice_button.setMinimumHeight(30)
        self.install_libreoffice_button.clicked.connect(self.open_dependency_action)
        dependency_row.addWidget(self.install_libreoffice_button, 0, Qt.AlignVCenter)
        advanced_layout.addLayout(dependency_row)
        self.advanced_panel.hide()
        settings_layout.addWidget(self.advanced_panel)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.event_log = QTextEdit()
        self.event_log.setReadOnly(True)
        self.event_log.setObjectName("eventLog")
        self.event_log.setPlainText(self.text["log_waiting"])
        self.event_log.setMinimumHeight(132)

        self.output_path_label = ElidedLabel(self.text["output_waiting"])
        self.output_path_label.setObjectName("outputHint")
        self.output_path_label.setToolTip(self.text["output_waiting"])
        self.output_path_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.output_path_label.setFixedHeight(30)
        self.output_path_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed
        )
        self.current_file_label = QLabel(self.text["status_ready"])
        self.current_file_label.setObjectName("currentFile")
        self.current_file_label.setWordWrap(True)
        self.current_file_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.current_file_label.setFixedHeight(38)
        self.current_file_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.current_file_label.hide()
        right_column.addWidget(right_card, 0)
        right_column.addStretch(1)

        preview_card = QFrame()
        preview_card.setObjectName("previewCard")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(12, 12, 12, 12)
        preview_layout.setSpacing(8)

        preview_header = QGridLayout()
        preview_header.setHorizontalSpacing(8)
        preview_header.setVerticalSpacing(6)
        preview_context = QWidget()
        preview_context_layout = QHBoxLayout(preview_context)
        preview_context_layout.setContentsMargins(0, 0, 0, 0)
        preview_context_layout.setSpacing(8)
        preview_title = QLabel(self.text["preview_title"])
        preview_title.setObjectName("sectionTitle")
        preview_context_layout.addWidget(preview_title)
        preview_theme_label = QLabel(self.text["preview_background"])
        preview_theme_label.setObjectName("fieldLabel")
        preview_context_layout.addSpacing(8)
        preview_context_layout.addWidget(preview_theme_label)
        self.preview_background_select = CleanComboBox()
        self.preview_background_select.setView(QListView())
        self.preview_background_select.view().setObjectName("comboPopup")
        for theme in PREVIEW_BACKGROUND_THEMES:
            self.preview_background_select.addItem(theme[self.language], theme)
        self.preview_background_select.setMinimumHeight(30)
        self.preview_background_select.setFixedWidth(114)
        self.preview_background_select.currentIndexChanged.connect(
            self.refresh_sample_preview_if_needed
        )
        preview_context_layout.addWidget(self.preview_background_select)
        preview_context_layout.addStretch(1)
        self.preview_page_label = QLabel(self.text["preview_waiting"])
        self.preview_page_label.setObjectName("fieldLabel")
        self.preview_page_label.setMinimumWidth(0)
        self.preview_page_label.setVisible(False)
        preview_actions = QWidget()
        preview_actions_layout = QHBoxLayout(preview_actions)
        preview_actions_layout.setContentsMargins(0, 0, 0, 0)
        preview_actions_layout.setSpacing(8)
        self.original_preview_button = QPushButton(self.text["preview_original"])
        self.original_preview_button.setObjectName("previewToolbarButton")
        self.original_preview_button.setCheckable(True)
        self.original_preview_button.setFixedWidth(46 if self.language == "zh" else 80)
        self.original_preview_button.clicked.connect(
            lambda: self.set_preview_mode("original")
        )
        self.rendered_preview_button = QPushButton(self.text["preview_output"])
        self.rendered_preview_button.setObjectName("previewToolbarButton")
        self.rendered_preview_button.setCheckable(True)
        self.rendered_preview_button.setFixedWidth(46 if self.language == "zh" else 80)
        self.rendered_preview_button.clicked.connect(
            lambda: self.set_preview_mode("preview")
        )
        self.preview_prev_button = QPushButton(self.text["preview_prev"])
        self.preview_prev_button.setObjectName("previewToolbarButton")
        self.preview_prev_button.setText(
            "上一组" if self.language == "zh" else "Previous"
        )
        self.preview_prev_button.setToolTip(self.text["preview_prev"])
        self.preview_prev_button.setAccessibleName(self.text["preview_prev"])
        self.preview_prev_button.setFixedWidth(62 if self.language == "zh" else 76)
        self.preview_prev_button.clicked.connect(lambda: self.change_preview_page(-1))
        self.preview_next_button = QPushButton(self.text["preview_next"])
        self.preview_next_button.setObjectName("previewToolbarButton")
        self.preview_next_button.setText("下一组" if self.language == "zh" else "Next")
        self.preview_next_button.setToolTip(self.text["preview_next"])
        self.preview_next_button.setAccessibleName(self.text["preview_next"])
        self.preview_next_button.setFixedWidth(62 if self.language == "zh" else 64)
        self.preview_next_button.clicked.connect(lambda: self.change_preview_page(1))
        self.preview_refresh_button = QPushButton(self.text["preview_refresh"])
        self.preview_refresh_button.setObjectName("previewToolbarButton")
        self.preview_refresh_button.setFixedWidth(50 if self.language == "zh" else 64)
        self.preview_refresh_button.clicked.connect(self.force_preview_refresh)
        self.preview_note_label = QLabel(self.text["preview_note"])
        self.preview_note_label.setObjectName("subtleNote")
        self.preview_note_label.setToolTip(self.text["preview_note"])
        self.preview_note_label.setVisible(False)
        preview_actions_layout.addStretch(1)
        preview_actions_layout.addWidget(self.preview_prev_button)
        preview_actions_layout.addWidget(self.preview_page_label)
        preview_actions_layout.addWidget(self.preview_next_button)
        preview_actions_layout.addWidget(self.preview_refresh_button)
        preview_actions_layout.addWidget(self.original_preview_button)
        preview_actions_layout.addWidget(self.rendered_preview_button)
        preview_header.addWidget(preview_context, 0, 0)
        preview_header.addWidget(preview_actions, 0, 1)
        self.preview_header_layout = preview_header
        self.preview_context_widget = preview_context
        self.preview_actions_widget = preview_actions
        preview_layout.addLayout(preview_header)

        self.preview_image_label = PreviewCanvas(self.text["preview_waiting"])
        self.preview_image_label.setObjectName("previewArea")
        self.preview_image_label.set_watermark_font_family(self.watermark_font_family)
        self.preview_image_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.preview_image_label.setMinimumHeight(100)
        self.preview_image_label_secondary = PreviewCanvas()
        self.preview_image_label_secondary.setObjectName("previewArea")
        self.preview_image_label_secondary.set_watermark_font_family(
            self.watermark_font_family
        )
        self.preview_image_label_secondary.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.preview_image_label_secondary.setMinimumHeight(100)
        self.preview_image_label_secondary.clear_preview(
            self.text["preview_waiting"],
            self.sample_watermark_options(),
            self.sample_preview_theme(),
        )

        self.preview_pages_widget = QWidget()
        self.preview_pages_widget.setObjectName("previewPages")
        self.preview_pages_layout = QVBoxLayout(self.preview_pages_widget)
        self.preview_pages_layout.setContentsMargins(0, 0, 0, 0)
        self.preview_pages_layout.setSpacing(12)
        self.preview_pages_layout.addWidget(
            self.preview_image_label, 0, Qt.AlignmentFlag.AlignHCenter
        )
        self.preview_pages_layout.addWidget(
            self.preview_image_label_secondary, 0, Qt.AlignmentFlag.AlignHCenter
        )
        self.preview_pages_layout.addStretch(1)

        self.preview_scroll_area = QScrollArea()
        self.preview_scroll_area.setObjectName("previewScrollArea")
        self.preview_scroll_area.setWidgetResizable(True)
        self.preview_scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.preview_scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.preview_scroll_area.setWidget(self.preview_pages_widget)
        self.preview_scroll_viewport = self.preview_scroll_area.viewport()
        self.preview_scroll_viewport.installEventFilter(self)
        preview_layout.addWidget(self.preview_scroll_area, 1)

        self.preview_thumbnail_list = QListWidget()
        self.preview_thumbnail_list.setObjectName("previewThumbnails")
        self.preview_thumbnail_list.setViewMode(QListView.ViewMode.IconMode)
        self.preview_thumbnail_list.setFlow(QListView.Flow.LeftToRight)
        self.preview_thumbnail_list.setMovement(QListView.Movement.Static)
        self.preview_thumbnail_list.setResizeMode(QListView.ResizeMode.Adjust)
        self.preview_thumbnail_list.setWrapping(False)
        self.preview_thumbnail_list.setSelectionMode(
            QListWidget.SelectionMode.SingleSelection
        )
        self.preview_thumbnail_list.setIconSize(QSize(96, 54))
        self.preview_thumbnail_list.setFixedHeight(82)
        self.preview_thumbnail_list.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.preview_thumbnail_list.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.preview_thumbnail_list.currentRowChanged.connect(self.select_preview_page)
        placeholder = QListWidgetItem(self.text["preview_thumbnails_empty"])
        placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
        self.preview_thumbnail_list.addItem(placeholder)
        preview_layout.addWidget(self.preview_thumbnail_list)
        center_column.addWidget(preview_card, 1)

        left_pane = QWidget()
        left_pane.setObjectName("leftPane")
        left_pane.setMinimumWidth(0)
        left_pane.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding
        )
        left_pane.setLayout(left_column)

        center_pane = QWidget()
        center_pane.setObjectName("centerPane")
        center_pane.setMinimumWidth(520)
        center_pane.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        center_pane.setLayout(center_column)

        right_pane = QWidget()
        right_pane.setObjectName("rightPane")
        right_pane.setMinimumWidth(410)
        right_pane.setMaximumWidth(460)
        right_pane.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding
        )
        right_pane.setLayout(right_column)
        left_pane.setMinimumWidth(320)
        left_pane.setMaximumWidth(320)
        left_pane.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding
        )
        self.left_pane = left_pane
        self.center_pane = center_pane
        self.right_pane = right_pane

        body_row.addWidget(left_pane, 0)
        body_row.addWidget(center_pane, 1)
        body_row.addWidget(right_pane, 0)
        body_row.setStretch(0, 0)
        body_row.setStretch(1, 1)
        body_row.setStretch(2, 0)
        root.addLayout(body_row, 1)

        self.log_shelf = QPushButton(
            "状态与日志 · 等待开始"
            if self.language == "zh"
            else "Status and logs · Ready"
        )
        self.log_shelf.setObjectName("logShelf")
        self.log_shelf.setAccessibleName(self.text["details_title"])
        self.log_shelf.setToolTip(
            "悬停 1 秒或点击查看详细日志"
            if self.language == "zh"
            else "Hover for 1 second or click to view detailed logs"
        )
        self.log_shelf.setFixedHeight(26)
        self.log_shelf.clicked.connect(self.toggle_log_drawer)
        self.log_shelf.installEventFilter(self)
        root.addWidget(self.log_shelf)

        self.log_drawer = QFrame(central)
        self.log_drawer.setObjectName("logDrawer")
        self.log_drawer.installEventFilter(self)
        self.event_log.installEventFilter(self)
        log_drawer_layout = QVBoxLayout(self.log_drawer)
        log_drawer_layout.setContentsMargins(12, 10, 12, 12)
        log_drawer_layout.setSpacing(6)
        log_drawer_header = QHBoxLayout()
        log_drawer_title = QLabel(self.text["details_title"])
        log_drawer_title.setObjectName("sectionTitle")
        log_drawer_header.addWidget(log_drawer_title)
        log_drawer_header.addStretch(1)
        close_log_drawer = QPushButton(self.text["close_button"])
        close_log_drawer.setObjectName("disclosureButton")
        close_log_drawer.setFixedHeight(28)
        close_log_drawer.clicked.connect(self.hide_log_drawer)
        log_drawer_header.addWidget(close_log_drawer)
        log_drawer_layout.addLayout(log_drawer_header)
        log_drawer_layout.addWidget(self.progress_bar)
        log_drawer_layout.addWidget(self.output_path_label)
        log_drawer_layout.addWidget(self.current_file_label)
        log_drawer_layout.addWidget(self.event_log, 1)
        self.log_drawer.hide()
        self.log_drawer_timer = QTimer(self)
        self.log_drawer_timer.setSingleShot(True)
        self.log_drawer_timer.timeout.connect(self.hide_log_drawer_if_idle)
        self.log_hover_timer = QTimer(self)
        self.log_hover_timer.setSingleShot(True)
        self.log_hover_timer.setInterval(1000)
        self.log_hover_timer.timeout.connect(self.show_log_drawer_after_hover)

        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(420)
        self.preview_timer.timeout.connect(self.start_preview_for_selection)

        self.set_preview_mode("preview")
        self.update_image_quality_control()
        self.update_watermark_controls()
        self.update_preset_hint()
        self.update_font_check()
        self.refresh_file_list()
        self.update_dependency_hint()
        self.apply_styles()
        for control in (
            self.pick_button,
            self.remove_button,
            self.output_format_select,
            self.output_mode_select,
            self.dpi_select,
            self.keep_videos_checkbox,
            self.run_button,
            self.watermark_checkbox,
            self.watermark_text_kind_button,
            self.watermark_image_kind_button,
            self.color_select,
            self.opacity_input,
            self.watermark_text_input,
            self.pick_watermark_image_button,
            self.preset_select,
            self.image_width_input,
            self.font_size_input,
            self.spacing_input,
            self.angle_input,
            self.advanced_button,
            self.font_fix_button,
            self.install_libreoffice_button,
            self.preview_background_select,
            self.original_preview_button,
            self.rendered_preview_button,
            self.preview_prev_button,
            self.preview_next_button,
            self.preview_refresh_button,
        ):
            control.setMinimumHeight(
                max(control.minimumHeight(), control.sizeHint().height())
            )
        install_control_help(self)
        for control in (
            self.output_format_select,
            self.output_mode_select,
            self.dpi_select,
            self.color_select,
        ):
            control.setMinimumWidth(
                max(control.minimumWidth(), control.sizeHint().width())
            )
        self.clear_preview(self.text["preview_waiting"])
        QTimer.singleShot(0, self.update_preview_geometry)

    def apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget#central { background: #0b1017; }
            QLabel {
                color: #dbe4f0;
                font-size: 13px;
            }
            QFrame#headerCard, QFrame#sideCard, QFrame#detailsCard, QFrame#queueCard, QFrame#leftCard, QFrame#previewCard, QFrame#rightCard {
                background: #121a24;
                border: 1px solid #273244;
                border-radius: 12px;
            }
            QFrame#softDivider {
                color: #273244;
                background: #273244;
                max-height: 1px;
                border: 0;
            }
            QLabel#eyebrow {
                color: #f97316;
                font-size: 11px;
                font-weight: 700;
                letter-spacing: 0.08em;
            }
            QLabel#title { font-size: 18px; font-weight: 700; color: #f8fafc; }
            QLabel#subtitle { color: #94a3b8; font-size: 12px; }
            QLabel#sectionTitle {
                color: #f8fafc;
                font-size: 15px;
                font-weight: 600;
            }
            QLabel#fieldLabel {
                color: #94a3b8;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#estimatePill {
                color: #f8fafc;
                background: #18212d;
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 4px 12px;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#currentFile {
                color: #cbd5e1;
                background: #0f1720;
                border: 1px solid #273244;
                border-radius: 10px;
                padding: 0 10px;
                font-size: 12px;
            }
            QLabel#compactMeta {
                color: #cbd5e1;
                background: #0f1720;
                border: 1px solid #273244;
                border-radius: 10px;
                padding: 6px 8px;
                font-size: 12px;
            }
            QWidget#controlGroup {
                background: rgba(15, 23, 32, 0.62);
                border: 1px solid rgba(59, 74, 95, 0.88);
                border-radius: 8px;
                padding: 0 10px;
            }
            QWidget#controlGroup[plain="true"] {
                background: transparent;
                border: 0;
                border-radius: 0;
                padding: 0;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 0 26px 0 12px;
                selection-background-color: #f97316;
                selection-color: #ffffff;
                font-size: 13px;
                font-weight: 600;
                min-height: 28px;
            }
            QLineEdit { padding: 0 12px; }
            QSpinBox, QDoubleSpinBox { padding-right: 10px; }
            QComboBox#colorSelect {
                padding: 0 26px 0 8px;
                font-size: 13px;
            }
            QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled {
                background: #0b111a;
                color: #64748b;
                border: 1px solid #273244;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
            QPushButton:focus, QListWidget:focus, QTextEdit:focus {
                border: 1px solid #fb923c;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 26px;
                border: 0;
                background: transparent;
            }
            QComboBox::down-arrow { image: none; width: 0; height: 0; }
            QListView#comboPopup {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 4px;
                outline: 0;
                selection-background-color: #243244;
                selection-color: #f8fafc;
                font-size: 13px;
                font-weight: 600;
            }
            QListView#comboPopup::item {
                min-height: 26px;
                padding: 5px 10px;
                border-radius: 6px;
            }
            QListView#comboPopup::item:selected { background: #243244; }
            QPushButton {
                background: #18212d;
                color: #ffffff;
                border: 1px solid #334155;
                border-radius: 12px;
                padding: 6px 12px;
                min-width: 72px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover { background: #202b39; }
            QPushButton:disabled {
                background: #334155;
                color: #94a3b8;
            }
            QPushButton#primaryButton {
                background: #f97316;
                border: 1px solid #fb923c;
                color: #ffffff;
                min-width: 102px;
                max-width: 102px;
            }
            QPushButton#primaryButton:disabled {
                background: #334155;
                border: 1px solid #475569;
                color: #94a3b8;
            }
            QPushButton#primaryButton:hover { background: #ea580c; }
            QPushButton#primaryButton[mode="stop"] {
                background: #dc2626;
                border: 1px solid #ef4444;
            }
            QPushButton#primaryButton[mode="stop"]:hover { background: #b91c1c; }
            QPushButton#secondaryButton {
                min-width: 0;
                padding: 8px 12px;
            }
            QPushButton#secondaryButton:checked {
                background: #243244;
                border: 1px solid #475569;
            }
            QPushButton#previewToolbarButton {
                min-width: 0;
                padding: 4px 6px;
                font-size: 13px;
            }
            QPushButton#previewToolbarButton:checked {
                background: #243244;
                border: 1px solid #475569;
            }
            QPushButton#disclosureButton {
                background: #18212d;
                color: #94a3b8;
                border: 1px solid #334155;
                border-radius: 8px;
                min-width: 0;
                padding: 0 10px;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton#disclosureButton:hover {
                background: #18212d;
                color: #f8fafc;
            }
            QPushButton#advancedDisclosure {
                background: transparent;
                color: #cbd5e1;
                border: 0;
                border-top: 1px solid #273244;
                border-radius: 0;
                min-width: 0;
                padding: 0 2px;
                text-align: left;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton#advancedDisclosure:hover, QPushButton#advancedDisclosure:checked {
                color: #f8fafc;
                background: #18212d;
            }
            QPushButton#logShelf {
                background: #0f1720;
                color: #94a3b8;
                border: 1px solid #273244;
                border-radius: 8px;
                min-width: 0;
                padding: 0 10px;
                text-align: left;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton#logShelf:hover { color: #f8fafc; background: #18212d; }
            QFrame#logDrawer {
                background: #111827;
                border: 1px solid #334155;
                border-radius: 12px;
            }
            QPushButton#toggleButton {
                background: #18212d;
                color: #94a3b8;
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 0 8px;
                min-width: 0;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#toggleButton[active="true"] {
                background: #f97316;
                color: #ffffff;
                border: 1px solid #fb923c;
            }
            QPushButton#toggleButton[active="false"] {
                background: #1e293b;
                color: #dbe4f0;
                border: 1px solid #475569;
            }
            QPushButton#toggleButton[active="false"]:hover {
                background: #243244;
                border: 1px solid #64748b;
            }
            QPushButton#toggleButton:disabled {
                background: #334155;
                color: #94a3b8;
                border: 1px solid #475569;
            }
            QPushButton#segmentButton {
                background: #0f1720;
                color: #94a3b8;
                border: 1px solid #334155;
                border-radius: 0;
                padding: 0 6px;
                min-width: 42px;
                min-height: 28px;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#segmentButton[side="left"] {
                border-top-left-radius: 10px;
                border-bottom-left-radius: 10px;
            }
            QPushButton#segmentButton[side="right"] {
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }
            QPushButton#segmentButton:checked {
                background: #243244;
                color: #f8fafc;
                border: 1px solid #475569;
            }
            QPushButton#helpIconButton {
                color: #cbd5e1;
                background: #0f1720;
                border: 1px solid #334155;
                border-radius: 15px;
                padding: 0;
                min-width: 30px;
                max-width: 30px;
                min-height: 30px;
                max-height: 30px;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#helpIconButton:hover {
                background: #18212d;
                color: #f8fafc;
            }
            QCheckBox {
                color: #f8fafc;
                font-size: 13px;
                font-weight: 600;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 5px;
                border: 1px solid #475569;
                background: #0f1720;
            }
            QCheckBox::indicator:checked {
                background: #f97316;
                border: 1px solid #fb923c;
            }
            QProgressBar {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #273244;
                border-radius: 8px;
                text-align: center;
                min-height: 14px;
            }
            QProgressBar::chunk {
                background: #f97316;
                border-radius: 6px;
            }
            QLabel#outputHint {
                color: #94a3b8;
                background: #0f1720;
                border: 1px solid #273244;
                border-radius: 10px;
                padding: 8px;
                font-size: 13px;
            }
            QTextEdit#eventLog {
                background: #0f1720;
                color: #94a3b8;
                border: 1px solid #273244;
                border-radius: 10px;
                padding: 6px;
                font-size: 13px;
            }
            QListWidget#fileList {
                background: #0f1720;
                color: #cbd5e1;
                border: 1px solid #273244;
                border-radius: 10px;
                padding: 6px;
                font-size: 13px;
                outline: 0;
            }
            QListWidget#fileList::item {
                padding: 6px;
                border-radius: 6px;
            }
            QListWidget#fileList::item:selected {
                background: #1e293b;
                color: #f8fafc;
            }
            QListWidget#previewThumbnails {
                background: #0f1720;
                color: #dbe4f0;
                border: 1px solid #273244;
                border-radius: 8px;
                padding: 4px;
                outline: 0;
            }
            QListWidget#previewThumbnails::item {
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 3px;
                margin-right: 6px;
            }
            QListWidget#previewThumbnails::item:selected {
                background: #18212d;
                color: #f8fafc;
                border: 2px solid #f97316;
            }
            QWidget#fileItem {
                background: #09131f;
                border: 1px solid #1d3043;
                border-radius: 8px;
            }
            QWidget#fileItem[included="true"] {
                background: #14273a;
                border-color: #36516a;
            }
            QWidget#fileItem[included="false"] QLabel#fileName,
            QWidget#fileItem[included="false"] QLabel#fileMeta {
                color: #64748b;
            }
            QWidget#fileItem[selected="true"] {
                border-color: #f97316;
            }
            QWidget#fileItem[state="running"] {
                background: #17202c;
                border-radius: 8px;
            }
            QLabel#fileName {
                color: #cbd5e1;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#fileMeta {
                color: #94a3b8;
                font-size: 12px;
            }
            QLabel#fileState {
                color: #94a3b8;
                font-size: 12px;
                font-weight: 700;
                border: 1px solid #475569;
                border-radius: 6px;
                padding: 2px 5px;
            }
            QLabel#fileState[state="queued"], QLabel#fileState[state="pending"] {
                color: #94a3b8;
            }
            QLabel#fileState[state="running"], QLabel#fileState[state="stopped"] {
                color: #fb923c;
                border-color: #c2410c;
            }
            QLabel#fileState[state="done"] {
                color: #4ade80;
                border-color: #15803d;
            }
            QLabel#fileState[state="failed"] {
                color: #f87171;
                border-color: #b91c1c;
            }
            QLabel#selectionSummary {
                color: #94a3b8;
                font-size: 12px;
                padding: 2px 4px;
            }
            QPushButton#fileTypeToggle {
                background: transparent;
                border: 0;
                border-radius: 6px;
                padding: 3px;
            }
            QPushButton#fileTypeToggle:hover {
                background: #20364b;
            }
            QLabel#previewArea {
                background: #0f1720;
                border: 1px solid #273244;
                border-radius: 8px;
                color: #64748b;
                padding: 0;
            }
            QScrollArea#previewScrollArea,
            QScrollArea#previewScrollArea > QWidget > QWidget,
            QWidget#previewPages {
                background: #0f1720;
                border: 0;
            }
            QScrollBar:vertical, QScrollBar:horizontal {
                background: #0f1720;
                border: 0;
                margin: 4px 2px 4px 2px;
                border-radius: 5px;
            }
            QScrollBar:vertical { width: 10px; }
            QScrollBar:horizontal { height: 10px; }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: #334155;
                border-radius: 5px;
                min-height: 32px;
                min-width: 32px;
            }
            QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
                background: #475569;
            }
            QScrollBar::add-line, QScrollBar::sub-line,
            QScrollBar::add-page, QScrollBar::sub-page {
                background: transparent;
                border: 0;
                width: 0;
                height: 0;
            }
            QLabel, QPushButton, QLineEdit, QComboBox, QSpinBox,
            QDoubleSpinBox, QCheckBox, QListWidget, QTextEdit {
                font-size: 12px;
            }
            QPushButton, QPushButton#primaryButton, QPushButton#secondaryButton,
            QPushButton#previewToolbarButton, QPushButton#disclosureButton,
            QPushButton#advancedDisclosure, QPushButton#logShelf,
            QPushButton#toggleButton, QPushButton#segmentButton,
            QPushButton#helpIconButton, QComboBox#colorSelect {
                font-size: 12px;
            }
            QLabel#fieldLabel { font-size: 12px; }
            QLabel#title { font-size: 16px; }
            QLabel#sectionTitle { font-size: 13px; }
            QLabel#eyebrow, QLabel#subtitle, QLabel#currentFile,
            QLabel#compactMeta, QLabel#fileMeta, QLabel#fileState,
            QLabel#selectionSummary, QLabel#subtleNote,
            QPushButton#logShelf, QPushButton#disclosureButton {
                font-size: 11px;
            }
            """
            + SHARED_MAIN_QSS
        )

    def toggle_advanced_settings(self, expanded: bool) -> None:
        self.advanced_panel.setVisible(expanded)
        self.advanced_button.setText(
            f"{self.text['advanced_settings']} · "
            f"{self.text['advanced_collapse'] if expanded else self.text['advanced_expand']}"
        )

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        if (
            watched is getattr(self, "content_widget", None)
            and event.type() == QEvent.Type.Resize
        ):
            self.update_responsive_layout(event.size().width())
        if (
            watched is getattr(self, "preview_scroll_viewport", None)
            and event.type() == QEvent.Type.Resize
        ):
            self.update_preview_geometry()
        if watched in {
            getattr(self, "log_shelf", None),
            getattr(self, "log_drawer", None),
            getattr(self, "event_log", None),
        }:
            if event.type() == QEvent.Type.Enter:
                self.log_drawer_timer.stop()
                if watched is self.log_shelf and self.log_drawer.isHidden():
                    self.log_hover_timer.start()
            elif event.type() == QEvent.Type.Leave:
                self.log_hover_timer.stop()
                if not self.is_running:
                    self.log_drawer_timer.start(900)
        return super().eventFilter(watched, event)

    def update_responsive_layout(self, width: int) -> None:
        if not all(
            hasattr(self, name) for name in ("left_pane", "center_pane", "right_pane")
        ):
            return
        compact = width < 1180
        left_width = 260 if compact else 320
        right_width = 310 if compact else 410
        if self.language == "en":
            self.keep_videos_checkbox.setFixedWidth(58 if compact else 108)
            self.preset_select.setFixedWidth(108 if compact else 126)
            self.keep_videos_label_widget.setText(
                "Video" if compact else "Embedded video"
            )
            self.keep_videos_label_widget.setFixedWidth(40 if compact else 96)
            self.keep_videos_group.setMinimumWidth(128 if compact else 218)
        self.left_pane.setFixedWidth(left_width)
        self.right_pane.setMinimumWidth(right_width)
        self.right_pane.setMaximumWidth(330 if compact else 460)
        self.center_pane.setMinimumWidth(0 if compact else 520)
        layout = self.preview_header_layout
        layout.removeWidget(self.preview_context_widget)
        layout.removeWidget(self.preview_actions_widget)
        layout.addWidget(self.preview_context_widget, 0, 0)
        layout.addWidget(
            self.preview_actions_widget,
            1 if compact else 0,
            0 if compact else 1,
        )
        self.update_preview_geometry()

    def position_log_drawer(self) -> None:
        if not hasattr(self, "log_drawer") or self.log_drawer.isHidden():
            return
        shelf_origin = self.log_shelf.mapTo(self.content_widget, QPoint(0, 0))
        height_ratio = 0.35 if self.content_widget.width() < 1180 else 0.45
        height = min(340, max(180, int(self.content_widget.height() * height_ratio)))
        height = min(height, max(160, shelf_origin.y() - 8))
        self.log_drawer.setGeometry(
            shelf_origin.x(),
            max(4, shelf_origin.y() - height - 5),
            self.log_shelf.width(),
            height,
        )
        self.log_drawer.raise_()

    def show_log_drawer(self, *, auto_hide: bool = True) -> None:
        self.log_hover_timer.stop()
        self.log_drawer.show()
        self.position_log_drawer()
        QTimer.singleShot(0, self.position_log_drawer)
        if auto_hide:
            self.log_drawer_timer.start(2500)
        else:
            self.log_drawer_timer.stop()

    def show_log_drawer_after_hover(self) -> None:
        if self.log_shelf.underMouse():
            self.show_log_drawer(auto_hide=False)

    def hide_log_drawer_if_idle(self) -> None:
        if self.is_running:
            return
        if any(
            widget.underMouse()
            for widget in (self.log_shelf, self.log_drawer, self.event_log)
        ):
            return
        self.hide_log_drawer()

    def hide_log_drawer(self) -> None:
        self.log_hover_timer.stop()
        self.log_drawer_timer.stop()
        self.log_drawer.hide()

    def toggle_log_drawer(self) -> None:
        if self.log_drawer.isHidden():
            self.show_log_drawer(auto_hide=False)
        else:
            self.hide_log_drawer()

    def refresh_log_shelf(self) -> None:
        status = self.current_file_label.text().strip() or self.text["status_ready"]
        prefix = "状态与日志" if self.language == "zh" else "Status and logs"
        self.log_shelf.setText(f"{prefix} · {self.progress_bar.value()}% · {status}")

    def show_dialog(self, title: str, message: str) -> None:
        dialog_parent = QApplication.activeWindow() or self
        dialog = StyledDialog(dialog_parent, title, message)
        ok_button = dialog.findChild(QPushButton)
        if ok_button is not None:
            ok_button.setText(self.text["ok_button"])
        dialog.exec()

    def show_help(self) -> None:
        self.show_dialog(self.text["help_title"], self.text["help_body"])

    def open_dependency_action(self) -> None:
        QDesktopServices.openUrl(
            QUrl(getattr(self, "pdf_engine_download_url", LIBREOFFICE_DOWNLOAD_URL))
        )

    def confirm_runtime_dependencies(
        self, settings: GuiExportSettings, paths: list[Path]
    ) -> bool:
        missing = []
        for path in paths:
            for status in dependency_statuses(settings.export_options(path)):
                if status.required and not status.available:
                    key = (status.status_code, status.name, status.action_url)
                    if key not in {
                        (item.status_code, item.name, item.action_url)
                        for item in missing
                    }:
                        missing.append(status)
        if not missing:
            return True

        dialog = QMessageBox(QApplication.activeWindow() or self)
        dialog.setIcon(QMessageBox.Icon.Warning)
        dialog.setWindowTitle(self.text["dependency_title"])
        dialog.setText(
            self.text["dependency_missing_ffmpeg"]
            if all(item.status_code == "ffmpeg_missing" for item in missing)
            else self.text["dependency_missing_libreoffice"]
        )
        dialog.setInformativeText("\n".join(f"• {item.detail}" for item in missing))
        dialog.setStandardButtons(QMessageBox.StandardButton.Cancel)

        primary = missing[0]
        action_url = primary.action_url or LIBREOFFICE_DOWNLOAD_URL
        if primary.status_code.endswith("permission_denied"):
            action_label = self.text["dependency_open_settings"]
        elif primary.status_code == "ffmpeg_missing":
            action_label = self.text["dependency_install_ffmpeg"]
        elif primary.status_code == "keynote_missing":
            action_label = self.text["dependency_install_keynote"]
        elif primary.status_code == "pages_missing":
            action_label = self.text["dependency_install_pages"]
        else:
            action_label = self.text["dependency_install_libreoffice"]
        action_button = dialog.addButton(
            action_label, QMessageBox.ButtonRole.ActionRole
        )
        dialog.exec()
        if dialog.clickedButton() is action_button:
            QDesktopServices.openUrl(QUrl(action_url))
        return False

    def load_bundled_watermark_font(self) -> str:
        font_path = bundled_watermark_font_path()
        if font_path.exists():
            font_id = QFontDatabase.addApplicationFont(str(font_path))
            if font_id >= 0:
                families = QFontDatabase.applicationFontFamilies(font_id)
                if families:
                    return families[0]
        return BUNDLED_WATERMARK_FONT_FAMILY

    def collect_available_font_families(self) -> set[str]:
        try:
            families = set(QFontDatabase.families())
        except TypeError:
            families = set(QFontDatabase().families())
        return families

    def source_font_replacement_family(self) -> str:
        if sys.platform == "darwin":
            preferred = ("PingFang SC", "Arial Unicode MS", "Arial")
        elif sys.platform == "win32":
            preferred = ("Microsoft YaHei", "SimHei", "DengXian", "Arial")
        else:
            preferred = (
                "Noto Sans CJK SC",
                "Noto Sans SC",
                "WenQuanYi Micro Hei",
                "Arial",
            )
        normalized = {
            family.casefold(): family for family in self.available_font_families
        }
        for family in preferred:
            if family.casefold() in normalized:
                return normalized[family.casefold()]
        return default_source_font_family()

    def choose_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            self.text["choose_title"],
            "",
            self.text["choose_filter"],
        )
        if selected:
            self.set_files([Path(path) for path in selected])

    def choose_watermark_image(self) -> None:
        image_filter = (
            "Images ("
            + " ".join(f"*{suffix}" for suffix in WATERMARK_IMAGE_INPUT_EXTENSIONS)
            + ")"
        )
        selected, _ = QFileDialog.getOpenFileName(
            self,
            self.text["watermark_pick_image"],
            "",
            image_filter,
        )
        if not selected:
            return
        self.apply_watermark_image_path(Path(selected))

    def apply_watermark_image_path(self, path: Path) -> None:
        resolved = path.expanduser().resolve()
        if (
            not resolved.exists()
            or resolved.suffix.lower() not in WATERMARK_IMAGE_INPUT_EXTENSIONS
        ):
            return
        self.watermark_image_path = resolved
        prefix = "图片" if self.language == "zh" else "Image"
        self.pick_watermark_image_button.setText(f"{prefix}: {resolved.name}")
        self.pick_watermark_image_button.setToolTip(str(resolved))
        self.watermark_image_name_label.set_full_text(resolved.name)
        self.watermark_image_name_label.setToolTip(str(resolved))
        self.on_settings_changed()

    def set_files(self, paths: list[Path]) -> None:
        if self.is_running:
            return
        seen = set(self.input_paths)
        for path in paths:
            resolved = path.expanduser().resolve()
            if (
                resolved.suffix.lower() not in SUPPORTED_WATERMARK_INPUT_EXTENSIONS
                or resolved in seen
            ):
                continue
            seen.add(resolved)
            self.input_paths.append(resolved)
            self.checked_paths.add(resolved)
            self.file_statuses[resolved] = "pending"
        self.refresh_file_list()
        if self.input_paths and not self.file_list.selectedItems():
            self.file_list.setCurrentRow(0)
        self.update_idle_status_label()
        self.update_font_check()
        self.update_output_path_hint()
        self.schedule_preview_refresh()

    def remove_selected_files(self) -> None:
        if self.is_running:
            return
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
            if item.data(Qt.ItemDataRole.UserRole)
        }
        if not selected_paths:
            selected_paths = set(self.checked_paths)
        if not selected_paths:
            return
        self.input_paths = [
            path for path in self.input_paths if path not in selected_paths
        ]
        for path in selected_paths:
            self.checked_paths.discard(path)
            self.file_statuses.pop(path, None)
            self.file_outputs.pop(path, None)
            self.drop_preview_source(path)
            source_path = str(path.expanduser().resolve())
            for key in list(self.font_report_cache):
                if key[0] == source_path:
                    self.font_report_cache.pop(key, None)
        self.refresh_file_list()
        self.update_idle_status_label()
        self.update_output_path_hint()
        self.update_font_check()
        self.schedule_preview_refresh()

    def set_path_checked(
        self, path: Path, checked: bool, row: QWidget | None = None
    ) -> None:
        if checked:
            self.checked_paths.add(path)
        else:
            self.checked_paths.discard(path)
        if row is not None:
            row.setProperty("included", "true" if checked else "false")
            row.style().unpolish(row)
            row.style().polish(row)
        self.update_selection_summary()
        self.update_remove_button_state()
        self.sync_run_button_state()

    def update_selection_summary(self) -> None:
        selected = sum(path in self.checked_paths for path in self.input_paths)
        self.selection_summary_label.setText(
            self.text["selected_summary"].format(
                selected=selected,
                total=len(self.input_paths),
            )
        )

    def on_selection_changed(self) -> None:
        self.update_remove_button_state()
        self.update_file_item_selection_styles()
        self.update_format_dropdown_state()
        self.update_image_quality_control()
        self.update_watermark_controls()
        self.update_idle_status_label()
        self.update_output_path_hint()
        self.update_font_check()
        self.update_dependency_hint()
        self.schedule_preview_refresh()

    def update_file_item_selection_styles(self) -> None:
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
            if item.data(Qt.ItemDataRole.UserRole)
        }
        for index in range(self.file_list.count()):
            item = self.file_list.item(index)
            raw_path = item.data(Qt.ItemDataRole.UserRole)
            widget = self.file_list.itemWidget(item)
            if not raw_path or widget is None:
                continue
            widget.setProperty(
                "selected", "true" if Path(raw_path) in selected_paths else "false"
            )
            widget.style().unpolish(widget)
            widget.style().polish(widget)

    def selected_path(self) -> Path | None:
        current_item = self.file_list.currentItem()
        if current_item is not None:
            raw_path = current_item.data(Qt.ItemDataRole.UserRole)
            if raw_path:
                return Path(raw_path)
        for item in self.file_list.selectedItems():
            raw_path = item.data(Qt.ItemDataRole.UserRole)
            if raw_path:
                return Path(raw_path)
        if self.input_paths:
            return self.input_paths[0]
        return None

    def current_output_context_kind(
        self,
        path: Path | None = None,
        source_paths: list[Path] | None = None,
    ) -> str:
        if source_paths:
            if any(item.suffix.lower() == ".pptx" for item in source_paths):
                return "pptx"
            if any(item.suffix.lower() in (".docx", ".pdf") for item in source_paths):
                return "document"
            if any(
                item.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS)
                for item in source_paths
            ):
                return "media"
        probe = path or self.selected_path()
        if probe is None:
            return "pptx"
        suffix = probe.suffix.lower()
        if suffix == ".pptx":
            return "pptx"
        if suffix in (".docx", ".pdf"):
            return "document"
        if suffix in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS):
            return "media"
        return "pptx"

    @staticmethod
    def set_combo_item_enabled(combo: QComboBox, value: str, enabled: bool) -> None:
        index = combo.findData(value)
        if index < 0:
            return
        model = combo.model()
        item = model.item(index) if hasattr(model, "item") else None
        if item is not None and item.isEnabled() != enabled:
            item.setEnabled(enabled)

    def resolved_primary_values(
        self, source_paths: list[Path] | None = None
    ) -> tuple[str, str]:
        context_kind = self.current_output_context_kind(source_paths=source_paths)
        current_format = self.output_format_select.currentData() or "pdf"
        current_mode = self.output_mode_select.currentData() or "editable"

        if source_paths:
            if context_kind == "pptx":
                output_format = (
                    current_format
                    if current_format in {"pdf", "pptx"}
                    else self._preferred_output_format
                )
                output_mode = (
                    current_mode
                    if current_mode in {"editable", "image"}
                    else self._preferred_output_mode
                )
                return output_format, output_mode
            if context_kind == "document":
                output_mode = (
                    current_mode
                    if current_mode in {"editable", "image"}
                    else self._preferred_output_mode
                )
                return "pdf", output_mode
            return "source", self._preferred_output_mode

        if context_kind == "pptx":
            output_format = (
                current_format
                if current_format in {"pdf", "pptx"}
                else self._preferred_output_format
            )
            output_mode = (
                current_mode
                if current_mode in {"editable", "image"}
                else self._preferred_output_mode
            )
            return output_format, output_mode
        if context_kind == "document":
            output_mode = (
                current_mode
                if current_mode in {"editable", "image"}
                else self._preferred_output_mode
            )
            return "pdf", output_mode
        return "source", self._preferred_output_mode

    def sync_primary_preferences_from_controls(self) -> None:
        if self._primary_context_lock:
            return
        context_kind = self.current_output_context_kind()
        current_format = self.output_format_select.currentData()
        current_mode = self.output_mode_select.currentData()
        if context_kind == "pptx" and current_format in {"pdf", "pptx"}:
            self._preferred_output_format = current_format
        if context_kind in {"pptx", "document"} and current_mode in {
            "editable",
            "image",
        }:
            self._preferred_output_mode = current_mode

    def queue_primary_values(
        self, source_paths: list[Path] | None = None
    ) -> tuple[str, str]:
        paths = source_paths or self.input_paths
        context_kind = self.current_output_context_kind(source_paths=paths)
        preferred_format = (
            self._preferred_output_format
            if self._preferred_output_format in {"pdf", "pptx"}
            else "pdf"
        )
        preferred_mode = (
            self._preferred_output_mode
            if self._preferred_output_mode in {"editable", "image"}
            else "editable"
        )
        if context_kind == "pptx":
            return preferred_format, preferred_mode
        if context_kind == "document":
            return "pdf", preferred_mode
        return "source", preferred_mode

    def selected_path_uses_fixed_pdf(self) -> bool:
        selected = self.selected_path()
        queue_output_format, _ = self.queue_primary_values(self.input_paths)
        has_pptx = any(path.suffix.lower() == ".pptx" for path in self.input_paths)
        has_media = any(
            path.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS)
            for path in self.input_paths
        )
        return bool(
            selected is not None
            and selected.suffix.lower() in (".docx", ".pdf")
            and (
                (has_pptx and queue_output_format == "pptx")
                or (has_media and not has_pptx)
            )
        )

    def queue_uses_fixed_pdf_outputs(self) -> bool:
        queue_output_format, _ = self.queue_primary_values(self.input_paths)
        has_pptx = any(path.suffix.lower() == ".pptx" for path in self.input_paths)
        has_doc_pdf = any(
            path.suffix.lower() in (".docx", ".pdf") for path in self.input_paths
        )
        has_media = any(
            path.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS)
            for path in self.input_paths
        )
        return bool(
            self.input_paths
            and has_doc_pdf
            and (
                (has_pptx and queue_output_format == "pptx")
                or (has_media and not has_pptx)
            )
        )

    def selected_path_uses_fixed_media_output(self) -> bool:
        selected = self.selected_path()
        return bool(
            selected is not None
            and selected.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS)
        )

    def queue_uses_fixed_media_outputs(self) -> bool:
        has_media = any(
            path.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS)
            for path in self.input_paths
        )
        if not has_media:
            return False
        return any(
            path.suffix.lower() in (".pptx", ".docx", ".pdf")
            for path in self.input_paths
        )

    def update_idle_status_label(self) -> None:
        count = len(self.input_paths)
        if count <= 0:
            self.queue_count_label.setText(self.text["queue_summary_empty"])
            if not self.is_running:
                self.current_file_label.setText(self.text["status_ready"])
                self.queue_count_label.setToolTip(self.text["status_ready"])
                self.sync_run_button_state()
            return
        summary_key = "queue_summary"
        if self.queue_uses_fixed_media_outputs():
            summary_key = "queue_summary_mixed_fixed_output"
        elif self.queue_uses_fixed_pdf_outputs():
            summary_key = "queue_summary_mixed_fixed_pdf"
        summary = self.text[summary_key].format(count=count)
        self.queue_count_label.setText(self.text["queue_summary"].format(count=count))
        self.queue_count_label.setToolTip(summary)
        if self.is_running:
            return
        if self.selected_path_uses_fixed_media_output():
            message_key = "selection_fixed_media"
        elif self.selected_path_uses_fixed_pdf():
            message_key = "selection_fixed_pdf"
        else:
            message_key = "selection_ready"
        message = self.text[message_key].format(count=count)
        self.current_file_label.setText(message)
        self.queue_count_label.setToolTip(f"{summary}\n{message}")
        self.sync_run_button_state()

    def sync_run_button_state(self) -> None:
        if not hasattr(self, "run_button"):
            return
        if self.is_running:
            if self.run_button.text() != self.text["stopping"]:
                self.run_button.setEnabled(True)
            self.run_button.setToolTip(self.text["stop"])
            return
        enabled = bool(self.pending_paths())
        self.run_button.setEnabled(enabled)
        tooltip = self.current_file_label.text().strip() or self.text["status_ready"]
        self.run_button.setToolTip(tooltip)

    def current_image_quality_profile(self) -> tuple[int, int]:
        profile = self.dpi_select.currentData()
        if isinstance(profile, tuple) and len(profile) == 2:
            return (int(profile[0]), int(profile[1]))
        return (240, 90)

    def current_video_quality_profile(self) -> str:
        index = self.dpi_select.currentIndex()
        if index <= 0:
            return "none"
        if index == 1:
            return "high"
        if index == 2:
            return "balanced"
        return "aggressive"

    def effective_output_format_for_path(
        self, path: Path, requested_output_format: str | None = None
    ) -> str:
        return effective_export_output_format(
            path,
            requested_output_format
            or (self.output_format_select.currentData() or "pdf"),
        )

    def font_names_by_path(self, paths: list[Path]) -> dict[str, tuple[str, ...]]:
        result: dict[str, tuple[str, ...]] = {}
        for path in paths:
            try:
                key = preview_cache_key(path)
                report = self.font_report_cache.get(key)
                if report is None:
                    report = scan_missing_fonts(path, self.available_font_families)
                    self.font_report_cache[key] = report
                result[str(path.expanduser().resolve())] = report.missing
            except Exception:
                result[str(path.expanduser().resolve())] = ()
        return result

    def current_settings(
        self, source_paths: list[Path] | None = None
    ) -> GuiExportSettings:
        replace_source_fonts = self.font_fix_button.isChecked()
        source_font_names_by_path = (
            self.font_names_by_path(source_paths)
            if replace_source_fonts and source_paths is not None
            else None
        )
        dpi, jpeg_quality = self.current_image_quality_profile()
        if source_paths is None:
            output_format, output_mode = self.resolved_primary_values()
        else:
            output_format, output_mode = self.queue_primary_values(source_paths)
        return GuiExportSettings(
            output_format=output_format,
            output_mode=output_mode,
            dpi=dpi,
            jpeg_quality=jpeg_quality,
            video_quality_profile=self.current_video_quality_profile(),
            preserve_videos_in_image_pptx=self.keep_videos_checkbox.isChecked(),
            video_encoder="auto",
            watermark_enabled=self.watermark_checkbox.isChecked(),
            watermark_kind="image"
            if self.watermark_image_kind_button.isChecked()
            else "text",
            watermark_text=self.watermark_text_input.text().strip()
            or DEFAULT_WATERMARK_TEXT,
            watermark_image_path=self.watermark_image_path,
            watermark_image_width=int(self.image_width_input.value()),
            watermark_color=self.color_select.currentData(),
            watermark_opacity=float(self.opacity_input.value()),
            watermark_font_size=int(self.font_size_input.value()),
            watermark_spacing=int(self.spacing_input.value()),
            watermark_angle=float(self.angle_input.value()),
            watermark_bold=True,
            replace_source_fonts=replace_source_fonts
            and (bool(self.current_missing_fonts) or bool(source_font_names_by_path)),
            source_font_family=self.source_font_replacement_family(),
            source_font_names=self.current_missing_fonts,
            source_font_names_by_path=source_font_names_by_path,
        )

    def on_settings_changed(self) -> None:
        self.sync_primary_preferences_from_controls()
        self.update_image_quality_control()
        self.update_watermark_controls()
        self.update_preset_hint()
        self.update_font_check()
        self.update_dependency_hint()
        self.mark_completed_files_pending_due_to_settings()
        self.update_idle_status_label()
        self.update_output_path_hint()
        self.refresh_sample_preview_if_needed()
        self.schedule_preview_refresh()
        self.refresh_file_list()

    def apply_watermark_preset_from_select(self) -> None:
        if self.watermark_image_kind_button.isChecked():
            self.update_preset_hint()
            return
        preset = self.preset_select.currentData()
        if not preset:
            return
        widgets = [
            self.color_select,
            self.opacity_input,
            self.font_size_input,
            self.spacing_input,
            self.angle_input,
        ]
        previous_blocks = [widget.blockSignals(True) for widget in widgets]
        try:
            color_index = self.color_select.findData(preset["color"])
            if color_index >= 0:
                self.color_select.setCurrentIndex(color_index)
            self.opacity_input.setValue(float(preset["opacity"]))
            self.font_size_input.setValue(int(preset["font_size"]))
            self.spacing_input.setValue(int(preset["spacing"]))
            self.angle_input.setValue(int(preset["angle"]))
        finally:
            for widget, was_blocked in zip(widgets, previous_blocks):
                widget.blockSignals(was_blocked)
        self.on_settings_changed()

    def update_preset_hint(self) -> None:
        if not hasattr(self, "preset_hint_label"):
            return
        if (
            hasattr(self, "watermark_image_kind_button")
            and self.watermark_image_kind_button.isChecked()
        ):
            image_name = (
                self.watermark_image_path.name
                if self.watermark_image_path
                else self.text["watermark_no_image"]
            )
            detail = (
                f"{image_name}  "
                f"{self.opacity_input.value():.2f} · "
                f"{int(self.image_width_input.value())}px · "
                f"{int(self.spacing_input.value())}px · "
                f"{int(self.angle_input.value())}°"
            )
            self.preset_hint_label.setText(image_name)
            self.preset_hint_label.setToolTip(
                str(self.watermark_image_path) if self.watermark_image_path else detail
            )
            return
        preset = self.preset_select.currentData()
        if not preset:
            self.preset_hint_label.setText("")
            self.preset_hint_label.setToolTip("")
            return
        desc = preset[f"{self.language}_desc"]
        detail = (
            f"{desc}  "
            f"{self.color_select.currentData()} · "
            f"{self.opacity_input.value():.2f} · "
            f"{int(self.font_size_input.value())}pt · "
            f"{int(self.spacing_input.value())}px · "
            f"{int(self.angle_input.value())}°"
        )
        self.preset_hint_label.setText(desc)
        self.preset_hint_label.setToolTip(detail)

    def update_font_check(self) -> None:
        if not hasattr(self, "font_check_label"):
            return
        selected = self.selected_path()
        if selected is None or not selected.exists():
            self.current_font_report = None
            self.current_missing_fonts = ()
            self.font_check_label.setText(self.text["font_check_waiting"])
            self.font_check_label.setToolTip("")
            self.font_fix_button.setChecked(False)
            self._sync_font_fix_button()
            return
        if selected.suffix.lower() != ".pptx":
            self.current_font_report = None
            self.current_missing_fonts = ()
            self.font_check_label.setText(self.text["font_check_non_pptx"])
            self.font_check_label.setToolTip(self.text["font_check_non_pptx"])
            self.font_fix_button.setChecked(False)
            self._sync_font_fix_button()
            return
        try:
            key = preview_cache_key(selected)
            report = self.font_report_cache.get(key)
            if report is None:
                report = scan_missing_fonts(selected, self.available_font_families)
                self.font_report_cache[key] = report
        except Exception as exc:
            self.current_font_report = None
            self.current_missing_fonts = ()
            self.font_check_label.setText(f"{self.text['font_check_title']}: {exc}")
            self.font_check_label.setToolTip(str(exc))
            self.font_fix_button.setChecked(False)
            self._sync_font_fix_button()
            return

        self.current_font_report = report
        self.current_missing_fonts = report.missing
        if not report.missing:
            self.font_check_label.setText(self.text["font_check_ok"])
            self.font_check_label.setToolTip(
                ", ".join(report.referenced)
                if report.referenced
                else self.text["font_check_ok"]
            )
            self.font_fix_button.setChecked(False)
        else:
            visible_fonts = ", ".join(report.missing[:4])
            if len(report.missing) > 4:
                visible_fonts += f" +{len(report.missing) - 4}"
            message = self.text["font_check_missing"].format(
                count=len(report.missing),
                fonts=visible_fonts,
                family=self.source_font_replacement_family(),
            )
            self.font_check_label.setText(message)
            self.font_check_label.setToolTip(", ".join(report.missing))
        self._sync_font_fix_button()

    def _sync_font_fix_button(self) -> None:
        if not hasattr(self, "font_fix_button"):
            return
        enabled = bool(self.current_missing_fonts) and not self.is_running
        self.font_fix_button.setEnabled(enabled)
        self.font_fix_button.setText(
            self.text["font_fix_on"]
            if self.font_fix_button.isChecked()
            else self.text["font_fix_off"]
        )
        self.font_fix_button.setProperty(
            "active", "true" if self.font_fix_button.isChecked() else "false"
        )
        self.font_fix_button.style().unpolish(self.font_fix_button)
        self.font_fix_button.style().polish(self.font_fix_button)

    def mark_completed_files_pending_due_to_settings(self) -> None:
        if self.is_running:
            return
        changed = False
        for path, state in list(self.file_statuses.items()):
            if state in {"done", "failed", "stopped"}:
                self.file_statuses[path] = "pending"
                self.file_outputs.pop(path, None)
                changed = True
        if changed:
            self.update_idle_status_label()

    def update_watermark_controls(self) -> None:
        enabled = self.watermark_checkbox.isChecked() and not self.is_running
        is_image_watermark = self.watermark_image_kind_button.isChecked()
        self.watermark_checkbox.setText(
            self.text["watermark_toggle_on"]
            if self.watermark_checkbox.isChecked()
            else self.text["watermark_toggle_off"]
        )
        self.watermark_checkbox.setProperty(
            "active", "true" if self.watermark_checkbox.isChecked() else "false"
        )
        self.watermark_checkbox.style().unpolish(self.watermark_checkbox)
        self.watermark_checkbox.style().polish(self.watermark_checkbox)
        self.watermark_text_group.setVisible(not is_image_watermark)
        self.watermark_preset_group.setVisible(not is_image_watermark)
        self.watermark_image_group.setVisible(is_image_watermark)
        self.watermark_image_width_group.setVisible(is_image_watermark)
        for widget in (
            self.watermark_text_kind_button,
            self.watermark_image_kind_button,
            self.watermark_text_input,
            self.preset_select,
            self.opacity_input,
            self.spacing_input,
            self.angle_input,
            self.pick_watermark_image_button,
            self.image_width_input,
        ):
            widget.setEnabled(enabled)
        self.color_select.setEnabled(enabled and not is_image_watermark)
        self.font_size_input.setEnabled(enabled and not is_image_watermark)

    def update_image_quality_control(self) -> None:
        has_pptx = any(path.suffix.lower() == ".pptx" for path in self.input_paths)
        context_kind = (
            self.current_output_context_kind(source_paths=self.input_paths)
            if has_pptx
            else self.current_output_context_kind()
        )
        resolved_format, resolved_mode = self.resolved_primary_values(
            self.input_paths if has_pptx else None
        )
        is_image_mode = resolved_mode == "image"
        is_image_pptx = is_image_mode and resolved_format == "pptx"
        enabled = (
            (context_kind in {"pptx", "document"} and is_image_mode)
            or context_kind == "media"
        ) and not self.is_running
        self.dpi_select.setEnabled(enabled)
        if context_kind == "media":
            quality_tooltip = self.text["media_quality_locked_hint"]
        elif context_kind in {"pptx", "document"} and not is_image_mode:
            quality_tooltip = self.text["image_quality_mode_hint"]
        else:
            quality_tooltip = ""
        self.dpi_select.setToolTip(quality_tooltip)
        self.keep_videos_group.setVisible(context_kind == "pptx")
        can_reinsert_videos = is_image_pptx and not self.is_running
        self.keep_videos_checkbox.setEnabled(can_reinsert_videos)
        self.keep_videos_checkbox.setText(
            self.text["image_keep_videos_on"]
            if self.keep_videos_checkbox.isChecked()
            else self.text["image_keep_videos_off"]
        )
        self.keep_videos_checkbox.setToolTip(
            (
                f"{self.text['image_keep_videos_hint']}\n"
                f"{self.text['image_keep_videos_warning']}"
                if can_reinsert_videos
                else self.text["image_keep_videos_unavailable"]
            )
        )
        self.keep_videos_checkbox.setProperty(
            "active",
            "true"
            if self.keep_videos_checkbox.isChecked() and is_image_pptx
            else "false",
        )
        self.keep_videos_checkbox.style().unpolish(self.keep_videos_checkbox)
        self.keep_videos_checkbox.style().polish(self.keep_videos_checkbox)

    def update_dependency_hint(self) -> None:
        probe_path = self.selected_path() or Path("placeholder.pptx")
        options = self.current_settings().export_options(probe_path)
        statuses = [
            status for status in dependency_statuses(options) if status.required
        ]
        missing_statuses = [status for status in statuses if not status.available]
        primary_missing = missing_statuses[0] if missing_statuses else None
        status_codes = {status.status_code for status in statuses}
        using_keynote_fallback = "keynote_fallback" in status_codes
        using_pages_fallback = "pages_fallback" in status_codes

        if primary_missing is not None and primary_missing.status_code in {
            "keynote_permission_denied",
            "pages_permission_denied",
        }:
            text = (
                self.text["dependency_permission_keynote"]
                if primary_missing.status_code == "keynote_permission_denied"
                else self.text["dependency_permission_pages"]
            )
            self.install_libreoffice_button.setText(
                self.text["dependency_open_settings"]
            )
            self.pdf_engine_download_url = (
                primary_missing.action_url or LIBREOFFICE_DOWNLOAD_URL
            )
        elif (
            primary_missing is not None
            and primary_missing.status_code == "ffmpeg_missing"
        ):
            text = self.text["dependency_missing_ffmpeg"]
            self.pdf_engine_download_url = LIBREOFFICE_DOWNLOAD_URL
            self.install_libreoffice_button.setText(
                self.text["dependency_install_libreoffice"]
            )
        elif primary_missing is not None:
            text = self.text["dependency_missing_libreoffice"]
            self.pdf_engine_download_url = (
                primary_missing.action_url or LIBREOFFICE_DOWNLOAD_URL
            )
            if primary_missing.status_code.startswith("pages_"):
                self.install_libreoffice_button.setText(
                    self.text["dependency_install_pages"]
                    if "missing" in primary_missing.status_code
                    else self.text["dependency_install_libreoffice"]
                )
            elif primary_missing.status_code.startswith("keynote_"):
                self.install_libreoffice_button.setText(
                    self.text["dependency_install_keynote"]
                    if "missing" in primary_missing.status_code
                    else self.text["dependency_install_libreoffice"]
                )
            else:
                self.install_libreoffice_button.setText(
                    self.text["dependency_install_libreoffice"]
                )
        elif using_keynote_fallback:
            text = self.text["dependency_keynote_fallback"]
            self.install_libreoffice_button.setText(
                self.text["dependency_install_libreoffice"]
            )
            self.pdf_engine_download_url = LIBREOFFICE_DOWNLOAD_URL
        elif using_pages_fallback:
            text = self.text["dependency_pages_fallback"]
            self.install_libreoffice_button.setText(
                self.text["dependency_install_libreoffice"]
            )
            self.pdf_engine_download_url = LIBREOFFICE_DOWNLOAD_URL
        else:
            names = ", ".join(status.name for status in statuses)
            text = f"{self.text['dependency_ok_short']} {names}".strip()
            self.pdf_engine_download_url = LIBREOFFICE_DOWNLOAD_URL

        details = missing_dependency_message(options) or "\n".join(
            f"{status.name}: {status.detail or status.path or 'available'}"
            for status in statuses
        )
        self.dependency_hint_label.setText(text)
        self.dependency_hint_label.setToolTip(details)
        should_show_action = (
            primary_missing is not None
            or using_keynote_fallback
            or using_pages_fallback
        )
        if (
            primary_missing is not None
            and primary_missing.status_code == "ffmpeg_missing"
        ):
            should_show_action = False
        self.install_libreoffice_button.setVisible(should_show_action)

    def schedule_preview_refresh(self) -> None:
        if self.is_running:
            return
        selected = self.selected_path()
        if selected is None:
            self.clear_preview(self.text["preview_waiting"])
            return
        try:
            source_key = preview_cache_key(selected)
        except OSError:
            self.clear_preview(self.text["preview_waiting"])
            return
        if (
            selected.stat().st_size > AUTO_PREVIEW_MAX_BYTES
            and source_key not in self.preview_source_cache
        ):
            self.clear_preview(self.text["preview_manual"])
            return
        self.preview_timer.start()

    def force_preview_refresh(self) -> None:
        self.preview_timer.stop()
        selected = self.selected_path()
        if selected is not None:
            self.drop_preview_source(selected)
        self.start_preview_for_selection()

    def drop_preview_source(self, path: Path) -> None:
        source_path = str(path.expanduser().resolve())
        for key in list(self.preview_source_cache):
            if key[0] != source_path:
                continue
            source = self.preview_source_cache.pop(key)
            if self.preview_worker is not None and self.preview_worker.source is source:
                self.preview_source_cleanup_pending.add(source.temp_root)
                continue
            shutil.rmtree(source.temp_root, ignore_errors=True)

    def start_preview_for_selection(self) -> None:
        selected = self.selected_path()
        if selected is None or not selected.exists():
            self.clear_preview(self.text["preview_waiting"])
            return
        settings = self.current_settings()
        if (
            settings.watermark_enabled
            and settings.watermark_kind == "image"
            and settings.watermark_image_path is None
        ):
            self.clear_preview(self.text["watermark_image_missing_body"])
            return
        preview_options = ExportOptions(
            input_path=selected,
            output_format="pdf",
            output_mode="image",
            dpi=self.current_image_quality_profile()[0],
            jpeg_quality=self.current_image_quality_profile()[1],
            keep_artifacts=False,
            watermark=settings.watermark(),
        )
        dependency_message = missing_dependency_message(preview_options)
        if dependency_message:
            self.clear_preview(f"{self.text['preview_failed']}: {dependency_message}")
            return
        if self.preview_worker is not None:
            self.preview_dirty = True
            return
        try:
            source_key = preview_cache_key(selected)
        except OSError as exc:
            self.clear_preview(f"{self.text['preview_failed']}: {exc}")
            return
        source = self.preview_source_cache.get(source_key)
        self.preview_request_id += 1
        request_id = self.preview_request_id
        self.preview_page_label.setText(self.text["preview_loading"])
        # Page progress lives in the thumbnail strip; keeping transient status in
        # the compact toolbar squeezes the three preview actions at 960 px.
        self.preview_page_label.hide()
        self.preview_image_label.clear_preview(
            self.text["preview_loading"],
            self.sample_watermark_options(),
            self.sample_preview_theme(),
        )
        self.preview_thread = QThread(self)
        self.preview_worker = PreviewWorker(
            request_id, selected, settings, source_key, source
        )
        self.preview_worker.moveToThread(self.preview_thread)
        self.preview_thread.started.connect(self.preview_worker.run)
        self.preview_worker.finished.connect(self.on_preview_finished)
        self.preview_worker.failed.connect(self.on_preview_failed)
        self.preview_worker.finished.connect(self.preview_worker.deleteLater)
        self.preview_worker.failed.connect(self.preview_worker.deleteLater)
        self.preview_thread.finished.connect(self.preview_thread.deleteLater)
        self.preview_thread.start()

    def on_preview_finished(
        self, request_id: int, input_path: object, artifacts: object, source_obj: object
    ) -> None:
        path = Path(input_path)
        preview: PreviewArtifacts = artifacts
        source: PreviewSource = source_obj
        if request_id != self.preview_request_id or path != self.selected_path():
            if preview.temp_root is not None:
                shutil.rmtree(preview.temp_root, ignore_errors=True)
            if source.key not in self.preview_source_cache:
                shutil.rmtree(source.temp_root, ignore_errors=True)
        else:
            self.preview_source_cache[source.key] = source
            self.cleanup_current_preview()
            self.current_preview = preview
            self.current_preview_page = 0
            self.update_preview_display()
        self.finish_preview_thread()

    def on_preview_failed(
        self, request_id: int, input_path: object, message: str
    ) -> None:
        if (
            request_id == self.preview_request_id
            and Path(input_path) == self.selected_path()
        ):
            self.clear_preview(f"{self.text['preview_failed']}: {message}")
        self.finish_preview_thread()

    def finish_preview_thread(self) -> None:
        if self.preview_thread is not None:
            thread = self.preview_thread
            thread.quit()
            if not thread.wait(5000):
                while not thread.wait(250):
                    terminate_active_processes(grace_seconds=0.1)
            self.preview_thread = None
        self.preview_worker = None
        for temp_root in self.preview_source_cleanup_pending:
            shutil.rmtree(temp_root, ignore_errors=True)
        self.preview_source_cleanup_pending.clear()
        if self.preview_dirty:
            self.preview_dirty = False
            self.preview_timer.start()

    @staticmethod
    def shutdown_thread(thread: QThread | None, wait_ms: int = 5000) -> None:
        if thread is None or not thread.isRunning():
            return
        terminate_active_processes(grace_seconds=0.1)
        thread.quit()
        while not thread.wait(min(wait_ms, 250)):
            terminate_active_processes(grace_seconds=0.1)

    def cleanup_current_preview(self) -> None:
        if self.current_preview is not None:
            if self.current_preview.temp_root is not None:
                shutil.rmtree(self.current_preview.temp_root, ignore_errors=True)
            self.current_preview = None

    def cleanup_preview_sources(self) -> None:
        for source in self.preview_source_cache.values():
            shutil.rmtree(source.temp_root, ignore_errors=True)
        self.preview_source_cache.clear()
        for temp_root in self.preview_source_cleanup_pending:
            shutil.rmtree(temp_root, ignore_errors=True)
        self.preview_source_cleanup_pending.clear()

    def sample_watermark_options(self) -> WatermarkOptions | None:
        if not hasattr(self, "watermark_checkbox"):
            return None
        return self.current_settings().watermark()

    def sample_preview_theme(self) -> dict[str, str]:
        if not hasattr(self, "preview_background_select"):
            return PREVIEW_BACKGROUND_THEMES[0]
        theme = self.preview_background_select.currentData()
        return theme if theme else PREVIEW_BACKGROUND_THEMES[0]

    def refresh_sample_preview_if_needed(self) -> None:
        if self.current_preview is not None or self.preview_worker is not None:
            return
        message = (
            self.text["preview_waiting"]
            if not self.input_paths
            else self.preview_page_label.text().strip() or self.text["preview_waiting"]
        )
        self.preview_image_label.clear_preview(
            message, self.sample_watermark_options(), self.sample_preview_theme()
        )
        show_spread = self.preview_uses_spread()
        if show_spread:
            self.preview_image_label_secondary.clear_preview(
                message,
                self.sample_watermark_options(),
                self.sample_preview_theme(),
            )
        self.preview_image_label_secondary.setVisible(show_spread)
        self.preview_thumbnail_list.show()
        self.update_preview_geometry()

    def preview_uses_spread(self) -> bool:
        selected = self.selected_path()
        return selected is None or selected.suffix.lower() == ".pptx"

    def clear_preview(self, message: str) -> None:
        self.cleanup_current_preview()
        self.preview_thumbnail_list.clear()
        placeholder = QListWidgetItem(self.text["preview_thumbnails_empty"])
        placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
        self.preview_thumbnail_list.addItem(placeholder)
        self.preview_page_label.setText(message)
        self.preview_page_label.setVisible(False)
        self.preview_image_label.clear_preview(
            message, self.sample_watermark_options(), self.sample_preview_theme()
        )
        show_spread = self.preview_uses_spread()
        if show_spread:
            self.preview_image_label_secondary.clear_preview(
                message,
                self.sample_watermark_options(),
                self.sample_preview_theme(),
            )
        self.preview_image_label_secondary.setVisible(show_spread)
        self.preview_thumbnail_list.show()
        self.update_preview_geometry()
        self.update_preview_buttons()

    def set_preview_mode(self, mode: str) -> None:
        self.preview_mode = mode
        self.original_preview_button.setChecked(mode == "original")
        self.rendered_preview_button.setChecked(mode == "preview")
        self.update_preview_display()

    def change_preview_page(self, delta: int) -> None:
        if self.current_preview is None:
            return
        page_count = len(self.current_preview.original_paths)
        if page_count <= 0:
            return
        step = self.preview_group_size()
        self.current_preview_page = max(
            0, min(page_count - 1, self.current_preview_page + (delta * step))
        )
        self.update_preview_display()

    def select_preview_page(self, page_index: int) -> None:
        if self.current_preview is None or page_index < 0:
            return
        page_count = len(self.preview_paths_for_mode())
        if page_index >= page_count or page_index == self.current_preview_page:
            return
        self.current_preview_page = page_index
        self.update_preview_display()

    def update_preview_thumbnails(self, paths: list[Path]) -> None:
        with QSignalBlocker(self.preview_thumbnail_list):
            self.preview_thumbnail_list.clear()
            for index, path in enumerate(paths):
                pixmap = QPixmap(str(path))
                item = QListWidgetItem(
                    QIcon(pixmap),
                    str(index + 1),
                )
                item.setToolTip(
                    self.text["preview_page"].format(
                        index=index + 1,
                        rendered=len(paths),
                        total=self.current_preview.total_pages
                        if self.current_preview
                        else len(paths),
                    )
                )
                item.setSizeHint(QSize(106, 68))
                self.preview_thumbnail_list.addItem(item)
            self.preview_thumbnail_list.setCurrentRow(self.current_preview_page)
        self.preview_thumbnail_list.show()

    def update_preview_buttons(self) -> None:
        page_count = (
            len(self.current_preview.original_paths) if self.current_preview else 0
        )
        has_preview = page_count > 0
        step = self.preview_group_size()
        self.preview_prev_button.show()
        self.preview_next_button.show()
        self.preview_prev_button.setEnabled(
            has_preview and self.current_preview_page > 0
        )
        self.preview_next_button.setEnabled(
            has_preview and self.current_preview_page + step < page_count
        )

    def preview_paths_for_mode(self) -> list[Path]:
        if self.current_preview is None:
            return []
        return (
            self.current_preview.original_paths
            if self.preview_mode == "original"
            else self.current_preview.preview_paths
        )

    def preview_group_size(self) -> int:
        paths = self.preview_paths_for_mode()
        if not paths:
            return 1
        page_index = min(self.current_preview_page, len(paths) - 1)
        pixmap = QPixmap(str(paths[page_index]))
        if not pixmap.isNull() and pixmap.height() > pixmap.width():
            return 1
        next_index = page_index + 1
        if next_index < len(paths):
            next_pixmap = QPixmap(str(paths[next_index]))
            if not next_pixmap.isNull() and next_pixmap.height() > next_pixmap.width():
                return 1
        return 2

    def update_preview_display(self) -> None:
        if self.current_preview is None:
            self.update_preview_buttons()
            return
        paths = self.preview_paths_for_mode()
        if not paths:
            self.clear_preview(self.text["preview_empty_page"])
            return
        page_index = max(0, min(len(paths) - 1, self.current_preview_page))
        self.current_preview_page = page_index
        pixmap = QPixmap(str(paths[page_index]))
        if pixmap.isNull():
            self.clear_preview(self.text["preview_empty_page"])
            return
        self.preview_image_label.set_preview_pixmap(pixmap)
        next_index = page_index + 1
        has_second_page = False
        if self.preview_group_size() == 2 and next_index < len(paths):
            next_pixmap = QPixmap(str(paths[next_index]))
            if next_pixmap.isNull():
                self.preview_image_label_secondary.hide()
            else:
                self.preview_image_label_secondary.set_preview_pixmap(next_pixmap)
                self.preview_image_label_secondary.show()
                has_second_page = True
        else:
            self.preview_image_label_secondary.hide()
        self.update_preview_thumbnails(paths)
        preview_layout = self.preview_thumbnail_list.parentWidget().layout()
        if preview_layout is not None:
            preview_layout.activate()
        self.update_preview_geometry()
        QTimer.singleShot(0, self.update_preview_geometry)
        if has_second_page:
            self.preview_page_label.setText(
                f"{page_index + 1}-{next_index + 1} / "
                f"{self.current_preview.total_pages}"
            )
        else:
            self.preview_page_label.setText(
                self.text["preview_page"].format(
                    index=page_index + 1,
                    rendered=len(paths),
                    total=self.current_preview.total_pages,
                )
            )
        self.preview_page_label.hide()
        self.update_preview_buttons()

    def default_output_for_path(self, path: Path) -> Path:
        queue_paths = self.input_paths if path in self.input_paths else [path]
        settings = self.current_settings(queue_paths)
        return default_output_path(
            path,
            settings.effective_output_format_for_path(path),
            settings.output_mode,
        )

    @staticmethod
    def source_format_label(path: Path) -> str:
        suffix = path.suffix.lower().lstrip(".")
        return suffix.upper() if suffix else "FILE"

    def output_format_label_for_path(
        self,
        path: Path,
        output_mode: str,
        requested_output_format: str | None = None,
    ) -> str:
        actual_format = self.effective_output_format_for_path(
            path, requested_output_format
        ).upper()
        if path.suffix.lower() in (
            ".docx",
            ".pdf",
            *IMAGE_EXTENSIONS,
            *VIDEO_EXTENSIONS,
        ):
            actual_format = self.text["output_format_fixed"].format(
                format=actual_format
            )
        if path.suffix.lower() in (*IMAGE_EXTENSIONS, *VIDEO_EXTENSIONS):
            return actual_format
        return f"{actual_format} · {output_mode}"

    def update_output_path_hint(self) -> None:
        selected = self.selected_path()
        if selected is None:
            self.output_path_label.set_full_text(self.text["output_waiting"])
            self.output_path_label.setToolTip(self.text["output_waiting"])
            return
        output_path = self.file_outputs.get(selected) or self.default_output_for_path(
            selected
        )
        display = f"{self.text['output_path_label']}: {output_path.name}"
        self.output_path_label.set_full_text(display)
        self.output_path_label.setToolTip(str(output_path))

    def state_label(self, state: str | None) -> str:
        if not state:
            return ""
        key = {
            "pending": "pending_marker",
            "queued": "queued_marker",
            "running": "running_marker",
            "done": "done_marker",
            "failed": "failed_marker",
            "stopped": "stopped_marker",
        }.get(state)
        return self.text[key] if key else ""

    def update_format_dropdown_state(self) -> None:
        context_kind = self.current_output_context_kind()
        self.set_combo_item_enabled(
            self.output_format_select, "pdf", context_kind in {"pptx", "document"}
        )
        self.set_combo_item_enabled(
            self.output_format_select, "pptx", context_kind == "pptx"
        )
        self.set_combo_item_enabled(
            self.output_format_select, "source", context_kind == "media"
        )
        self.set_combo_item_enabled(
            self.output_mode_select, "editable", context_kind in {"pptx", "document"}
        )
        self.set_combo_item_enabled(
            self.output_mode_select, "image", context_kind in {"pptx", "document"}
        )
        self.set_combo_item_enabled(
            self.output_mode_select, "na", context_kind == "media"
        )

        self._primary_context_lock = True
        try:
            if context_kind == "pptx":
                target_format = (
                    self._preferred_output_format
                    if self._preferred_output_format in {"pdf", "pptx"}
                    else "pdf"
                )
                target_mode = (
                    self._preferred_output_mode
                    if self._preferred_output_mode in {"editable", "image"}
                    else "editable"
                )
                format_idx = self.output_format_select.findData(target_format)
                mode_idx = self.output_mode_select.findData(target_mode)
                if (
                    format_idx >= 0
                    and self.output_format_select.currentIndex() != format_idx
                ):
                    blocker = QSignalBlocker(self.output_format_select)
                    self.output_format_select.setCurrentIndex(format_idx)
                    del blocker
                if mode_idx >= 0 and self.output_mode_select.currentIndex() != mode_idx:
                    blocker = QSignalBlocker(self.output_mode_select)
                    self.output_mode_select.setCurrentIndex(mode_idx)
                    del blocker
                self.output_format_select.setDisabled(self.is_running)
                self.output_mode_select.setDisabled(self.is_running)
                self.output_format_select.setToolTip("")
                self.output_mode_select.setToolTip("")
            elif context_kind == "document":
                format_idx = self.output_format_select.findData("pdf")
                target_mode = (
                    self._preferred_output_mode
                    if self._preferred_output_mode in {"editable", "image"}
                    else "editable"
                )
                mode_idx = self.output_mode_select.findData(target_mode)
                if (
                    format_idx >= 0
                    and self.output_format_select.currentIndex() != format_idx
                ):
                    blocker = QSignalBlocker(self.output_format_select)
                    self.output_format_select.setCurrentIndex(format_idx)
                    del blocker
                if mode_idx >= 0 and self.output_mode_select.currentIndex() != mode_idx:
                    blocker = QSignalBlocker(self.output_mode_select)
                    self.output_mode_select.setCurrentIndex(mode_idx)
                    del blocker
                self.output_format_select.setDisabled(True)
                self.output_format_select.setToolTip(
                    self.text["output_format_document_locked_hint"]
                )
                self.output_mode_select.setDisabled(self.is_running)
                self.output_mode_select.setToolTip("")
            else:
                format_idx = self.output_format_select.findData("source")
                mode_idx = self.output_mode_select.findData("na")
                if (
                    format_idx >= 0
                    and self.output_format_select.currentIndex() != format_idx
                ):
                    blocker = QSignalBlocker(self.output_format_select)
                    self.output_format_select.setCurrentIndex(format_idx)
                    del blocker
                if mode_idx >= 0 and self.output_mode_select.currentIndex() != mode_idx:
                    blocker = QSignalBlocker(self.output_mode_select)
                    self.output_mode_select.setCurrentIndex(mode_idx)
                    del blocker
                self.output_format_select.setDisabled(True)
                self.output_format_select.setToolTip(
                    self.text["output_format_media_locked_hint"]
                )
                self.output_mode_select.setDisabled(True)
                self.output_mode_select.setToolTip(self.text["media_mode_locked_hint"])
        finally:
            self._primary_context_lock = False

    def refresh_file_list(self) -> None:
        self.update_format_dropdown_state()

        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
            if item.data(Qt.ItemDataRole.UserRole)
        }
        self.file_list.clear()
        if not self.input_paths:
            self.update_idle_status_label()
            self.update_remove_button_state()
            self.update_selection_summary()
            self.file_list.viewport().update()
            self.sync_run_button_state()
            return

        settings = self.current_settings(self.input_paths)
        icon_provider = QFileIconProvider()
        for path in self.input_paths:
            export_tag = self.output_format_label_for_path(
                path, settings.output_mode, settings.output_format
            )
            source_tag = self.source_format_label(path)
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            item_width = max(1, self.file_list.viewport().width() - 16)
            item.setSizeHint(QSize(item_width, 62))

            row = QWidget()
            row.setObjectName("fileItem")
            state = self.file_statuses.get(path, "pending")
            row.setProperty("state", state)
            row.setProperty("selected", "true" if path in selected_paths else "false")
            row.setProperty(
                "included", "true" if path in self.checked_paths else "false"
            )
            row.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(8, 4, 8, 4)
            row_layout.setSpacing(8)

            type_toggle = QPushButton()
            type_toggle.setObjectName("fileTypeToggle")
            type_toggle.setCheckable(True)
            type_toggle.setChecked(path in self.checked_paths)
            type_toggle.setDisabled(self.is_running)
            type_toggle.setFixedSize(34, 34)
            type_toggle.setIcon(icon_provider.icon(QFileInfo(str(path))))
            type_toggle.setIconSize(QSize(24, 24))
            type_toggle.setAccessibleName(
                self.text["file_toggle_accessible"].format(source=source_tag)
            )
            type_toggle.setToolTip(
                self.text["file_toggle_tooltip"].format(source=source_tag)
            )
            type_toggle.toggled.connect(
                lambda checked, checked_path=path, checked_row=row: (
                    self.set_path_checked(checked_path, checked, checked_row)
                )
            )

            text_widget = QWidget(row)
            text_widget.setSizePolicy(
                QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
            )
            text_col = QVBoxLayout(text_widget)
            text_col.setContentsMargins(0, 0, 0, 0)
            text_col.setSpacing(2)
            name_label = ElidedLabel(path.name)
            name_label.setObjectName("fileName")
            name_label.setToolTip(str(path))
            meta = path.stat().st_size if path.exists() else 0
            meta_label = ElidedLabel(
                f"{format_file_size(meta)} · {source_tag} → {export_tag}"
            )
            meta_label.setObjectName("fileMeta")
            meta_label.setToolTip(
                f"{format_file_size(meta)} · {source_tag} → {export_tag}"
            )
            text_col.addWidget(name_label)
            text_col.addWidget(meta_label)

            status_widget = QWidget(row)
            status_widget.setFixedWidth(52 if self.language == "zh" else 66)
            status_col = QVBoxLayout(status_widget)
            status_col.setContentsMargins(0, 0, 0, 0)
            status_col.setSpacing(2)
            status_col.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            output_name = self.file_outputs.get(path)
            state_label = QLabel(self.state_label(state))
            state_label.setObjectName("fileState")
            state_label.setProperty("state", state)
            state_label.setAlignment(Qt.AlignRight)
            status_col.addWidget(state_label)
            row.setToolTip(
                f"{path}\n"
                f"{self.text['output_path_label']}: "
                f"{output_name if output_name else export_tag}"
            )

            row_layout.addWidget(type_toggle, 0, Qt.AlignVCenter)
            row_layout.addWidget(text_widget, 1)
            row_layout.addWidget(status_widget, 0)
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, row)
            if path in selected_paths:
                item.setSelected(True)

        if not self.file_list.selectedItems() and self.file_list.count() > 0:
            self.file_list.setCurrentRow(0)
        self.update_selection_summary()
        self.update_remove_button_state()
        self.sync_run_button_state()

    def update_remove_button_state(self) -> None:
        has_targets = bool(self.file_list.selectedItems()) or bool(self.checked_paths)
        self.remove_button.setDisabled(self.is_running or not has_targets)

    def mark_item_pending(self, item: QListWidgetItem) -> None:
        if self.is_running:
            return
        raw_path = item.data(Qt.ItemDataRole.UserRole)
        if not raw_path:
            return
        path = Path(raw_path)
        if self.file_statuses.get(path) in {"done", "failed", "stopped"}:
            self.file_statuses[path] = "pending"
            self.file_outputs.pop(path, None)
            self.refresh_file_list()
            self.update_output_path_hint()

    def pending_paths(self) -> list[Path]:
        runnable_states = {"pending", "queued", "running", "stopped"}
        return [
            path
            for path in self.input_paths
            if path in self.checked_paths
            if self.file_statuses.get(path, "pending") in runnable_states
        ]

    def run_or_stop(self) -> None:
        if self.is_running:
            self.stop_job()
        else:
            self.start_job()

    def stop_job(self) -> None:
        if not self.is_running or self.worker is None:
            return
        self.current_file_label.setText(self.text["status_stopping"])
        self.run_button.setText(self.text["stopping"])
        self.run_button.setDisabled(True)
        self.run_button.setToolTip(self.text["status_stopping"])
        self.worker.cancel()

    def set_running(self, running: bool) -> None:
        self.is_running = running
        self.current_file_label.setVisible(running)
        self.run_button.setText(self.text["stop"] if running else self.text["run"])
        self.run_button.setProperty("mode", "stop" if running else "run")
        self.run_button.style().unpolish(self.run_button)
        self.run_button.style().polish(self.run_button)
        self.pick_button.setDisabled(running)
        self.update_format_dropdown_state()
        self.watermark_checkbox.setDisabled(running)
        self.keep_videos_checkbox.setDisabled(running)
        self.update_image_quality_control()
        self.update_watermark_controls()
        self._sync_font_fix_button()
        self.update_remove_button_state()
        self.sync_run_button_state()
        self.refresh_log_shelf()
        if running:
            self.show_log_drawer()

    def start_job(self) -> None:
        if not self.input_paths:
            self.show_dialog(
                self.text["missing_file_title"], self.text["missing_file_body"]
            )
            return
        pending_paths = self.pending_paths()
        if not pending_paths:
            self.show_dialog(
                self.text["no_pending_title"], self.text["no_pending_body"]
            )
            return
        settings = self.current_settings(pending_paths)
        if not self.confirm_runtime_dependencies(settings, pending_paths):
            return

        has_pptx = any(p.suffix.lower() == ".pptx" for p in pending_paths)
        has_non_pptx = any(p.suffix.lower() in (".docx", ".pdf") for p in pending_paths)
        if has_pptx and has_non_pptx and settings.output_format == "pptx":
            dialog_parent = QApplication.activeWindow() or self
            dialog = QMessageBox(dialog_parent)
            dialog.setWindowTitle(self.windowTitle())
            text = self.text["mixed_queue_pptx_body"]
            dialog.setText(text)
            dialog.setStandardButtons(
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
            )
            if dialog.exec() != QMessageBox.StandardButton.Ok:
                return

        if (
            settings.watermark_enabled
            and settings.watermark_kind == "image"
            and settings.watermark_image_path is None
        ):
            self.show_dialog(
                self.text["watermark_image_missing_title"],
                self.text["watermark_image_missing_body"],
            )
            return
        self.event_log.setPlainText(
            self.text["job_started"].format(count=len(pending_paths))
        )
        self.current_file_label.setText(self.text["status_running"])
        self.progress_bar.setValue(0)
        self.set_running(True)
        self.worker_thread = QThread(self)
        for path in pending_paths:
            self.file_statuses[path] = "queued"
        self.refresh_file_list()
        self.worker = ExportWorker(pending_paths, settings, self.text)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.log.connect(self.append_log)
        self.worker.fileStarted.connect(self.on_file_started)
        self.worker.fileCompleted.connect(self.on_file_completed)
        self.worker.fileFailed.connect(self.on_file_failed)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def on_progress(self, percent: int, label: str) -> None:
        self.progress_bar.setValue(percent)
        self.current_file_label.setText(label)
        self.refresh_log_shelf()

    def append_log(self, message: str) -> None:
        LOGGER.info("%s", message)
        current = self.event_log.toPlainText().strip()
        if current == self.text["log_waiting"]:
            current = ""
        current = f"{current}\n{message}".strip() if current else message
        self.event_log.setPlainText(current)
        cursor = self.event_log.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.event_log.setTextCursor(cursor)
        self.refresh_log_shelf()
        if self.is_running:
            self.show_log_drawer(auto_hide=False)

    def on_file_started(self, path: object) -> None:
        source_path = Path(path)
        self.file_statuses[source_path] = "running"
        self.refresh_file_list()
        self.update_output_path_hint()
        self.current_file_label.setText(
            f"{self.text['current_processing']}: {source_path.name}"
        )

    def on_file_completed(
        self, path: object, output_path: object, size_bytes: int
    ) -> None:
        source_path = Path(path)
        output = Path(output_path)
        self.file_statuses[source_path] = "done"
        self.file_outputs[source_path] = output
        self.refresh_file_list()
        self.update_output_path_hint()
        self.append_log(f"[DONE] {output} ({format_file_size(size_bytes)})")

    def on_file_failed(self, path: object, message: str) -> None:
        source_path = Path(path)
        self.file_statuses[source_path] = "failed"
        self.refresh_file_list()
        self.update_output_path_hint()
        self.append_log(f"[ERR] {source_path.name}: {message}")
        self.show_log_drawer(auto_hide=False)

    def on_finished(self, results: list, failures: list, stopped: bool) -> None:
        total = len(results) + len(failures)
        if stopped:
            for path, state in list(self.file_statuses.items()):
                if state == "queued":
                    self.file_statuses[path] = "pending"
                elif state == "running":
                    self.file_statuses[path] = "stopped"
            self.current_file_label.setText(self.text["status_stopping"])
            self.append_log(self.text["job_stopped"])
        else:
            self.progress_bar.setValue(100)
            self.current_file_label.setText(
                self.text["done_summary"].format(
                    success=len(results), total=max(1, total)
                )
            )
        self.refresh_file_list()
        self.set_running(False)
        if self.worker_thread is not None:
            thread = self.worker_thread
            thread.quit()
            if not thread.wait(5000):
                while not thread.wait(250):
                    terminate_active_processes(grace_seconds=0.1)
            self.worker_thread = None
        self.worker = None
        if stopped:
            self.current_file_label.setText(self.text["job_stopped"])
        elif failures:
            self.current_file_label.setText(self.text["status_failed"])
            self.show_log_drawer(auto_hide=False)
        else:
            self.current_file_label.setText(self.text["status_done"])
            if not self.log_drawer.isHidden():
                self.log_drawer_timer.start(2500)
        self.refresh_log_shelf()
        self.update_output_path_hint()
        self.schedule_preview_refresh()

    def resizeEvent(self, event) -> None:  # noqa: N802, ANN001
        super().resizeEvent(event)
        self.update_preview_geometry()
        self.update_preview_display()
        self.position_log_drawer()

    def update_preview_geometry(self) -> None:
        if not hasattr(self, "preview_image_label"):
            return
        parent = getattr(self, "preview_scroll_area", None)
        parent_width = parent.viewport().width() if parent is not None else 0
        parent_height = parent.viewport().height() if parent is not None else 0
        preferred_width = PreviewCanvas.PREFERRED_WIDTH
        width = preferred_width
        page_count = 1 if self.preview_image_label_secondary.isHidden() else 2
        if parent_width > 0 and parent_height > 0:
            page_spacing = self.preview_pages_layout.spacing() * (page_count - 1)
            page_height = max(1, (parent_height - page_spacing - 16) // page_count)
            width_for_height = int(
                page_height * self.preview_image_label.current_aspect_ratio()
            )
            width = min(parent_width - 12, width_for_height)
        minimum_width = 300 if page_count == 1 else 220
        width = max(minimum_width, width)
        for preview in (
            self.preview_image_label,
            self.preview_image_label_secondary,
        ):
            preview.setFixedWidth(width)
            target_height = max(100, preview.heightForWidth(width))
            preview.setFixedHeight(target_height)

    def closeEvent(self, event) -> None:  # noqa: N802, ANN001
        self.preview_timer.stop()
        self.preview_dirty = False
        self.preview_request_id += 1
        if bool(getattr(self, "is_running", False)) and self.worker is not None:
            try:
                self.worker.cancel()
            except Exception:
                pass
        self.shutdown_thread(self.preview_thread)
        self.shutdown_thread(self.worker_thread)
        self.preview_thread = None
        self.preview_worker = None
        self.worker_thread = None
        self.worker = None
        self.cleanup_current_preview()
        self.cleanup_preview_sources()
        super().closeEvent(event)


def main() -> int:
    configure_app_logging()
    app = QApplication([])
    configure_ui_font(app)
    language = detect_language()
    app.setApplicationName("Doc Media Toolkit")
    app.setApplicationDisplayName(
        "文档媒体工具箱" if language == "zh" else "Doc Media Toolkit"
    )
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
