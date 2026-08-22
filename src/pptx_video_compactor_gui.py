#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path
from zipfile import ZipFile

from PySide6.QtCore import (
    QEvent,
    QFileInfo,
    QObject,
    QPoint,
    QRect,
    QSettings,
    QSize,
    QThread,
    QTimer,
    Qt,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QDragEnterEvent,
    QDragMoveEvent,
    QDropEvent,
    QIcon,
    QKeyEvent,
    QPainter,
    QPen,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFileIconProvider,
    QFrame,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QListWidgetItem,
    QListWidget,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pptx_video_compactor import (
    CancelledError,
    IMAGE_EXTENSIONS,
    IMAGE_QUALITY_RULES,
    PROFILE_QUALITY_RULES,
    VIDEO_EXTENSIONS,
    compact_input_path,
    compact_failed_assets_into_output,
    is_experimental_runtime,
    load_json_file,
)
from pptx_quality_audit import QualityAuditWorker
from pptx_tools.app_logging import configure_app_logging
from pptx_tools.language import detect_language as detect_system_language
from pptx_tools.image_manager import ImageProject
from pptx_tools.ui_theme import (
    SHARED_DIALOG_QSS,
    SHARED_MAIN_QSS,
    configure_ui_font,
    format_user_file_size,
    install_control_help,
)
from pptx_tools.video_manager import (
    VideoProject,
    normalize_library_category,
    sha256_file,
)

LOGGER = logging.getLogger("pptx_video_compactor_gui")

DOCUMENT_INPUT_EXTENSIONS = (".docx", ".docm", ".pdf", ".xlsx", ".xlsm")
SUPPORTED_COMPACTOR_INPUT_EXTENSIONS = (
    ".pptx",
    *IMAGE_EXTENSIONS,
    *VIDEO_EXTENSIONS,
    *DOCUMENT_INPUT_EXTENSIONS,
)


def app_settings() -> QSettings:
    name = (
        "Doc Media Toolkit Experimental"
        if is_experimental_runtime()
        else "Doc Media Toolkit"
    )
    return QSettings(name, name)


def resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    source_root = Path(__file__).resolve().parents[1]
    return source_root if (source_root / "assets").is_dir() else Path(sys.prefix)


def cleanup_stale_runtime_temp_dirs(max_age_hours: int = 24) -> None:
    cutoff = time.time() - max_age_hours * 3600
    temp_root = Path(tempfile.gettempdir())
    for pattern in ("pptx_incremental_*", "pptx_audit_*", "pptx_compact_*"):
        for path in temp_root.glob(pattern):
            try:
                if path.is_dir() and path.stat().st_mtime < cutoff:
                    shutil.rmtree(path, ignore_errors=True)
            except OSError:
                pass


def format_file_size(size_bytes: int) -> str:
    return format_user_file_size(size_bytes)


PRESET_OPTIONS = ("none", "high", "balanced", "aggressive")
IMAGE_PRESET_OPTIONS = ("none", "lossless", "high", "balanced", "aggressive")
VIDEO_PRESET_THRESHOLDS = {"high": 0.95, "balanced": 0.93, "aggressive": 0.90}
STRINGS = {
    "zh": {
        "window_title": "文档及媒体动态压缩",
        "eyebrow": "DOC MEDIA TOOLKIT",
        "title": "文档及媒体动态压缩",
        "subtitle": "压缩 PPTX 内嵌媒体，也支持独立图片、视频与 DOCX/PDF/XLSX 压缩",
        "settings_title": "压缩设置",
        "target_label": "目标大小 (MB)",
        "target_group_label": "文件大小",
        "target_placeholder": "留空按预设",
        "queue_title": "文件列表",
        "empty_queue": "拖入 PPTX / 图片 / 视频 / 文档",
        "empty_queue_hint": "支持拖放、多选与 Delete 删除；文档指 DOCX/PDF/XLSX",
        "estimate_label": "估算",
        "queued_marker": "排队",
        "running_marker": "处理中",
        "done_marker": "完成",
        "skip_marker": "跳过",
        "fail_marker": "失败",
        "profile_label": "视频预设",
        "image_profile_label": "图片预设",
        "profile_group_label": "视频",
        "image_profile_group_label": "图片",
        "profile_none": "不压缩",
        "image_profile_none": "不压缩",
        "image_profile_lossless": "PNG 无损",
        "image_profile_high": "高保真",
        "image_profile_balanced": "平衡",
        "image_profile_aggressive": "低体积",
        "profile_balanced": "平衡",
        "profile_high": "高保真",
        "profile_aggressive": "低体积",
        "pick_file": "添加文件",
        "remove_file": "移除选中",
        "help_button": "使用说明",
        "total_estimate": "共 {count} 个文件 · 预计输出 {size:.1f} MB",
        "total_estimate_ready": "共 {count} 个文件 · 预计输出 {size:.1f} MB",
        "current_processing": "正在处理",
        "no_current_file": "等待开始",
        "run": "开始压缩",
        "audit_button": "画质评估",
        "optimize_button": "提档优化",
        "auto_optimize": "自动提档",
        "target_gpu": "目标容量用 GPU",
        "target_gpu_tooltip": "默认关闭。开启后目标容量模式优先尝试 GPU，单个素材失败时回退 CPU。",
        "auto_optimize_tooltip": "压缩后自动评估；PPTX 低于阈值时自动提档并再次评估。",
        "archive_videos": "高清源视频入库",
        "archive_videos_tooltip": "可选：压缩前归档视频源。可限制为 1080p、仅转兼容 MP4，或保留原片。",
        "archive_off": "视频不入库",
        "archive_1080p": "1080p 高清入库",
        "archive_mp4": "兼容 MP4 入库",
        "archive_original": "原片入库",
        "archive_library_unset": "选择视频库…",
        "archive_library_off": "视频库：未启用",
        "archive_library_name": "视频库：{name}",
        "archive_library_tooltip": "压缩前的视频源和压缩后哈希关联将写入此视频库。点击可切换。",
        "archive_category": "入库分类",
        "archive_category_placeholder": "例如：示例项目/2026",
        "archive_category_invalid": "入库分类无效",
        "resource_archive": "资源归档",
        "resource_archive_expand": "展开设置",
        "resource_archive_collapse": "收起设置",
        "resource_archive_summary": "视频：{video}｜图片：{image}｜自动去重",
        "archive_images_label": "图片入库（可选）",
        "archive_images_off": "图片不入库",
        "archive_images_original": "原图入库",
        "archive_images_processed": "处理后图片入库",
        "image_library_unset": "选择图片库…",
        "image_library_off": "图片库：未启用",
        "image_library_name": "图片库：{name}",
        "image_library_tooltip": "压缩成功后写入图片库，按内容哈希自动去重。点击可切换。",
        "image_archive_category": "图片分类",
        "overwrite_original": "覆盖原 PPTX",
        "overwrite_original_tooltip": "仅允许在视频已入库、只压缩视频时使用；采用临时文件原子替换。",
        "overwrite_invalid_title": "不能覆盖原 PPTX",
        "overwrite_invalid_body": "覆盖原 PPTX 仅支持 PPTX 文件，并要求开启视频入库、图片预设设为“不压缩”、关闭自动评估。",
        "overwrite_confirm_title": "确认覆盖原 PPTX",
        "overwrite_confirm_body": "原 PPTX 将被压缩版替换。视频源会先入库，但图片等其他内容不会备份。是否继续？",
        "overwrite_confirm_button": "覆盖并压缩",
        "overwrite_cancel_button": "取消",
        "audit_button_tooltip": "完成至少一个压缩任务后，可对比原素材与压缩结果。",
        "optimize_button_tooltip": "画质评估发现低于阈值的 PPTX 素材后可用。",
        "video_threshold_label": "视频≥",
        "image_threshold_label": "图片≥",
        "video_threshold_tooltip": (
            "SSIM 质量阈值，范围 0.00-1.00，数值越高越严格。\n"
            "视频推荐 0.90；更高阈值会显著降低压缩收益。"
        ),
        "image_threshold_tooltip": (
            "SSIM 质量阈值，范围 0.00-1.00，数值越高越严格。\n"
            "图片推荐 0.99；重视体积时可使用 0.98。"
        ),
        "stop": "停止",
        "stopping": "正在停止",
        "pending_marker": "待处理",
        "details_title": "状态与日志",
        "initial_status": "等待开始。添加文件后可直接压缩。",
        "ready_status": "已选择 {count} 个文件，点击开始压缩。",
        "running_status": "正在压缩，请稍候。",
        "waiting_log": "等待开始。",
        "job_started": "任务开始：{count} 个文件。",
        "current_file": "当前文件",
        "output_path": "输出路径",
        "output_same_folder": "输出位置：源文件同目录。完整路径见日志。",
        "output_each_folder": "输出位置：各自源文件同目录。完整路径见日志。",
        "output_overwrite_hint": "输出位置：覆盖原 PPTX；视频源会先入库。",
        "output_done_hint": "输出完成。完整路径见日志。",
        "output_waiting_hint": "输出位置：源文件同目录。完整路径见日志。",
        "total_estimate_empty": "未选择文件",
        "choose_title": "选择 PPTX / 图片 / 视频 / 文档 文件",
        "choose_filter": "Supported Files (*.pptx *.png *.jpg *.jpeg *.jpe *.webp *.mp4 *.m4v *.mov *.wmv *.asf *.avi *.mpg *.mpeg *.mpe *.webm *.mkv *.ts *.m2ts *.3gp *.3g2 *.docx *.docm *.pdf *.xlsx *.xlsm)",
        "missing_file_title": "缺少文件",
        "missing_file_body": "先选择一个或多个 PPTX、图片、视频或文档（DOCX/PDF/XLSX）文件。",
        "no_media_profile_title": "未选择压缩内容",
        "no_media_profile_body": "视频预设和图片预设不能同时设为“不压缩”。",
        "invalid_size_title": "大小设置错误",
        "invalid_size_number": "目标大小必须是数字。",
        "invalid_size_positive": "目标大小必须大于 0。",
        "document_target_required_title": "需要目标大小",
        "document_target_required_body": "队列包含 DOCX/PDF/XLSX 文件，这类文件必须显式填写目标大小（MB）后才能开始压缩。",
        "document_image_profile_none_title": "图片预设不适用",
        "document_image_profile_none_body": "DOCX/PDF/XLSX 压缩仅处理嵌入图片，图片预设不能设为“不压缩”。",
        "batch_size_title": "批量目标大小",
        "batch_size_body": "已选择多个文件。建议留空目标大小，按预设为每个文件分别压缩。\n\n仍然使用同一个目标大小继续吗？",
        "target_too_large": "目标大小不小于源文件，已跳过。",
        "done_title": "压缩完成",
        "done_label": "完成",
        "output_file": "输出文件",
        "actual_size": "实际大小",
        "batch_done": "批量完成",
        "batch_summary": "成功 {success}/{total} 个文件。",
        "batch_failed_summary": "失败 {failed}/{total} 个文件，详情见日志。",
        "batch_skipped_summary": "跳过 {skipped} 个文件。",
        "completion_hint": "输出文件保存在源文件同目录，完整路径见任务日志。",
        "completion_single": "输出：{name} ({size:.1f} MB)",
        "processing_file": "正在处理 {index}/{total}: {name}",
        "failed_title": "压缩失败",
        "failed_status": "压缩失败。",
        "stopped_marker": "已停止",
        "stopped_status": "已停止。再次点击开始会继续处理未完成文件。",
        "no_pending_title": "没有待处理文件",
        "no_pending_body": "列表里的文件都已完成或跳过。双击对应条目可切回待处理。",
        "cancelled_log": "任务已停止。",
        "help_title": "使用说明",
        "help_body": (
            "1. 把一个或多个 PPTX、图片、视频或 DOCX/PDF/XLSX 文件拖入左侧列表，或点击添加文件。\n"
            "2. 目标大小留空时，视频和图片分别按各自预设压缩；两个预设都可以选择不压缩。DOCX/PDF/XLSX 必须显式填写目标大小，且图片预设不能为不压缩。\n"
            "3. 视频高保真偏 1080p 且保留帧率，平衡在 720/1080p 动态选择，极限压缩会更主动降低视频分辨率和帧率。\n"
            "4. 图片按原格式处理：JPEG/WebP 使用编码质量，高保真约 95，平衡约 85，低体积约 75；PNG 默认只做无损优化，不改格式。\n"
            "5. 填写目标大小时优先压缩视频；仍不够时再降低图片质量。图片低体积在目标压力下最多缩到 80% 像素尺寸。\n"
            "6. 勾选自动评估优化后，压缩结束会自动评估；PPTX 低于阈值时自动提档并复评。独立图片和视频只自动评估。\n"
            "7. 开始后按钮会变为停止；停止后再次开始会继续未完成文件。\n"
            "8. 默认优先使用 GPU 编码视频；不可用或失败时自动切回 CPU。只压图片时不需要 FFmpeg。\n"
            "9. 双击完成、跳过、失败或已停止条目，可切回待处理。\n"
            "10. 过程视频默认在临时目录，成功或停止会清理；普通错误才会在源文件旁保留诊断目录。"
        ),
        "optimize_prompt_restore": "💡 已提至高保真后仍有 {count} 个素材低于阈值；没有更高的压缩档位，将保留当前压缩结果。",
        "optimize_terminal_original": "⛔ 有 {count} 个素材在高保真压缩后仍低于阈值，已保留压缩结果并停止提档。建议降低 SSIM 阈值，或替换源素材。",
        "optimize_terminal_mixed": "⛔ 其中 {count} 个素材在高保真压缩后仍低于阈值，已保留压缩结果并停止提档。",
        "ok_button": "知道了",
        "continue_button": "继续",
        "preset_button": "留空按预设",
    },
    "en": {
        "window_title": "Document & Media Dynamic Compression",
        "eyebrow": "DOC MEDIA TOOLKIT",
        "title": "Document & Media Dynamic Compression",
        "subtitle": "Compress PPTX embedded media plus standalone images, videos, and DOCX/PDF/XLSX files",
        "settings_title": "Compression",
        "target_label": "Target size (MB)",
        "target_group_label": "Size (MB)",
        "target_placeholder": "Blank = presets",
        "queue_title": "Files",
        "empty_queue": "Drop PPTX, images, videos, or documents here",
        "empty_queue_hint": "Supports drag & drop, multi-select, and Delete; documents = DOCX/PDF/XLSX",
        "estimate_label": "Est.",
        "queued_marker": "Queued",
        "running_marker": "Processing",
        "done_marker": "Done",
        "skip_marker": "Skipped",
        "fail_marker": "Failed",
        "profile_label": "Video preset",
        "image_profile_label": "Image preset",
        "profile_group_label": "Video",
        "image_profile_group_label": "Image",
        "profile_none": "Off",
        "profile_balanced": "Balanced",
        "profile_high": "High",
        "profile_aggressive": "Low",
        "image_profile_none": "Off",
        "image_profile_lossless": "PNG lossless",
        "image_profile_high": "High",
        "image_profile_balanced": "Balanced",
        "image_profile_aggressive": "Low",
        "pick_file": "Add Files",
        "remove_file": "Remove",
        "help_button": "Help",
        "total_estimate": "{count} file(s) · est. {size:.1f} MB",
        "total_estimate_ready": "{count} file(s) · est. {size:.1f} MB",
        "current_processing": "Processing",
        "no_current_file": "Waiting",
        "run": "Run",
        "audit_button": "Audit",
        "optimize_button": "Optimize",
        "auto_optimize": "Audit && optimize after compression",
        "target_gpu": "GPU for target size",
        "target_gpu_tooltip": "Off by default. Target-size jobs try GPU first and fall back to CPU per asset.",
        "auto_optimize_tooltip": "Audit after compression; automatically optimize and re-audit PPTX files below the threshold.",
        "archive_videos": "Archive source videos",
        "archive_videos_tooltip": "Optional: archive sources before compression. Limit to 1080p, normalize to compatible MP4, or keep original bytes.",
        "archive_off": "Do not archive videos",
        "archive_1080p": "Archive high-quality 1080p",
        "archive_mp4": "Archive compatible MP4",
        "archive_original": "Archive original media",
        "archive_library_unset": "Choose library…",
        "archive_library_off": "Library: disabled",
        "archive_library_name": "Library: {name}",
        "archive_library_tooltip": "Source videos and compressed hash aliases will be recorded in this library. Click to change it.",
        "archive_category": "Library folder",
        "archive_category_placeholder": "e.g. Client A/2026",
        "archive_category_invalid": "Invalid library folder",
        "resource_archive": "Resource archive",
        "resource_archive_expand": "Show settings",
        "resource_archive_collapse": "Hide settings",
        "resource_archive_summary": "Video: {video} | Images: {image} | deduplicated",
        "archive_images_label": "Image archive (optional)",
        "archive_images_off": "Do not archive images",
        "archive_images_original": "Archive originals",
        "archive_images_processed": "Archive processed images",
        "image_library_unset": "Choose image library…",
        "image_library_off": "Image library: off",
        "image_library_name": "Image library: {name}",
        "image_library_tooltip": "Archive images after successful compression with hash deduplication. Click to switch.",
        "image_archive_category": "Image category",
        "overwrite_original": "Replace source PPTX",
        "overwrite_original_tooltip": "Available only when videos are archived and images are not compressed; replacement is atomic.",
        "overwrite_invalid_title": "Cannot replace source PPTX",
        "overwrite_invalid_body": "Replacing source files requires PPTX inputs, video archiving enabled, image compression set to Off, and automatic audit disabled.",
        "overwrite_confirm_title": "Replace source PPTX files?",
        "overwrite_confirm_body": "Each source PPTX will be replaced by its compressed copy. Video sources are archived first, but other content is not backed up. Continue?",
        "overwrite_confirm_button": "Replace and compress",
        "overwrite_cancel_button": "Cancel",
        "audit_button_tooltip": "Available after at least one compression finishes.",
        "optimize_button_tooltip": "Available when a PPTX quality audit finds assets below the threshold.",
        "video_threshold_label": "Video≥",
        "image_threshold_label": "Image≥",
        "video_threshold_tooltip": (
            "SSIM quality threshold, from 0.00 to 1.00. Higher is stricter.\n"
            "0.90 is recommended for video; higher values reduce compression savings."
        ),
        "image_threshold_tooltip": (
            "SSIM quality threshold, from 0.00 to 1.00. Higher is stricter.\n"
            "0.99 is recommended for images; use 0.98 when size matters more."
        ),
        "stop": "Stop",
        "stopping": "Stopping",
        "pending_marker": "Pending",
        "details_title": "Status & Log",
        "initial_status": "Waiting to start. Add files to compress.",
        "ready_status": "{count} files selected. Click Compress.",
        "running_status": "Compressing, please wait.",
        "waiting_log": "Waiting to start.",
        "job_started": "Job started: {count} files.",
        "current_file": "Current file",
        "output_path": "Output path",
        "output_same_folder": "Output: beside the source file. Full path is in the log.",
        "output_each_folder": "Output: beside each source file. Full paths are in the log.",
        "output_overwrite_hint": "Output: replace each source PPTX after archiving its videos.",
        "output_done_hint": "Output complete. Full path is in the log.",
        "output_waiting_hint": "Output is written beside the source file. Full paths appear in the log.",
        "total_estimate_empty": "No files selected",
        "choose_title": "Choose PPTX / Image / Video / Document Files",
        "choose_filter": "Supported Files (*.pptx *.png *.jpg *.jpeg *.jpe *.webp *.mp4 *.m4v *.mov *.wmv *.asf *.avi *.mpg *.mpeg *.mpe *.webm *.mkv *.ts *.m2ts *.3gp *.3g2 *.docx *.docm *.pdf *.xlsx *.xlsm)",
        "missing_file_title": "Missing file",
        "missing_file_body": "Choose one or more PPTX, image, video, or DOCX/PDF/XLSX files first.",
        "no_media_profile_title": "Nothing selected",
        "no_media_profile_body": "Video preset and image preset cannot both be Off.",
        "invalid_size_title": "Invalid size",
        "invalid_size_number": "Target size must be a number.",
        "invalid_size_positive": "Target size must be greater than 0.",
        "document_target_required_title": "Target size required",
        "document_target_required_body": "The queue contains DOCX/PDF/XLSX files. An explicit target size (MB) is required before compression.",
        "document_image_profile_none_title": "Image preset not applicable",
        "document_image_profile_none_body": "DOCX/PDF/XLSX compression only re-encodes embedded images; the image preset cannot be Off.",
        "batch_size_title": "Batch Target Size",
        "batch_size_body": "Multiple files are selected. Leaving target size blank is recommended so each file uses the selected preset.\n\nContinue with one shared target size anyway?",
        "target_too_large": "Target size is not smaller than source; skipped.",
        "done_title": "Compression Complete",
        "done_label": "Complete",
        "output_file": "Output file",
        "actual_size": "Actual size",
        "batch_done": "Batch Complete",
        "batch_summary": "Succeeded on {success}/{total} files.",
        "batch_failed_summary": "{failed}/{total} files failed. See the log for details.",
        "batch_skipped_summary": "{skipped} files skipped.",
        "completion_hint": "Output files are saved beside the source files. Full paths are in the task log.",
        "completion_single": "Output: {name} ({size:.1f} MB)",
        "processing_file": "Processing {index}/{total}: {name}",
        "failed_title": "Compression Failed",
        "failed_status": "Compression failed.",
        "stopped_marker": "Stopped",
        "stopped_status": "Stopped. Click Compress again to continue unfinished files.",
        "no_pending_title": "Nothing to process",
        "no_pending_body": "All files are complete or skipped. Double-click an item to mark it pending again.",
        "cancelled_log": "Job stopped.",
        "help_title": "How to use",
        "help_body": (
            "1. Drop one or more PPTX, image, video, or DOCX/PDF/XLSX files into the list on the left, or click Add Files.\n"
            "2. Leave target size blank to use the selected video and image presets. Either preset can be Off. DOCX/PDF/XLSX files require an explicit target size and the image preset cannot be Off.\n"
            "3. Video High favors 1080p and keeps FPS, Balanced chooses 720/1080p dynamically, and Aggressive lowers resolution and FPS more readily.\n"
            "4. Images keep their original format: JPEG/WebP use encoder quality around 95/85/75; PNG is optimized losslessly by default.\n"
            "5. With a target size, videos are compressed first. If more reduction is needed, image quality is lowered; Image Low may downsample to 80% pixel dimensions.\n"
            "6. Auto audit & optimize audits after compression, then optimizes and re-audits PPTX files below the threshold. Standalone media is audited only.\n"
            "7. GPU video encoding is preferred by default and falls back to CPU. Image-only runs do not require FFmpeg.\n"
            "8. Compress changes to Stop while running; starting again continues unfinished files.\n"
            "9. Double-click a done, skipped, failed, or stopped item to mark it pending."
        ),
        "optimize_prompt_restore": "💡 {count} asset(s) are still below the threshold after the High preset. There is no higher compressed preset, so the current result will be kept.",
        "optimize_terminal_original": "⛔ {count} asset(s) are still below the threshold at the High preset. The compressed result was kept; lower the SSIM threshold or replace the source media.",
        "optimize_terminal_mixed": "⛔ {count} asset(s) are still below the threshold at the High preset. The compressed results were kept and will not be retried automatically.",
        "ok_button": "OK",
        "continue_button": "Continue",
        "preset_button": "Use preset",
    },
}


def detect_language() -> str:
    return detect_system_language("PPTX_VIDEO_COMPACTOR_LANG")


def localize_progress_label(label: str, language: str) -> str:
    if language == "zh":
        return label
    exact = {
        "正在解析演示文稿": "Analyzing presentation",
        "正在提取内嵌视频": "Extracting embedded videos",
        "正在提取内嵌图片": "Extracting embedded images",
        "正在计算压缩计划": "Calculating compression plan",
        "正在重新打包 PPTX": "Repacking PPTX",
        "压缩完成": "Compression complete",
        "未找到可压缩媒体，已跳过": "No compressible media found; skipped",
    }
    if label in exact:
        return exact[label]
    extract_match = re.match(r"正在提取视频 (\d+)/(\d+)", label)
    if extract_match:
        return f"Extracting video {extract_match.group(1)}/{extract_match.group(2)}"
    encode_match = re.match(r"正在处理视频 (\d+)/(\d+): (.+)", label)
    if encode_match:
        return f"Processing video {encode_match.group(1)}/{encode_match.group(2)}: {encode_match.group(3)}"
    image_extract_match = re.match(r"正在提取图片 (\d+)/(\d+)", label)
    if image_extract_match:
        return f"Extracting image {image_extract_match.group(1)}/{image_extract_match.group(2)}"
    image_encode_match = re.match(r"正在处理图片 (\d+)/(\d+): (.+)", label)
    if image_encode_match:
        return f"Processing image {image_encode_match.group(1)}/{image_encode_match.group(2)}: {image_encode_match.group(3)}"
    return label


def build_namespace(
    input_pptx: Path,
    target_size_mb: float | None,
    profile: str,
    image_profile: str,
    output_path: Path | None = None,
    target_gpu_enabled: bool = False,
    video_ssim_threshold: float = 0.95,
    image_ssim_threshold: float = 0.99,
    quality_mode: str = "safe",
    standard_encoder_strategy: str = "auto",
) -> argparse.Namespace:
    return argparse.Namespace(
        input_pptx=input_pptx,
        target_size_mb=target_size_mb,
        config=None,
        profile=profile,
        image_profile=image_profile,
        output=output_path,
        video_output_dir=None,
        slide_render_width=1920,
        slide_render_height=1080,
        min_height=480,
        max_height=1080,
        overscan=1.2,
        reserve_mb=None,
        preset="medium",
        work_dir=None,
        keep_work_dir=False,
        keep_artifacts=False,
        dry_run=False,
        encoder=(
            "gpu"
            if target_size_mb is not None and target_gpu_enabled
            else "cpu"
            if target_size_mb is not None
            else standard_encoder_strategy
        ),
        video_ssim_threshold=video_ssim_threshold,
        image_ssim_threshold=image_ssim_threshold,
        quality_mode=quality_mode,
    )


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
                .endswith(SUPPORTED_COMPACTOR_INPUT_EXTENSIONS)
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
            if path.suffix.lower() in SUPPORTED_COMPACTOR_INPUT_EXTENSIONS:
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


class ElidedLabel(QLabel):
    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.full_text = text
        self.setText(text)

    def set_full_text(self, text: str) -> None:
        self.full_text = text
        self.update_elide()

    def resizeEvent(self, event) -> None:  # noqa: N802, ANN001
        super().resizeEvent(event)
        self.update_elide()

    def update_elide(self) -> None:
        width = max(20, self.width())
        self.setText(
            self.fontMetrics().elidedText(
                self.full_text, Qt.TextElideMode.ElideMiddle, width
            )
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
        center_x = self.width() - 18
        center_y = self.height() // 2 + 1
        painter.drawLine(center_x - 5, center_y - 4, center_x, center_y + 1)
        painter.drawLine(center_x, center_y + 1, center_x + 5, center_y - 4)


class CleanDoubleSpinBox(QDoubleSpinBox):
    def paintEvent(self, event) -> None:  # noqa: N802, ANN001
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        color = QColor("#94a3b8" if self.isEnabled() else "#64748b")
        pen = QPen(color, 1.6)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        center_x = self.width() - 11
        upper_y = self.height() // 2 - 6
        lower_y = self.height() // 2 + 7
        painter.drawLine(center_x - 3, upper_y + 2, center_x, upper_y - 1)
        painter.drawLine(center_x, upper_y - 1, center_x + 3, upper_y + 2)
        painter.drawLine(center_x - 3, lower_y - 2, center_x, lower_y + 1)
        painter.drawLine(center_x, lower_y + 1, center_x + 3, lower_y - 2)


class CompressionWorker(QObject):
    finished = Signal(list, list, bool, object)
    failed = Signal(str)
    status = Signal(str)
    progress = Signal(int, str)
    log = Signal(str)
    file_started = Signal(object)
    file_completed = Signal(object, bool, object, int, str, str)
    file_failed = Signal(object, str)

    def __init__(
        self,
        input_paths: list[Path],
        target_size_mb: float | None,
        profile: str,
        image_profile: str,
        language: str,
        text: dict[str, str],
        video_library_root: Path | None = None,
        archive_source_quality: str = "1080p",
        archive_category: str = "",
        image_library_root: Path | None = None,
        image_archive_mode: str = "off",
        image_archive_category: str = "",
        overwrite_original: bool = False,
        target_gpu_enabled: bool = False,
        video_ssim_threshold: float = 0.95,
        image_ssim_threshold: float = 0.99,
        quality_mode: str = "safe",
        standard_encoder_strategy: str = "auto",
    ) -> None:
        super().__init__()
        self.input_paths = input_paths
        self.target_size_mb = target_size_mb
        self.profile = profile
        self.image_profile = image_profile
        self.language = language
        self.text = text
        self.video_library_root = video_library_root
        self.archive_source_quality = archive_source_quality
        self.archive_category = archive_category
        self.image_library_root = image_library_root
        self.image_archive_mode = image_archive_mode
        self.image_archive_category = image_archive_category
        self.overwrite_original = overwrite_original
        self.target_gpu_enabled = target_gpu_enabled
        self.video_ssim_threshold = video_ssim_threshold
        self.image_ssim_threshold = image_ssim_threshold
        self.quality_mode = quality_mode
        self.standard_encoder_strategy = standard_encoder_strategy
        self.cancel_requested = False
        self.asset_pre_encoded_map: dict[Path, dict[str, str]] | None = None
        self.incremental_patch_map: dict[Path, dict[str, object]] | None = None
        self.total_bytes = sum(
            path.stat().st_size for path in self.input_paths if path.exists()
        )
        self.processed_bytes = 0
        self.file_progress = 0.0
        self.last_progress_percent = 0

    def cancel(self) -> None:
        self.cancel_requested = True

    def is_cancelled(self) -> bool:
        return self.cancel_requested

    def run(self) -> None:
        results: list[tuple[Path, Path, int, bool, str]] = []
        failures: list[tuple[Path, str]] = []
        total_files = len(self.input_paths)
        cancelled = False
        stopped_path: Path | None = None
        for file_index, input_pptx in enumerate(self.input_paths, start=1):
            if self.is_cancelled():
                cancelled = True
                break
            self.file_started.emit(input_pptx)
            self.file_progress = 0.0
            self.status.emit(
                self.text["processing_file"].format(
                    index=file_index, total=total_files, name=input_pptx.name
                )
            )
            file_size = input_pptx.stat().st_size if input_pptx.exists() else 0
            try:
                library = None
                archive_result = None
                registered_deck = None
                if (
                    self.video_library_root is not None
                    and input_pptx.suffix.lower() == ".pptx"
                ):
                    library = VideoProject.open(self.video_library_root)
                    archive_result = library.archive_and_register_pptx(
                        input_pptx,
                        source_quality=self.archive_source_quality,
                        category=self.archive_category,
                        progress_callback=self._log,
                        cancel_callback=self.is_cancelled,
                    )
                    self._log(
                        f"[LIBRARY] 新增 {archive_result['added']} 个，"
                        f"复用 {archive_result['reused']} 个视频，"
                        f"待确认高清候选 {archive_result.get('candidates_added', 0)} 个。"
                    )
                    registered_deck = archive_result["deck"]
                    if self.overwrite_original and not archive_result["media_families"]:
                        raise RuntimeError("PPTX 中没有可归档视频，拒绝覆盖原文件。")
                image_archive_before_overwrite = (
                    self.image_library_root is not None
                    and self.image_archive_mode == "original"
                    and self.overwrite_original
                    and input_pptx.suffix.lower() == ".pptx"
                )
                if image_archive_before_overwrite:
                    self._archive_images(input_pptx, required=True)
                patch_config = (
                    self.incremental_patch_map.get(input_pptx)
                    if self.incremental_patch_map
                    else None
                )
                if patch_config:
                    result = compact_failed_assets_into_output(
                        build_namespace(
                            input_pptx, None, self.profile, self.image_profile
                        ),
                        failed_media_paths=set(patch_config["failed_media_paths"]),
                        base_output_pptx=Path(patch_config["base_output_pptx"]),
                        report_path=Path(patch_config["report_path"]),
                        logger=self._log,
                        progress_callback=lambda done, total, label, fs=file_size: (
                            self._progress(
                                fs,
                                done,
                                total,
                                label,
                            )
                        ),
                        cancel_callback=self.is_cancelled,
                    )
                else:
                    result = compact_input_path(
                        build_namespace(
                            input_pptx,
                            self.target_size_mb,
                            self.profile,
                            self.image_profile,
                            input_pptx if self.overwrite_original else None,
                            self.target_gpu_enabled,
                            self.video_ssim_threshold,
                            self.image_ssim_threshold,
                            self.quality_mode,
                            self.standard_encoder_strategy,
                        ),
                        logger=self._log,
                        progress_callback=lambda done, total, label, fs=file_size: (
                            self._progress(
                                fs,
                                done,
                                total,
                                label,
                            )
                        ),
                        cancel_callback=self.is_cancelled,
                        asset_pre_encoded=(
                            self.asset_pre_encoded_map.get(input_pptx)
                            if self.asset_pre_encoded_map
                            else None
                        ),
                    )
                output_pptx = result["output_pptx"]
                if self.image_library_root is not None:
                    eligible = (
                        input_pptx.suffix.lower() == ".pptx"
                        or input_pptx.suffix.lower() in IMAGE_EXTENSIONS
                    )
                    if not eligible:
                        if input_pptx.suffix.lower() in VIDEO_EXTENSIONS:
                            self._log(
                                f"[IMAGE LIBRARY] 跳过 {input_pptx.name}："
                                "该文件不含可归档图片。"
                            )
                    elif not image_archive_before_overwrite:
                        image_source = (
                            input_pptx
                            if self.image_archive_mode == "original"
                            else output_pptx
                        )
                        self._archive_images(image_source, required=False)
                size = output_pptx.stat().st_size if output_pptx.exists() else 0
                skipped = bool(result.get("skipped"))
                reason = str(result.get("reason", ""))
                report_path = Path(result.get("report_path") or "")
                if (
                    library is not None
                    and archive_result is not None
                    and output_pptx.is_file()
                    and report_path.is_file()
                ):
                    aliases = library.register_compressed_pptx_hashes(
                        output_pptx,
                        report_path,
                        archive_result["media_families"],
                    )
                    self._log(f"[LIBRARY] 已登记 {aliases} 个压缩视频哈希。")
                    output_digest = sha256_file(output_pptx)
                    if self.overwrite_original and registered_deck is not None:
                        library.adopt_upgraded_deck_source(
                            registered_deck["id"], output_pptx
                        )
                        self._log("[LIBRARY] 已刷新覆盖后的 PPTX 关联。")
                    else:
                        library.register_optimized_output(
                            input_pptx, output_pptx, output_digest
                        )
                results.append((input_pptx, output_pptx, size, skipped, reason))
                digest = (
                    output_digest
                    if library is not None
                    and archive_result is not None
                    and output_pptx.is_file()
                    and report_path.is_file()
                    else sha256_file(output_pptx)
                    if input_pptx.suffix.lower() == ".pptx" and output_pptx.is_file()
                    else ""
                )
                self.file_completed.emit(
                    input_pptx, skipped, output_pptx, size, reason, digest
                )
            except CancelledError:
                cancelled = True
                stopped_path = input_pptx
                break
            except Exception as exc:  # pragma: no cover - GUI surface
                failures.append((input_pptx, str(exc)))
                self.file_failed.emit(input_pptx, str(exc))
            self.processed_bytes += file_size
        if total_files > 1 and self.input_paths:
            suffix = "_experimental" if is_experimental_runtime() else ""
            summary_path = self.input_paths[0].parent / (
                f"compression_batch_{time.strftime('%Y%m%d-%H%M%S')}{suffix}.md"
            )
            lines = ["# 批量压缩报告", "", f"- 文件数：{total_files}", ""]
            result_by_source = {Path(item[0]): item for item in results}
            failed_paths = {Path(item[0]) for item in failures}
            for source in self.input_paths:
                if source in failed_paths:
                    lines.append(f"- `{source.name}`：失败")
                elif source in result_by_source:
                    item = result_by_source[source]
                    status = "跳过" if item[3] else "完成"
                    lines.append(
                        f"- `{source.name}`：{status}，输出 {int(item[2]):,} bytes"
                    )
                else:
                    lines.append(f"- `{source.name}`：未处理")
            try:
                summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
                self._log(f"Batch report: {summary_path}")
            except OSError as exc:
                self._log(f"Batch report unavailable: {exc}")
        self.finished.emit(results, failures, cancelled, stopped_path)

    def _log(self, message: str) -> None:
        stripped = message.strip()
        if stripped:
            self.log.emit(stripped)

    def _archive_images(self, source: Path, *, required: bool) -> None:
        try:
            result = ImageProject.open(self.image_library_root).import_paths(
                [source], category=self.image_archive_category
            )
            self._log(
                f"[IMAGE LIBRARY] 新增 {result['added']} 张，"
                f"复用 {result['reused']} 张图片。"
            )
            for failure in result["failed"]:
                self._log(f"[IMAGE LIBRARY] 跳过：{failure['error']}")
            if required and result["failed"]:
                raise RuntimeError("原图未完整入库")
        except Exception as exc:
            if required:
                raise RuntimeError(f"原图入库失败，已停止覆盖：{exc}") from exc
            self._log(f"[IMAGE LIBRARY] 入库失败，压缩结果已保留：{exc}")

    def _progress(self, file_size: int, done: float, total: float, label: str) -> None:
        self.file_progress = max(
            self.file_progress, min(1.0, max(0.0, done / max(1.0, total)))
        )
        current_bytes = self.processed_bytes + (file_size * self.file_progress)
        percent = int((current_bytes / max(1, self.total_bytes)) * 100)
        percent = max(self.last_progress_percent, percent)
        self.last_progress_percent = percent
        self.progress.emit(percent, localize_progress_label(label, self.language))


class StyledDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        title: str,
        message: str,
        buttons: list[tuple[str, str, bool]],
    ) -> None:
        super().__init__(parent)
        self.result_key = buttons[0][1] if buttons else "ok"
        self.setModal(True)
        self.setWindowTitle(title)
        self.setMinimumWidth(440)
        self.setObjectName("styledDialog")

        # Start invisible to hide macOS window cascading jump
        self.setWindowOpacity(0.0)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Wait to center until layout is fully constructed
        title_label = QLabel(title)
        title_label.setObjectName("dialogTitle")
        layout.addWidget(title_label)

        body_label = QLabel(message)
        body_label.setObjectName("dialogBody")
        body_label.setWordWrap(True)
        layout.addWidget(body_label)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)
        button_row.addStretch(1)
        for label, key, primary in buttons:
            button = QPushButton(label)
            button.setObjectName("dialogPrimaryButton" if primary else "dialogButton")
            button.setFixedHeight(40)
            button.clicked.connect(
                lambda _checked=False, value=key: self.accept_result(value)
            )
            button_row.addWidget(button)
        layout.addLayout(button_row)

        self.setStyleSheet(
            """
            QDialog#styledDialog {
                background: #0b1017;
            }
            QLabel#dialogTitle {
                color: #f8fafc;
                font-size: 16px;
                font-weight: 600;
            }
            QLabel#dialogBody {
                color: #cbd5e1;
                font-size: 11px;
                line-height: 1.35;
            }
            QPushButton {
                background: #18212d;
                color: #ffffff;
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 6px 14px;
                min-width: 90px;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton#dialogPrimaryButton {
                background: #f97316;
                border: 1px solid #fb923c;
            }
            QPushButton:hover {
                background: #202b39;
            }
            QPushButton#dialogPrimaryButton:hover {
                background: #ea580c;
            }
            """
            + SHARED_DIALOG_QSS
        )

        # We use QTimer in showEvent instead of __init__ to beat macOS window cascading.

    def accept_result(self, key: str) -> None:
        self.result_key = key
        self.accept()

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


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        app = QApplication.instance()
        if app is not None:
            configure_ui_font(app)
        super().__init__()
        cleanup_stale_runtime_temp_dirs()
        self.language = detect_language()
        self.text = STRINGS[self.language]
        self.setWindowTitle(
            f"{self.text['window_title']} — Experimental"
            if is_experimental_runtime()
            else self.text["window_title"]
        )
        self.setMinimumSize(880, 560)
        self.resize(960, 620)

        self.input_paths: list[Path] = []
        self.file_statuses: dict[Path, str] = {}
        self.output_paths: dict[Path, Path] = {}
        self.failed_audits = {}
        self.forced_candidates: list[Path] = []
        self.session_report_paths: set[Path] = set()
        self.active_path: Path | None = None
        self.is_running = False
        self.worker_thread: QThread | None = None
        self.worker: CompressionWorker | None = None
        self.last_settings_signature: tuple[float | None, str, str, bool] | None = None
        self._suppress_settings_reset = False

        central = QWidget()
        central.setObjectName("central")
        self.content_widget = central
        central.installEventFilter(self)
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(14, 10, 14, 10)
        root.setSpacing(8)
        self.root_layout = root

        if is_experimental_runtime():
            experimental_banner = QLabel(
                "实验版 · 设置与正式版隔离 · 输出文件带 _experimental"
                if self.language == "zh"
                else "Experimental · isolated settings · _experimental outputs"
            )
            experimental_banner.setAlignment(Qt.AlignCenter)
            experimental_banner.setStyleSheet(
                "background:#7c2d12;color:#fff7ed;padding:5px 8px;font-weight:700;"
                "border-radius:6px;"
            )
            root.addWidget(experimental_banner)

        header = QFrame()
        header.setObjectName("headerCard")
        header.setFixedHeight(76)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 7, 16, 7)
        header_layout.setSpacing(10)

        title_stack_widget = QWidget(header)
        title_stack_widget.setObjectName("titleStack")
        title_stack = QVBoxLayout(title_stack_widget)
        title_stack.setContentsMargins(0, 0, 0, 0)
        title_stack.setSpacing(2)

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
        self.body_layout = body_row
        left_column = QVBoxLayout()
        left_column.setSpacing(8)
        left_column.setContentsMargins(0, 0, 0, 0)
        right_column = QVBoxLayout()
        right_column.setSpacing(5)
        right_column.setContentsMargins(0, 0, 0, 0)

        header_row_height = 30
        control_row_height = 30

        queue_card = QFrame()
        queue_card.setObjectName("queueCard")
        queue_card.setMinimumHeight(240)
        queue_card.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        queue_layout = QVBoxLayout(queue_card)
        queue_layout.setContentsMargins(10, 10, 10, 10)
        queue_layout.setSpacing(8)
        queue_header_widget = QWidget(queue_card)
        queue_header_widget.setFixedHeight(header_row_height)
        queue_header = QHBoxLayout(queue_header_widget)
        queue_header.setContentsMargins(0, 0, 0, 0)
        queue_header.setSpacing(10)
        queue_title = QLabel(self.text["queue_title"])
        queue_title.setObjectName("sectionTitle")
        queue_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        queue_title.setFixedHeight(header_row_height)
        self.total_estimate_label = QLabel(self.text["total_estimate_empty"])
        self.total_estimate_label.setObjectName("estimatePill")
        self.total_estimate_label.setAlignment(Qt.AlignCenter)
        self.total_estimate_label.setMinimumWidth(88 if self.language == "zh" else 104)
        self.total_estimate_label.setFixedHeight(header_row_height)
        queue_header.addWidget(queue_title, 1)
        queue_header.addWidget(self.total_estimate_label)
        queue_layout.addWidget(queue_header_widget)

        queue_buttons_widget = QWidget(queue_card)
        queue_buttons_widget.setMinimumHeight(control_row_height)
        queue_buttons = QHBoxLayout(queue_buttons_widget)
        queue_buttons.setContentsMargins(0, 0, 0, 0)
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
        queue_layout.addWidget(queue_buttons_widget)
        self.file_list = FileListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.set_empty_state(
            self.text["empty_queue"], self.text["empty_queue_hint"]
        )
        self.file_list.filesDropped.connect(self.set_files)
        self.file_list.deletePressed.connect(self.remove_selected_files)
        self.file_list.itemSelectionChanged.connect(self.on_file_selection_changed)
        self.file_list.itemDoubleClicked.connect(self.mark_item_pending)
        queue_layout.addWidget(self.file_list, 1)

        left_column.addWidget(queue_card)

        right_card = QFrame()
        self.settings_card = right_card
        right_card.setObjectName("rightCard")
        right_card.setMinimumHeight(180)
        right_card.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(10, 10, 10, 10)
        right_layout.setSpacing(8)

        settings_header_widget = QWidget(right_card)
        settings_header_widget.setFixedHeight(header_row_height)
        settings_header = QHBoxLayout()
        settings_header.setContentsMargins(0, 0, 0, 0)
        settings_header.setSpacing(6)
        settings_header.setAlignment(Qt.AlignVCenter)
        settings_header_widget.setLayout(settings_header)
        settings_title = QLabel(self.text["settings_title"])
        settings_title.setObjectName("sectionTitle")
        settings_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        settings_title.setFixedHeight(header_row_height)
        settings_header.addWidget(settings_title, 1)

        settings_controls_widget = QWidget(right_card)
        settings_controls_widget.setMinimumHeight(control_row_height)
        settings_controls = QGridLayout(settings_controls_widget)
        settings_controls.setContentsMargins(0, 0, 0, 0)
        settings_controls.setSpacing(7)

        assessment_row_primary_widget = QWidget(right_card)
        assessment_row_primary_widget.setMinimumHeight(control_row_height)
        assessment_controls = QHBoxLayout(assessment_row_primary_widget)
        assessment_controls.setContentsMargins(0, 0, 0, 0)
        assessment_controls.setSpacing(16)
        assessment_controls.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.assessment_row_primary_widget = assessment_row_primary_widget
        self.assessment_controls_layout = assessment_controls

        assessment_row_actions_widget = QWidget(right_card)
        assessment_row_actions_widget.setMinimumHeight(control_row_height)
        assessment_actions_layout = QHBoxLayout(assessment_row_actions_widget)
        assessment_actions_layout.setContentsMargins(0, 0, 0, 0)
        assessment_actions_layout.setSpacing(16)
        assessment_actions_layout.setAlignment(Qt.AlignVCenter)
        self.assessment_row_actions_widget = assessment_row_actions_widget
        self.assessment_actions_layout = assessment_actions_layout

        target_group_label = QLabel(self.text["target_group_label"])
        target_group_label.setObjectName("fieldLabel")
        target_group_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        target_group_label.setContentsMargins(8, 0, 0, 0)
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText(self.text["target_placeholder"])
        self.target_input.setMinimumWidth(108 if self.language == "zh" else 88)
        self.target_input.setMinimumHeight(control_row_height)
        self.target_input.setAlignment(Qt.AlignCenter)
        self.target_input.textChanged.connect(self.on_target_changed)

        target_group = QWidget()
        target_group.setObjectName("controlGroup")
        target_group.setProperty("plain", True)
        target_group_layout = QHBoxLayout(target_group)
        target_group_layout.setContentsMargins(0, 0, 0, 0)
        target_group_layout.setSpacing(4)
        target_group_label.setContentsMargins(0, 0, 0, 0)
        target_group_label.setFixedWidth(52 if self.language == "zh" else 68)
        target_group_layout.addWidget(target_group_label)
        target_group_layout.addWidget(self.target_input, 1)
        target_group.setMinimumWidth(166 if self.language == "zh" else 164)
        target_group.setMinimumHeight(control_row_height)
        target_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        profile_group_label = QLabel(self.text["profile_group_label"])
        profile_group_label.setObjectName("fieldLabel")
        profile_group_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        profile_group_label.setContentsMargins(8, 0, 0, 0)
        self.profile_select = CleanComboBox()
        self.profile_select.setView(QListView())
        self.profile_select.view().setObjectName("comboPopup")
        self.profile_select.setMaxVisibleItems(5)
        for profile in PRESET_OPTIONS:
            self.profile_select.addItem(self.text[f"profile_{profile}"], profile)
        self.profile_select.setCurrentIndex(PRESET_OPTIONS.index("high"))
        self.profile_select.setMinimumWidth(122 if self.language == "zh" else 108)
        self.profile_select.setMinimumHeight(control_row_height)
        self.profile_select.currentIndexChanged.connect(self.on_profile_changed)

        profile_group = QWidget()
        profile_group.setObjectName("controlGroup")
        profile_group.setProperty("plain", True)
        profile_group_layout = QHBoxLayout(profile_group)
        profile_group_layout.setContentsMargins(0, 0, 0, 0)
        profile_group_layout.setSpacing(4)
        profile_group_label.setContentsMargins(0, 0, 0, 0)
        profile_group_label.setFixedWidth(26 if self.language == "zh" else 36)
        profile_group_layout.addWidget(profile_group_label)
        profile_group_layout.addWidget(self.profile_select, 1)
        profile_group.setMinimumWidth(170 if self.language == "zh" else 150)
        profile_group.setMinimumHeight(control_row_height)
        profile_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        image_profile_group_label = QLabel(self.text["image_profile_group_label"])
        image_profile_group_label.setObjectName("fieldLabel")
        image_profile_group_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        image_profile_group_label.setContentsMargins(8, 0, 0, 0)
        self.image_profile_select = CleanComboBox()
        self.image_profile_select.setView(QListView())
        self.image_profile_select.view().setObjectName("comboPopup")
        self.image_profile_select.setMaxVisibleItems(5)
        for profile in IMAGE_PRESET_OPTIONS:
            self.image_profile_select.addItem(
                self.text[f"image_profile_{profile}"], profile
            )
        self.image_profile_select.setCurrentIndex(
            IMAGE_PRESET_OPTIONS.index("lossless")
        )
        self.image_profile_select.setMinimumWidth(122 if self.language == "zh" else 108)
        self.image_profile_select.setMinimumHeight(control_row_height)
        self.image_profile_select.currentIndexChanged.connect(self.on_profile_changed)

        image_profile_group = QWidget()
        image_profile_group.setObjectName("controlGroup")
        image_profile_group.setProperty("plain", True)
        image_profile_group_layout = QHBoxLayout(image_profile_group)
        image_profile_group_layout.setContentsMargins(0, 0, 0, 0)
        image_profile_group_layout.setSpacing(4)
        image_profile_group_label.setContentsMargins(0, 0, 0, 0)
        image_profile_group_label.setFixedWidth(26 if self.language == "zh" else 36)
        image_profile_group_layout.addWidget(image_profile_group_label)
        image_profile_group_layout.addWidget(self.image_profile_select, 1)
        image_profile_group.setMinimumWidth(170 if self.language == "zh" else 150)
        image_profile_group.setMinimumHeight(control_row_height)
        image_profile_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        settings = app_settings()

        self.video_threshold_label = QLabel(self.text["video_threshold_label"])
        self.video_threshold_label.setObjectName("fieldLabel")
        self.video_threshold_label.setToolTip(self.text["video_threshold_tooltip"])
        self.video_threshold_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.video_threshold_spinbox = CleanDoubleSpinBox()
        self.video_threshold_spinbox.setRange(0.0, 1.0)
        self.video_threshold_spinbox.setSingleStep(0.01)
        self.video_threshold_spinbox.setValue(
            settings.value("compression/video_ssim_threshold", 0.95, float)
        )
        self.video_threshold_spinbox.setDecimals(2)
        self.video_threshold_spinbox.setAlignment(Qt.AlignCenter)
        self._threshold_control_width = 80
        self.video_threshold_spinbox.setFixedWidth(self._threshold_control_width)
        self.video_threshold_spinbox.setMinimumHeight(control_row_height)
        self.video_threshold_spinbox.setToolTip(self.text["video_threshold_tooltip"])
        self._video_threshold_manual_override = (
            settings.value("compression/video_ssim_threshold_manual", False, bool)
            if settings.contains("compression/video_ssim_threshold_manual")
            else settings.contains("compression/video_ssim_threshold")
        )
        self.video_threshold_spinbox.valueChanged.connect(
            self.on_video_threshold_changed
        )

        self.image_threshold_label = QLabel(self.text["image_threshold_label"])
        self.image_threshold_label.setObjectName("fieldLabel")
        self.image_threshold_label.setToolTip(self.text["image_threshold_tooltip"])
        self.image_threshold_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.image_threshold_spinbox = CleanDoubleSpinBox()
        self.image_threshold_spinbox.setRange(0.0, 1.0)
        self.image_threshold_spinbox.setSingleStep(0.01)
        self.image_threshold_spinbox.setValue(
            settings.value("compression/image_ssim_threshold", 0.99, float)
        )
        self.image_threshold_spinbox.setDecimals(2)
        self.image_threshold_spinbox.setAlignment(Qt.AlignCenter)
        self.image_threshold_spinbox.setFixedWidth(self._threshold_control_width)
        self.image_threshold_spinbox.setMinimumHeight(control_row_height)
        self.image_threshold_spinbox.setToolTip(self.text["image_threshold_tooltip"])
        self.image_threshold_spinbox.valueChanged.connect(
            lambda value: settings.setValue("compression/image_ssim_threshold", value)
        )
        # Compatibility for callers that still use the old single-threshold field.
        self.threshold_spinbox = self.video_threshold_spinbox

        self.auto_optimize_checkbox = QCheckBox(self.text["auto_optimize"])
        self.auto_optimize_checkbox.setToolTip(self.text["auto_optimize_tooltip"])
        self.auto_optimize_checkbox.setMinimumHeight(control_row_height)
        self.auto_optimize_checkbox.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred
        )
        self.target_gpu_checkbox = QCheckBox(self.text["target_gpu"])
        self.target_gpu_checkbox.setToolTip(self.text["target_gpu_tooltip"])
        self.target_gpu_checkbox.setMinimumHeight(control_row_height)
        self.target_gpu_checkbox.setChecked(
            settings.value("compression/target_gpu", False, bool)
        )
        self.target_gpu_checkbox.toggled.connect(
            lambda checked: settings.setValue("compression/target_gpu", checked)
        )
        self.standard_encoder_select = CleanComboBox()
        self.standard_encoder_select.setMinimumHeight(control_row_height)
        self.standard_encoder_select.addItem(
            "自动硬件" if self.language == "zh" else "Auto hardware",
            "auto",
        )
        self.standard_encoder_select.addItem(
            "仅 CPU" if self.language == "zh" else "CPU only",
            "cpu",
        )
        self.standard_encoder_select.addItem(
            "优先 GPU" if self.language == "zh" else "Prefer GPU",
            "gpu",
        )
        saved_encoder_strategy = settings.value(
            "compression/standard_encoder_strategy", "auto", str
        )
        self.standard_encoder_select.setCurrentIndex(
            max(0, self.standard_encoder_select.findData(saved_encoder_strategy))
        )
        self.standard_encoder_select.currentIndexChanged.connect(
            lambda _index: settings.setValue(
                "compression/standard_encoder_strategy",
                self.standard_encoder_select.currentData(),
            )
        )

        self.archive_mode_select = CleanComboBox()
        self.archive_mode_select.addItem(self.text["archive_off"], "off")
        self.archive_mode_select.addItem(self.text["archive_1080p"], "1080p")
        self.archive_mode_select.addItem(self.text["archive_mp4"], "mp4")
        self.archive_mode_select.addItem(self.text["archive_original"], "original")
        self.archive_mode_select.setToolTip(self.text["archive_videos_tooltip"])
        self.archive_mode_select.setFixedWidth(150 if self.language == "zh" else 220)
        self.archive_mode_select.setMinimumHeight(control_row_height)
        saved_archive_quality = settings.value(
            "compression/archive_source_quality", "1080p", str
        )
        saved_archive_mode = (
            saved_archive_quality
            if settings.value("compression/archive_videos", False, bool)
            else "off"
        )
        self.archive_mode_select.setCurrentIndex(
            max(0, self.archive_mode_select.findData(saved_archive_mode))
        )

        def persist_archive_mode(_index: int) -> None:
            mode = str(self.archive_mode_select.currentData())
            settings.setValue("compression/archive_videos", mode != "off")
            if mode != "off":
                settings.setValue("compression/archive_source_quality", mode)
            elif hasattr(self, "overwrite_checkbox"):
                self.overwrite_checkbox.setChecked(False)

        self.archive_mode_select.currentIndexChanged.connect(persist_archive_mode)

        self.archive_category_input = QLineEdit()
        self.archive_category_input.setPlaceholderText(
            self.text["archive_category_placeholder"]
        )
        self.archive_category_input.setText(
            settings.value("compression/archive_category", "", str)
        )
        self.archive_category_input.setMinimumWidth(110)
        self.archive_category_input.textChanged.connect(
            lambda value: settings.setValue("compression/archive_category", value)
        )
        self.archive_library_button = QPushButton()
        self.archive_library_button.setObjectName("secondaryButton")
        self.archive_library_button.setMinimumHeight(control_row_height)
        self.archive_library_button.setFixedWidth(140 if self.language == "zh" else 180)
        self.archive_library_button.clicked.connect(self.choose_archive_library)
        self.image_archive_mode_select = CleanComboBox()
        self.image_archive_mode_select.addItem(self.text["archive_images_off"], "off")
        self.image_archive_mode_select.addItem(
            self.text["archive_images_original"], "original"
        )
        self.image_archive_mode_select.addItem(
            self.text["archive_images_processed"], "processed"
        )
        self.image_archive_mode_select.setFixedWidth(
            150 if self.language == "zh" else 220
        )
        self.image_archive_mode_select.setMinimumHeight(control_row_height)
        self.image_archive_mode_select.setCurrentIndex(
            max(
                0,
                self.image_archive_mode_select.findData(
                    settings.value("compression/image_archive_mode", "off", str)
                ),
            )
        )
        self.image_archive_mode_select.currentIndexChanged.connect(
            lambda _index: settings.setValue(
                "compression/image_archive_mode",
                self.image_archive_mode_select.currentData(),
            )
        )
        self.image_archive_mode_select.currentIndexChanged.connect(
            lambda _index: self.refresh_image_library_target()
        )
        self.image_archive_category_input = QLineEdit()
        self.image_archive_category_input.setPlaceholderText(
            self.text["archive_category_placeholder"]
        )
        self.image_archive_category_input.setText(
            settings.value("compression/image_archive_category", "", str)
        )
        self.image_archive_category_input.setMinimumWidth(110)
        self.image_archive_category_input.textChanged.connect(
            lambda value: settings.setValue("compression/image_archive_category", value)
        )
        self.image_library_button = QPushButton()
        self.image_library_button.setObjectName("secondaryButton")
        self.image_library_button.setMinimumHeight(control_row_height)
        self.image_library_button.setFixedWidth(140 if self.language == "zh" else 180)
        self.image_library_button.clicked.connect(self.choose_image_library)
        self.overwrite_checkbox = QCheckBox(self.text["overwrite_original"])
        self.overwrite_checkbox.setToolTip(self.text["overwrite_original_tooltip"])
        self.overwrite_checkbox.setChecked(
            self.archive_mode_select.currentData() != "off"
            and settings.value("compression/overwrite_original", False, bool)
        )
        self.overwrite_checkbox.setEnabled(
            self.archive_mode_select.currentData() != "off"
        )
        self.overwrite_checkbox.toggled.connect(
            lambda checked: settings.setValue("compression/overwrite_original", checked)
        )
        self.overwrite_checkbox.toggled.connect(self.on_target_changed)
        self.archive_category_input.setEnabled(
            self.archive_mode_select.currentData() != "off"
        )
        self.archive_mode_select.currentIndexChanged.connect(
            lambda _index: self.archive_category_input.setEnabled(
                self.archive_mode_select.currentData() != "off"
            )
        )
        self.archive_mode_select.currentIndexChanged.connect(
            lambda _index: self.overwrite_checkbox.setEnabled(
                self.archive_mode_select.currentData() != "off"
            )
        )
        self.archive_mode_select.currentIndexChanged.connect(
            lambda _index: self.refresh_archive_library_target()
        )
        self.refresh_archive_library_target()

        self.audit_button = QPushButton(self.text.get("audit_button", "画质评估"))
        self.audit_button.setObjectName("secondaryButton")
        self.audit_button.setMinimumHeight(control_row_height)
        self.audit_button.setDisabled(True)
        self.audit_button.setFixedWidth(72 if self.language == "zh" else 70)
        self.audit_button.setToolTip(self.text["audit_button_tooltip"])
        self.audit_button.clicked.connect(lambda: self.run_audit())

        self.optimize_button = QPushButton(self.text.get("optimize_button", "提档优化"))
        self.optimize_button.setObjectName("secondaryButton")
        self.optimize_button.setMinimumHeight(control_row_height)
        self.optimize_button.setDisabled(True)
        self.optimize_button.setFixedWidth(72)
        self.optimize_button.setToolTip(self.text["optimize_button_tooltip"])
        self.optimize_button.clicked.connect(self.on_optimize_clicked)

        self.forced_button = QPushButton(
            "尝试强制版" if self.language == "zh" else "Try forced"
        )
        self.forced_button.setObjectName("secondaryButton")
        self.forced_button.setMinimumHeight(control_row_height)
        self.forced_button.setFixedWidth(96 if self.language == "zh" else 80)
        self.forced_button.setVisible(False)
        self.forced_button.setToolTip(
            "仅对未达到目标容量的安全版可用；保留安全版并执行二次确认。"
            if self.language == "zh"
            else "Only for safe outputs above target; keeps the safe output and asks again."
        )
        self.forced_button.clicked.connect(self.run_forced_candidates)

        self.run_button = QPushButton(self.text["run"])
        self.run_button.setObjectName("primaryButton")
        self.run_button.setProperty("mode", "run")
        self.run_button.setFixedWidth(100 if self.language == "zh" else 94)
        self.run_button.setMinimumHeight(control_row_height)
        self.run_button.setDisabled(True)
        self.run_button.clicked.connect(self.run_or_stop)

        settings_header.addWidget(self.run_button)
        right_layout.addWidget(settings_header_widget)

        assessment_controls.addWidget(self.auto_optimize_checkbox)
        assessment_controls.addWidget(self.target_gpu_checkbox)
        assessment_controls.addWidget(self.standard_encoder_select)
        assessment_controls.addWidget(self.video_threshold_label)
        assessment_controls.addWidget(self.video_threshold_spinbox)
        assessment_controls.addWidget(self.image_threshold_label)
        assessment_controls.addWidget(self.image_threshold_spinbox)
        assessment_actions_layout.addWidget(self.audit_button)
        assessment_actions_layout.addWidget(self.optimize_button)
        assessment_actions_layout.addWidget(self.forced_button)
        assessment_actions_layout.addStretch(1)
        self.assessment_primary_controls = (
            self.auto_optimize_checkbox,
            self.target_gpu_checkbox,
            self.standard_encoder_select,
            self.video_threshold_label,
            self.video_threshold_spinbox,
            self.image_threshold_label,
            self.image_threshold_spinbox,
        )
        # Native Qt styles can report a one-pixel taller minimum hint on one
        # architecture. Give every control the same measured row height so
        # labels, checkboxes, combos, and spin boxes share one visual center.
        assessment_control_height = max(
            control_row_height,
            *(
                control.minimumSizeHint().height()
                for control in self.assessment_primary_controls
            ),
        )
        assessment_row_primary_widget.setFixedHeight(assessment_control_height)
        for control in self.assessment_primary_controls:
            control.setFixedHeight(assessment_control_height)
        self.assessment_action_buttons = (
            self.audit_button,
            self.optimize_button,
            self.forced_button,
        )
        settings_controls.addWidget(target_group, 0, 0)
        settings_controls.addWidget(profile_group, 0, 1)
        settings_controls.addWidget(image_profile_group, 0, 2)
        settings_controls.setColumnStretch(0, 1)
        settings_controls.setColumnStretch(1, 1)
        settings_controls.setColumnStretch(2, 1)
        self.settings_controls_widget = settings_controls_widget
        self.settings_controls_layout = settings_controls
        self.settings_control_groups = (
            target_group,
            profile_group,
            image_profile_group,
        )
        right_layout.addWidget(settings_controls_widget)

        archive_summary_widget = QWidget(right_card)
        archive_summary_widget.setMinimumHeight(control_row_height)
        archive_summary_layout = QHBoxLayout(archive_summary_widget)
        archive_summary_layout.setContentsMargins(0, 0, 0, 0)
        archive_summary_layout.setSpacing(8)
        archive_summary_title = QLabel(self.text["resource_archive"])
        archive_summary_title.setObjectName("sectionTitle")
        self.archive_summary_label = QLabel()
        self.archive_summary_label.setObjectName("fieldLabel")
        self.archive_summary_label.setMinimumWidth(0)
        self.archive_summary_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed
        )
        self.archive_summary_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.archive_disclosure_button = QPushButton(
            self.text["resource_archive_expand"]
        )
        self.archive_disclosure_button.setObjectName("disclosureButton")
        self.archive_disclosure_button.setCheckable(True)
        self.archive_disclosure_button.setMinimumHeight(control_row_height)
        self.archive_disclosure_button.clicked.connect(self.toggle_archive_settings)
        archive_summary_layout.addWidget(archive_summary_title)
        archive_summary_layout.addWidget(self.archive_summary_label, 1)
        archive_summary_layout.addWidget(self.archive_disclosure_button)
        right_layout.addWidget(archive_summary_widget)

        self.archive_settings_panel = QWidget(right_card)
        archive_panel_layout = QVBoxLayout(self.archive_settings_panel)
        archive_panel_layout.setContentsMargins(0, 0, 0, 0)
        archive_panel_layout.setSpacing(8)

        archive_category_widget = QWidget(self.archive_settings_panel)
        archive_category_widget.setMinimumHeight(control_row_height)
        archive_category_layout = QHBoxLayout(archive_category_widget)
        archive_category_layout.setContentsMargins(0, 0, 0, 0)
        archive_category_layout.setSpacing(7)
        self.archive_mode_label = QLabel(
            "视频入库（可选）" if self.language == "zh" else "Video archive (optional)"
        )
        self.archive_mode_label.setObjectName("fieldLabel")
        self.archive_mode_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        archive_category_label = QLabel(self.text["archive_category"])
        archive_category_label.setObjectName("fieldLabel")
        archive_category_layout.addWidget(self.archive_mode_label)
        archive_category_layout.addWidget(self.archive_mode_select)
        archive_category_layout.addWidget(self.archive_library_button)
        archive_category_layout.addWidget(archive_category_label)
        archive_category_layout.addWidget(self.archive_category_input, 1)
        archive_category_layout.addWidget(self.overwrite_checkbox)
        archive_panel_layout.addWidget(archive_category_widget)

        image_archive_widget = QWidget(self.archive_settings_panel)
        image_archive_widget.setMinimumHeight(control_row_height)
        image_archive_layout = QHBoxLayout(image_archive_widget)
        image_archive_layout.setContentsMargins(0, 0, 0, 0)
        image_archive_layout.setSpacing(7)
        image_archive_label = QLabel(self.text["archive_images_label"])
        image_archive_label.setObjectName("fieldLabel")
        image_archive_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        image_category_label = QLabel(self.text["image_archive_category"])
        image_category_label.setObjectName("fieldLabel")
        image_archive_layout.addWidget(image_archive_label)
        image_archive_layout.addWidget(self.image_archive_mode_select)
        image_archive_layout.addWidget(self.image_library_button)
        image_archive_layout.addWidget(image_category_label)
        image_archive_layout.addWidget(self.image_archive_category_input, 1)
        archive_panel_layout.addWidget(image_archive_widget)
        self.archive_settings_panel.hide()
        right_layout.addWidget(self.archive_settings_panel)
        self.refresh_image_library_target()
        self.refresh_archive_summary()

        details_layout = QVBoxLayout()
        details_layout.setContentsMargins(0, 0, 0, 0)
        details_layout.setSpacing(8)

        details_header = QHBoxLayout()
        details_header.setContentsMargins(0, 0, 0, 0)
        details_header.setSpacing(8)
        details_header.setAlignment(Qt.AlignVCenter)
        results_title = QLabel("压缩结果" if self.language == "zh" else "Results")
        results_title.setObjectName("sectionTitle")
        details_header.addWidget(results_title)
        details_header.addStretch(1)
        details_header.addWidget(self.assessment_row_actions_widget)
        self.details_header = details_header

        details_controls_layout = QVBoxLayout()
        details_controls_layout.setContentsMargins(0, 0, 0, 0)
        details_controls_layout.setSpacing(4)
        details_controls_layout.addWidget(self.assessment_row_primary_widget)
        self.details_controls_layout = details_controls_layout
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")

        self.output_hint_label = QLabel(self.text["output_waiting_hint"])
        self.output_hint_label.setObjectName("outputHint")
        self.output_hint_label.setWordWrap(True)
        self.output_hint_label.setFixedHeight(22)

        self.current_file_label = QLabel(self.text["initial_status"])
        self.current_file_label.setObjectName("currentFile")
        self.current_file_label.setWordWrap(True)
        self.current_file_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.current_file_label.setFixedHeight(24)
        self.current_file_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )

        self.results_tree = QTreeWidget()
        self.results_tree.setObjectName("resultsTree")
        self.results_tree.setHeaderLabels(
            ["文件", "类型", "原大小", "输出大小", "节省", "状态", "输出位置"]
            if self.language == "zh"
            else [
                "File",
                "Type",
                "Original",
                "Output",
                "Saved",
                "Status",
                "Output location",
            ]
        )
        self.results_tree.setRootIsDecorated(False)
        self.results_tree.setAlternatingRowColors(True)
        self.results_tree.setColumnWidth(0, 210)
        self.results_tree.setColumnWidth(1, 64)
        self.results_tree.setColumnWidth(2, 84)
        self.results_tree.setColumnWidth(3, 84)
        self.results_tree.setColumnWidth(4, 96)
        self.results_tree.setColumnWidth(5, 90)
        self.results_tree.header().setSectionResizeMode(
            6, QHeaderView.ResizeMode.Stretch
        )
        self.event_log = QTextEdit()
        self.event_log.setReadOnly(True)
        self.event_log.setObjectName("eventLog")
        self.event_log.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.event_log.setPlainText(self.text["waiting_log"])
        self.event_log.setMinimumWidth(400)
        self.event_log.setMinimumHeight(132)
        self.event_log.setMaximumHeight(150)
        details_layout.addLayout(details_header)
        details_layout.addLayout(details_controls_layout)
        details_layout.addWidget(self.results_tree, 1)
        right_layout.addLayout(details_layout, 1)
        right_column.addWidget(right_card, 1)
        left_column.setStretch(0, 1)
        right_column.setStretch(0, 1)
        left_pane = QWidget()
        left_pane.setLayout(left_column)
        left_pane.setMinimumWidth(320)
        left_pane.setMaximumWidth(320)
        right_pane = QWidget()
        right_pane.setLayout(right_column)
        right_pane.setMinimumWidth(0)
        self.left_pane = left_pane
        self.right_pane = right_pane
        body_row.addWidget(left_pane, 0)
        body_row.addWidget(right_pane, 1)
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
        log_drawer_layout.setContentsMargins(10, 8, 10, 10)
        log_drawer_layout.setSpacing(4)
        log_drawer_header = QHBoxLayout()
        log_drawer_header.setContentsMargins(0, 0, 0, 0)
        log_drawer_title = QLabel(self.text["details_title"])
        log_drawer_title.setObjectName("logDrawerTitle")
        log_drawer_header.addWidget(log_drawer_title)
        log_drawer_header.addStretch(1)
        close_log_drawer = QPushButton("收起" if self.language == "zh" else "Collapse")
        close_log_drawer.setObjectName("disclosureButton")
        close_log_drawer.setFixedHeight(24)
        close_log_drawer.clicked.connect(self.hide_log_drawer)
        log_drawer_header.addWidget(close_log_drawer)
        log_drawer_layout.addLayout(log_drawer_header)
        log_drawer_layout.addWidget(self.progress_bar)
        log_drawer_layout.addWidget(self.current_file_label)
        log_drawer_layout.addWidget(self.output_hint_label)
        log_drawer_layout.addWidget(self.event_log, 1)
        self.log_drawer.hide()
        self.log_drawer_timer = QTimer(self)
        self.log_drawer_timer.setSingleShot(True)
        self.log_drawer_timer.timeout.connect(self.hide_log_drawer_if_idle)
        self.log_hover_timer = QTimer(self)
        self.log_hover_timer.setSingleShot(True)
        self.log_hover_timer.setInterval(1000)
        self.log_hover_timer.timeout.connect(self.show_log_drawer_after_hover)

        self.setStyleSheet(
            """
            QMainWindow, QWidget#central { background: #0b1017; }
            QLabel {
                color: #dbe4f0;
                font-size: 13px;
            }
            QFrame#headerCard, QFrame#sideCard, QFrame#detailsCard, QFrame#queueCard, QFrame#rightCard {
                background: #121a24;
                border: 1px solid #273244;
                border-radius: 10px;
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
                font-weight: 500;
            }
            QLabel#estimatePill {
                color: #f8fafc;
                background: #18212d;
                border: 1px solid #334155;
                border-radius: 9px;
                padding: 2px 10px;
                font-size: 13px;
                font-weight: 500;
            }
            QLabel#currentFile {
                color: #cbd5e1;
                background: #0d1520;
                border: 1px solid #1e293b;
                border-radius: 6px;
                padding: 2px 8px;
                font-size: 13px;
            }
            QWidget#controlGroup {
                background: rgba(15, 23, 32, 0.62);
                border: 1px solid rgba(59, 74, 95, 0.88);
                border-radius: 10px;
                padding: 0 8px;
            }
            QWidget#controlGroup[plain="true"] {
                background: transparent;
                border: 0;
                border-radius: 0;
                padding: 0;
            }
            QLineEdit, QComboBox, QDoubleSpinBox {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 0 28px 0 10px;
                selection-background-color: #f97316;
                selection-color: #ffffff;
                font-size: 13px;
                font-weight: 500;
            }
            QLineEdit {
                padding: 0 10px;
            }
            QDoubleSpinBox {
                padding: 0 20px 0 8px;
            }
            QLineEdit:disabled, QComboBox:disabled, QDoubleSpinBox:disabled {
                background: #0b111a;
                color: #64748b;
                border: 1px solid #273244;
            }
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus,
            QPushButton:focus, QCheckBox:focus {
                border: 1px solid #fb923c;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                border: 0;
                background: transparent;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0;
                height: 0;
            }
            QDoubleSpinBox::up-button, QDoubleSpinBox::down-button {
                background: #18212d;
                border: 0;
                width: 16px;
                margin: 2px 2px 2px 0;
                border-radius: 4px;
            }
            QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover {
                background: #243244;
            }
            QCheckBox {
                color: #cbd5e1;
                spacing: 6px;
                font-size: 13px;
                font-weight: 500;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
                background: #0f1720;
                border: 1px solid #475569;
                border-radius: 4px;
            }
            QCheckBox::indicator:checked {
                background: #f97316;
                border-color: #fb923c;
            }
            QListView#comboPopup {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 3px;
                outline: 0;
                selection-background-color: #243244;
                selection-color: #f8fafc;
                font-size: 13px;
                font-weight: 500;
            }
            QListView#comboPopup::item {
                min-height: 22px;
                padding: 4px 8px;
                border-radius: 5px;
            }
            QListView#comboPopup::item:selected {
                background: #243244;
                color: #f8fafc;
            }
            QPushButton {
                background: #18212d;
                color: #ffffff;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 4px 10px;
                min-width: 90px;
                font-size: 13px;
                font-weight: 500;
            }
            QPushButton::menu-indicator {
                image: none;
                width: 0px;
                height: 0px;
                margin: 0px;
                padding: 0px;
            }
            QPushButton:hover {
                background: #202b39;
            }
            QPushButton:disabled {
                background: #334155;
                color: #94a3b8;
            }
            QPushButton#primaryButton {
                background: #f97316;
                border: 1px solid #fb923c;
                color: #ffffff;
                padding: 4px 0;
                min-width: 98px;
                max-width: 98px;
            }
            QPushButton#primaryButton:disabled {
                background: #334155;
                border: 1px solid #475569;
                color: #94a3b8;
            }
            QPushButton#primaryButton:hover {
                background: #ea580c;
            }
            QPushButton#primaryButton[mode="stop"] {
                background: #dc2626;
                border: 1px solid #ef4444;
            }
            QPushButton#primaryButton[mode="stop"]:hover {
                background: #b91c1c;
            }
            QPushButton#secondaryButton {
                min-width: 0;
                padding: 5px 10px;
            }
            QPushButton#disclosureButton {
                background: #18212d;
                color: #94a3b8;
                border: 1px solid #334155;
                border-radius: 6px;
                min-width: 0;
                padding: 0 8px;
                font-size: 13px;
                font-weight: 500;
            }
            QPushButton#logShelf {
                background: #0f1720;
                color: #94a3b8;
                border: 1px solid #273244;
                border-radius: 6px;
                min-width: 0;
                padding: 0 10px;
                text-align: left;
                font-size: 13px;
                font-weight: 500;
            }
            QPushButton#logShelf:hover { color: #f8fafc; background: #18212d; }
            QFrame#logDrawer {
                background: #0e1622;
                border: 1px solid #273244;
                border-radius: 8px;
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
            QProgressBar {
                background: #0f1720;
                color: #f8fafc;
                border: 1px solid #273244;
                border-radius: 6px;
                text-align: center;
                min-height: 14px;
            }
            QProgressBar::chunk {
                background: #f97316;
                border-radius: 5px;
            }
            QLabel#outputHint {
                color: #94a3b8;
                background: #0d1520;
                border: 1px solid #1e293b;
                border-radius: 6px;
                padding: 2px 8px;
                font-size: 13px;
            }
            QTextEdit#eventLog {
                background: #06090e;
                color: #38bdf8;
                border: 1px solid #1e293b;
                border-radius: 6px;
                padding: 6px 8px;
                font-family: "PingFang SC", "Hiragino Sans GB", "SF Pro Text", "Microsoft YaHei", Menlo, Monaco, Consolas, monospace;
                font-size: 13px;
                line-height: 1.35;
            }
            QListWidget#fileList {
                background: #0f1720;
                color: #cbd5e1;
                border: 1px solid #273244;
                border-radius: 8px;
                padding: 4px;
                font-size: 13px;
                outline: 0;
            }
            QListWidget#fileList::item {
                padding: 4px;
                border-radius: 5px;
            }
            QListWidget#fileList::item:selected {
                background: #1e293b;
                color: #f8fafc;
            }
            QTreeWidget#resultsTree {
                background: #0f1720;
                alternate-background-color: #111b28;
                color: #cbd5e1;
                border: 1px solid #273244;
                border-radius: 8px;
                font-size: 13px;
                outline: 0;
            }
            QTreeWidget#resultsTree::item {
                min-height: 30px;
                padding: 3px 5px;
            }
            QTreeWidget#resultsTree::item:selected {
                background: #12385f;
                color: #f8fafc;
            }
            QTreeWidget#resultsTree QHeaderView::section {
                background: #18212d;
                color: #dbe4f0;
                border: 0;
                padding: 6px 5px;
                font-size: 13px;
                font-weight: 600;
            }
            QScrollBar:vertical {
                background: #0f1720;
                border: 0;
                width: 8px;
                margin: 2px 2px 2px 2px;
                border-radius: 4px;
            }
            QScrollBar:horizontal {
                background: #0f1720;
                border: 0;
                height: 8px;
                margin: 2px 2px 2px 2px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: #334155;
                border-radius: 4px;
                min-height: 24px;
                min-width: 24px;
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
            QWidget#fileItem {
                background: #09131f;
                border: 1px solid #1d3043;
                border-radius: 6px;
            }
            QWidget#fileItem[selected="true"] {
                background: #1d2939;
                border: 1px solid #3b4a5f;
                border-radius: 6px;
            }
            QWidget#fileItem[state="running"] {
                background: #17202c;
                border-radius: 6px;
            }
            QWidget#fileItem[state="stopped"] {
                background: #151c27;
                border-radius: 6px;
            }
            QLabel#fileName {
                color: #cbd5e1;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#fileMeta {
                color: #94a3b8;
                font-size: 11px;
            }
            QLabel#fileEstimate {
                color: #94a3b8;
                background: transparent;
                border: 0;
                padding: 0;
                font-size: 11px;
                font-weight: 400;
            }
            QLabel#fileStatus {
                color: #94a3b8;
                font-size: 11px;
            }
            QLabel#fileState {
                color: #94a3b8;
                background: transparent;
                font-size: 11px;
                font-weight: 600;
            }
            QLabel#fileState[state="queued"] { color: #94a3b8; }
            QLabel#fileState[state="pending"] { color: #94a3b8; }
            QLabel#fileState[state="running"] { color: #f97316; }
            QLabel#fileState[state="done"] { color: #22c55e; }
            QLabel#fileState[state="skipped"] { color: #94a3b8; }
            QLabel#fileState[state="failed"] { color: #ef4444; }
            QLabel#fileState[state="stopped"] { color: #f97316; }
            QLabel, QPushButton, QLineEdit, QComboBox, QDoubleSpinBox,
            QCheckBox, QListWidget, QTreeWidget, QTextEdit {
                font-size: 12px;
            }
            QPushButton, QPushButton#primaryButton, QPushButton#secondaryButton,
            QPushButton#dialogPrimaryButton, QPushButton#disclosureButton,
            QPushButton#logShelf, QPushButton#helpIconButton {
                font-size: 12px;
            }
            QLabel#fieldLabel, QCheckBox { font-size: 12px; }
            QLabel#title { font-size: 16px; }
            QLabel#sectionTitle { font-size: 13px; }
            QLabel#logDrawerTitle { color: #f8fafc; font-size: 12px; font-weight: 600; }
            QLabel#eyebrow, QLabel#subtitle, QLabel#currentFile,
            QLabel#outputHint, QLabel#fileMeta, QLabel#fileEstimate,
            QLabel#fileStatus, QLabel#fileState,
            QPushButton#logShelf, QPushButton#disclosureButton {
                font-size: 11px;
            }
            QTextEdit#eventLog, QTreeWidget#resultsTree QHeaderView::section {
                font-size: 11px;
            }
            """
            + SHARED_MAIN_QSS
        )
        for control in (
            self.pick_button,
            self.remove_button,
            *self.assessment_primary_controls,
            *self.assessment_action_buttons,
            self.run_button,
            self.archive_mode_select,
            self.archive_library_button,
            self.image_archive_mode_select,
            self.image_library_button,
            self.archive_disclosure_button,
        ):
            control.setMinimumHeight(
                max(control.minimumHeight(), control.sizeHint().height())
            )
        self.refresh_file_list()
        install_control_help(self)

    def show_dialog(
        self,
        title: str,
        message: str,
        buttons: list[tuple[str, str, bool]] | None = None,
    ) -> str:
        dialog_parent = QApplication.activeWindow() or self
        dialog = StyledDialog(
            dialog_parent,
            title,
            message,
            buttons or [(self.text["ok_button"], "ok", True)],
        )
        dialog.exec()
        return dialog.result_key

    def show_help(self) -> None:
        self.show_dialog(self.text["help_title"], self.text["help_body"])

    def choose_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            self.text["choose_title"],
            "",
            self.text["choose_filter"],
        )
        if selected:
            self.set_files([Path(path) for path in selected])

    def set_files(self, paths: list[Path]) -> None:
        if self.is_running:
            return
        seen: set[Path] = set()
        resolved_paths: list[Path] = []
        for path in paths:
            resolved = path.expanduser().resolve()
            if (
                resolved.suffix.lower() not in SUPPORTED_COMPACTOR_INPUT_EXTENSIONS
                or resolved in seen
            ):
                continue
            seen.add(resolved)
            resolved_paths.append(resolved)
        existing = list(getattr(self, "input_paths", []))
        for path in resolved_paths:
            if path not in existing:
                existing.append(path)
                self.file_statuses[path] = "pending"
        self.input_paths = existing
        if not hasattr(self, "file_statuses"):
            self.file_statuses: dict[Path, str] = {}
        for path in self.input_paths:
            self.file_statuses.setdefault(path, "pending")
        self.last_settings_signature = self.current_settings_signature()
        self.refresh_file_list()
        self.refresh_info_text()
        self.current_file_label.setText(
            self.text["ready_status"].format(count=len(self.input_paths))
        )

    def state_label(self, state: str | None) -> str:
        if not state:
            return ""
        key = {
            "pending": "pending_marker",
            "queued": "queued_marker",
            "running": "running_marker",
            "done": "done_marker",
            "skipped": "skip_marker",
            "failed": "fail_marker",
            "stopped": "stopped_marker",
        }.get(state)
        return self.text[key] if key else ""

    def refresh_file_list(self) -> None:
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
            if item.data(Qt.ItemDataRole.UserRole)
        }
        self.file_list.clear()
        total_estimate = 0.0
        if not self.input_paths:
            self.refresh_results_tree()
            self.update_total_estimate_label(0, 0.0)
            self.update_remove_button_state()
            self.file_list.viewport().update()
            self.sync_run_button_state()
            return
        icon_provider = QFileIconProvider()
        for path in self.input_paths:
            try:
                size_mb = path.stat().st_size / 1024 / 1024
                estimate_mb = self.estimated_output_size_mb(path, size_mb)
                total_estimate += estimate_mb
            except OSError:
                size_mb = 0.0
                estimate_mb = 0.0

            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            item_width = max(0, self.file_list.viewport().width() - 16)
            item.setSizeHint(QSize(item_width, 58))
            row = QWidget()
            row.setObjectName("fileItem")
            state = self.file_statuses.get(path) or "pending"
            item.setData(Qt.ItemDataRole.UserRole + 1, state)
            row.setProperty("state", state)
            row.setProperty("selected", "true" if path in selected_paths else "false")
            row.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            row.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            row.setMinimumHeight(58)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(8, 4, 8, 4)
            row_layout.setSpacing(8)

            type_icon = QLabel(row)
            type_icon.setObjectName("fileTypeIcon")
            type_icon.setFixedSize(30, 30)
            type_icon.setPixmap(
                icon_provider.icon(QFileInfo(str(path))).pixmap(QSize(24, 24))
            )
            type_icon.setAlignment(Qt.AlignCenter)
            type_icon.setToolTip(f"{path.suffix.upper().lstrip('.')} 文件")

            text_widget = QWidget(row)
            text_layout = QVBoxLayout(text_widget)
            text_layout.setContentsMargins(0, 0, 0, 0)
            text_layout.setSpacing(2)
            name_label = ElidedLabel(path.name)
            name_label.setObjectName("fileName")
            name_label.setToolTip(str(path))
            name_label.setWordWrap(False)
            name_label.setSizePolicy(
                QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
            )
            suffix = path.suffix.lower().lstrip(".")
            source_tag = suffix.upper() if suffix else "FILE"
            state_label = self.state_label(state)
            metadata = (
                f"{format_file_size(int(size_mb * 1024 * 1024))} · {source_tag}"
                f" · {self.text['estimate_label']} {estimate_mb:.1f} MB · {state_label}"
            )
            metadata_label = ElidedLabel(metadata)
            metadata_label.setObjectName("fileMeta")
            metadata_label.setToolTip(f"{path}\n{metadata}")
            text_layout.addWidget(name_label)
            text_layout.addWidget(metadata_label)
            row_layout.addWidget(type_icon)
            row_layout.addWidget(text_widget, 1)
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, row)
            if path in selected_paths:
                item.setSelected(True)
        self.refresh_results_tree()
        self.update_total_estimate_label(len(self.input_paths), total_estimate)
        self.update_remove_button_state()
        self.sync_run_button_state()

    def refresh_results_tree(self) -> None:
        if not hasattr(self, "results_tree"):
            return
        self.results_tree.clear()
        for source in self.input_paths:
            output = self.output_paths.get(source)
            source_size = source.stat().st_size if source.is_file() else 0
            output_size = output.stat().st_size if output and output.is_file() else 0
            saved = max(0, source_size - output_size) if output_size else 0
            saving = (
                f"{format_file_size(saved)} ({saved / source_size:.0%})"
                if source_size and output_size
                else "—"
            )
            state = self.file_statuses.get(source, "pending")
            row = QTreeWidgetItem(
                [
                    source.name,
                    source.suffix.lstrip(".").upper(),
                    format_file_size(source_size),
                    format_file_size(output_size) if output_size else "—",
                    saving,
                    self.state_label(state),
                    str(output) if output else "—",
                ]
            )
            row.setToolTip(0, str(source))
            if output:
                row.setToolTip(6, str(output))
            self.results_tree.addTopLevelItem(row)

    def update_total_estimate_label(self, count: int, total_estimate: float) -> None:
        if count <= 0:
            message = self.text["total_estimate_empty"]
            tooltip = self.text["initial_status"]
        else:
            estimate = self.text["total_estimate_ready"].format(
                count=count, size=total_estimate
            )
            message = f"共 {count} 个" if self.language == "zh" else f"{count} file(s)"
            tooltip = self.text["ready_status"].format(count=count)
            tooltip = f"{estimate}\n{tooltip}"
        self.total_estimate_label.setText(message)
        self.total_estimate_label.setToolTip(tooltip)

    def sync_run_button_state(self) -> None:
        if not hasattr(self, "run_button"):
            return
        if self.is_running:
            if self.run_button.text() != self.text["stopping"]:
                self.run_button.setEnabled(True)
            self.run_button.setToolTip(self.text["stop"])
            return
        enabled = bool(self.input_paths)
        if hasattr(self, "file_list") and not self.file_list.isEnabled():
            enabled = False
        self.run_button.setEnabled(enabled)
        tooltip = self.current_file_label.text().strip() or self.text["initial_status"]
        self.run_button.setToolTip(tooltip)

    def on_file_selection_changed(self) -> None:
        self.update_remove_button_state()
        self.update_file_item_selection_styles()

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

    def estimated_output_size_mb(self, path: Path, source_size_mb: float) -> float:
        try:
            target_size = self.current_target_size()
        except ValueError:
            target_size = None
        if target_size is not None:
            return min(target_size, source_size_mb)
        video_profile = self.profile_select.currentData() or "high"
        image_profile = self.image_profile_select.currentData() or "high"
        video_ratio = float(PROFILE_QUALITY_RULES[video_profile]["estimate_ratio"])
        image_ratio = float(IMAGE_QUALITY_RULES[image_profile]["estimate_ratio"])
        suffix = path.suffix.lower()
        if suffix in VIDEO_EXTENSIONS:
            return max(0.01, source_size_mb * video_ratio)
        if suffix in IMAGE_EXTENSIONS:
            return max(0.01, source_size_mb * image_ratio)
        source_bytes = int(source_size_mb * 1024 * 1024)
        try:
            with ZipFile(path) as zf:
                video_bytes = 0
                image_bytes = 0
                for info in zf.infolist():
                    suffix = Path(info.filename).suffix.lower()
                    if (
                        info.filename.startswith("ppt/media/")
                        and suffix in VIDEO_EXTENSIONS
                    ):
                        video_bytes += int(info.compress_size or info.file_size)
                    elif (
                        info.filename.startswith(
                            ("ppt/media/", "word/media/", "xl/media/")
                        )
                        and suffix in IMAGE_EXTENSIONS
                    ):
                        image_bytes += int(info.compress_size or info.file_size)
        except Exception:
            ratio = min(video_ratio, image_ratio)
            return max(0.01, source_size_mb * ratio)
        other_bytes = max(0, source_bytes - video_bytes - image_bytes)
        estimated_bytes = int(
            other_bytes + video_bytes * video_ratio + image_bytes * image_ratio
        )
        return max(0.01, min(source_bytes, estimated_bytes) / 1024 / 1024)

    def current_target_size(self) -> float | None:
        value = self.target_input.text().strip()
        if not value:
            return None
        return float(value)

    def current_settings_signature(self) -> tuple[float | None, str, str, bool]:
        try:
            target_size = self.current_target_size()
        except ValueError:
            target_size = None
        return (
            target_size,
            self.profile_select.currentData() or "high",
            self.image_profile_select.currentData() or "high",
            self.overwrite_checkbox.isChecked(),
        )

    def reset_statuses_for_settings_change(self) -> None:
        if self._suppress_settings_reset:
            return
        if self.is_running or not self.input_paths:
            return
        signature = self.current_settings_signature()
        if (
            self.last_settings_signature is not None
            and signature != self.last_settings_signature
        ):
            self.file_statuses = {path: "pending" for path in self.input_paths}
            self.cleanup_session_reports()
            self.output_paths.clear()
            self.current_file_label.setText(
                self.text["ready_status"].format(count=len(self.input_paths))
            )
        self.last_settings_signature = signature

    def on_profile_changed(self) -> None:
        if hasattr(self, "video_threshold_spinbox") and not getattr(
            self, "_video_threshold_manual_override", False
        ):
            threshold = VIDEO_PRESET_THRESHOLDS.get(
                str(self.profile_select.currentData() or "high")
            )
            if threshold is not None:
                blocked = self.video_threshold_spinbox.blockSignals(True)
                self.video_threshold_spinbox.setValue(threshold)
                self.video_threshold_spinbox.blockSignals(blocked)
        self.reset_statuses_for_settings_change()
        self.refresh_file_list()
        self.refresh_info_text()

    def on_video_threshold_changed(self, value: float) -> None:
        settings = app_settings()
        settings.setValue("compression/video_ssim_threshold", value)
        settings.setValue("compression/video_ssim_threshold_manual", True)
        self._video_threshold_manual_override = True

    def on_target_changed(self) -> None:
        self.reset_statuses_for_settings_change()
        self.refresh_file_list()
        self.refresh_info_text()

    def mark_item_pending(self, item: QListWidgetItem) -> None:
        if self.is_running:
            return
        raw_path = item.data(Qt.ItemDataRole.UserRole)
        if not raw_path:
            return
        path = Path(raw_path)
        if self.file_statuses.get(path) in {"done", "skipped", "failed", "stopped"}:
            self.file_statuses[path] = "pending"
            self.refresh_file_list()
            self.current_file_label.setText(
                self.text["ready_status"].format(count=len(self.input_paths))
            )

    def pending_paths(self) -> list[Path]:
        runnable_states = {"pending", "queued", "running", "stopped"}
        return [
            path
            for path in self.input_paths
            if self.file_statuses.get(path, "pending") in runnable_states
        ]

    def remove_selected_files(self) -> None:
        if self.is_running:
            return
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
            if item.data(Qt.ItemDataRole.UserRole)
        }
        if not selected_paths:
            return
        self.input_paths = [
            path for path in self.input_paths if path not in selected_paths
        ]
        for path in selected_paths:
            self.file_statuses.pop(path, None)
            output_path = self.output_paths.pop(path, None)
            if output_path is not None:
                self.cleanup_report_path(output_path.with_suffix(".report.json"))
        self.refresh_file_list()
        self.refresh_info_text()
        if self.input_paths:
            self.current_file_label.setText(
                self.text["ready_status"].format(count=len(self.input_paths))
            )
        else:
            self.current_file_label.setText(self.text["initial_status"])
        self.sync_run_button_state()

    def restyle_widget(self, widget: QWidget) -> None:
        widget.style().unpolish(widget)
        widget.style().polish(widget)

    def configured_video_library_root(self) -> Path | None:
        settings = app_settings()
        configured = settings.value("video_manager/last_project", "", str)
        root = Path(configured) if configured else None
        if root is not None and (root / "video-project.json").is_file():
            return root
        return None

    def configured_image_library_root(self) -> Path | None:
        settings = app_settings()
        configured = settings.value("image_manager/last_project", "", str)
        root = Path(configured) if configured else None
        if root is not None and (root / "image-project.json").is_file():
            return root
        return None

    def on_activated(self) -> None:
        self.refresh_archive_library_target()
        self.refresh_image_library_target()

    def refresh_archive_library_target(self) -> None:
        if not hasattr(self, "archive_library_button"):
            return
        enabled = self.archive_mode_select.currentData() != "off"
        root = self.configured_video_library_root()
        if not enabled:
            self.archive_library_button.setText(self.text["archive_library_off"])
            self.archive_library_button.setToolTip(self.text["archive_videos_tooltip"])
        elif root is None:
            self.archive_library_button.setText(self.text["archive_library_unset"])
            self.archive_library_button.setToolTip(self.text["archive_library_tooltip"])
        else:
            self.archive_library_button.setText(
                self.text["archive_library_name"].format(name=root.name)
            )
            self.archive_library_button.setToolTip(
                f"{self.text['archive_library_tooltip']}\n{root}"
            )
        self.archive_library_button.setEnabled(enabled and not self.is_running)
        self.refresh_archive_summary()

    def choose_archive_library(self) -> Path | None:
        current = self.configured_video_library_root()
        directory = QFileDialog.getExistingDirectory(
            self,
            self.text["archive_library_unset"],
            str(current or Path.home()),
        )
        if not directory:
            return None
        root = Path(directory)
        try:
            if (root / "video-project.json").is_file():
                library = VideoProject.open(root)
            else:
                library = VideoProject.create(root, "视频库")
        except Exception as exc:
            self.show_dialog(self.text["archive_category_invalid"], str(exc))
            return None
        app_settings().setValue("video_manager/last_project", str(library.root))
        self.refresh_archive_library_target()
        return library.root

    def refresh_image_library_target(self) -> None:
        enabled = self.image_archive_mode_select.currentData() != "off"
        root = self.configured_image_library_root()
        if not enabled:
            self.image_library_button.setText(self.text["image_library_off"])
        elif root is None:
            self.image_library_button.setText(self.text["image_library_unset"])
        else:
            self.image_library_button.setText(
                self.text["image_library_name"].format(name=root.name)
            )
        self.image_library_button.setToolTip(self.text["image_library_tooltip"])
        self.image_library_button.setEnabled(enabled and not self.is_running)
        self.image_archive_category_input.setEnabled(enabled and not self.is_running)
        self.refresh_archive_summary()

    def refresh_archive_summary(self) -> None:
        if not hasattr(self, "archive_summary_label"):
            return
        video = self.archive_mode_select.currentText()
        video_root = self.configured_video_library_root()
        if self.archive_mode_select.currentData() != "off" and video_root is not None:
            video = f"{video} → {video_root.name}"
        image = self.image_archive_mode_select.currentText()
        image_root = self.configured_image_library_root()
        if (
            self.image_archive_mode_select.currentData() != "off"
            and image_root is not None
        ):
            image = f"{image} → {image_root.name}"
        summary = self.text["resource_archive_summary"].format(
            video=video,
            image=image,
        )
        self.archive_summary_label.setText(summary)
        self.archive_summary_label.setToolTip(summary)

    def toggle_archive_settings(self, expanded: bool) -> None:
        self.archive_settings_panel.setVisible(expanded)
        self.archive_disclosure_button.setText(
            self.text[
                "resource_archive_collapse" if expanded else "resource_archive_expand"
            ]
        )

    def choose_image_library(self) -> Path | None:
        current = self.configured_image_library_root()
        directory = QFileDialog.getExistingDirectory(
            self,
            self.text["image_library_unset"],
            str(current or Path.home()),
        )
        if not directory:
            return None
        root = Path(directory)
        try:
            library = (
                ImageProject.open(root)
                if (root / "image-project.json").is_file()
                else ImageProject.create(root, "图片库")
            )
        except Exception as exc:
            self.show_dialog(self.text["archive_category_invalid"], str(exc))
            return None
        app_settings().setValue("image_manager/last_project", str(library.root))
        self.refresh_image_library_target()
        return library.root

    def set_running(self, running: bool) -> None:
        self.is_running = running
        self.run_button.setText(self.text["stop"] if running else self.text["run"])
        self.run_button.setProperty("mode", "stop" if running else "run")
        self.restyle_widget(self.run_button)
        self.pick_button.setDisabled(running)
        self.target_input.setDisabled(running)
        self.profile_select.setDisabled(running)
        self.image_profile_select.setDisabled(running)
        self.archive_mode_select.setDisabled(running)
        self.archive_library_button.setDisabled(
            running or self.archive_mode_select.currentData() == "off"
        )
        self.overwrite_checkbox.setDisabled(
            running or self.archive_mode_select.currentData() == "off"
        )
        self.archive_category_input.setDisabled(
            running or self.archive_mode_select.currentData() == "off"
        )
        self.image_archive_mode_select.setDisabled(running)
        self.image_library_button.setDisabled(
            running or self.image_archive_mode_select.currentData() == "off"
        )
        self.image_archive_category_input.setDisabled(
            running or self.image_archive_mode_select.currentData() == "off"
        )
        self.auto_optimize_checkbox.setDisabled(running)
        self.target_gpu_checkbox.setDisabled(running)
        self.standard_encoder_select.setDisabled(running)
        if hasattr(self, "forced_button"):
            self.forced_button.setDisabled(running or not self.forced_candidates)
        self.video_threshold_spinbox.setDisabled(running)
        self.image_threshold_spinbox.setDisabled(running)
        self.update_remove_button_state()
        self.update_audit_button_state()
        self.sync_run_button_state()
        self.refresh_log_shelf()
        if running:
            self.show_log_drawer()

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        if (
            watched is getattr(self, "content_widget", None)
            and event.type() == QEvent.Type.Resize
        ):
            self.update_responsive_layout(event.size().width())
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
            hasattr(self, name) for name in ("settings_controls_layout", "results_tree")
        ):
            return
        compact = width < 1180
        left_width = 260 if compact else 320
        self.left_pane.setFixedWidth(left_width)
        column_widths = (
            (150, 48, 62, 62, 56, 56) if compact else (210, 64, 84, 84, 96, 90)
        )
        for column, column_width in enumerate(column_widths):
            self.results_tree.setColumnWidth(column, column_width)
        layout = self.settings_controls_layout
        for widget in self.settings_control_groups:
            layout.removeWidget(widget)
        target_group, profile_group, image_group = self.settings_control_groups
        layout.addWidget(target_group, 0, 0)
        layout.addWidget(profile_group, 0, 1)
        layout.addWidget(image_group, 0, 2)
        self.settings_controls_widget.setMinimumHeight(30)
        root_margins = self.root_layout.contentsMargins()
        settings_margins = self.settings_card.layout().contentsMargins()
        assessment_width = (
            width
            - root_margins.left()
            - root_margins.right()
            - left_width
            - self.body_layout.spacing()
            - settings_margins.left()
            - settings_margins.right()
            - self.settings_card.frameWidth() * 2
        )
        self._update_assessment_row_layout(
            assessment_width < self._assessment_row_required_width()
        )
        self._fit_assessment_row_controls()

    def _fit_assessment_row_controls(self) -> None:
        """Keep fixed threshold fields inside the row on native-style variants."""
        if not hasattr(self, "assessment_controls_layout"):
            return
        if not self.isVisible():
            return
        for spinbox in (self.video_threshold_spinbox, self.image_threshold_spinbox):
            spinbox.setFixedWidth(self._threshold_control_width)
        self.assessment_controls_layout.activate()
        row_width = self.assessment_row_primary_widget.width()
        if row_width <= 0:
            return
        right_edge = max(
            (
                control.mapTo(self.assessment_row_primary_widget, QPoint(0, 0)).x()
                + control.width()
                for control in self.assessment_primary_controls
            ),
            default=0,
        )
        overflow = max(0, right_edge - row_width)
        if overflow:
            self.image_threshold_spinbox.setFixedWidth(
                max(64, self._threshold_control_width - overflow)
            )
            self.assessment_controls_layout.activate()

    def _assessment_row_required_width(self) -> int:
        widgets = [
            *self.assessment_primary_controls,
            self.audit_button,
            self.optimize_button,
        ]
        if not self.forced_button.isHidden():
            widgets.append(self.forced_button)
        widths = [
            max(
                widget.minimumWidth(),
                min(widget.sizeHint().width(), widget.maximumWidth()),
            )
            for widget in widgets
        ]
        # Keep a conservative estimate for the fixed controls and inter-control gaps.
        return sum(widths) + self.assessment_controls_layout.spacing() * len(widgets)

    def _update_assessment_row_layout(self, compact: bool) -> None:
        """宽窗口：评估控件与按钮组成右对齐的一组；窄窗口：拆两行。"""
        if not all(
            hasattr(self, name)
            for name in (
                "assessment_controls_layout",
                "assessment_actions_layout",
                "assessment_row_actions_widget",
            )
        ):
            return
        primary = self.assessment_controls_layout
        actions = self.assessment_actions_layout
        action_buttons = self.assessment_action_buttons
        if compact:
            # 窄窗口：按钮回到标题行右侧，阈值控件单独一行
            primary.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            for button in action_buttons:
                primary.removeWidget(button)
            for button in action_buttons:
                actions.addWidget(button)
            self.assessment_row_actions_widget.setVisible(True)
        else:
            # 宽窗口：按钮移入 primary 行，整组控件靠右对齐
            primary.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            for button in action_buttons:
                actions.removeWidget(button)
            for button in action_buttons:
                primary.addWidget(button)
            self.assessment_row_actions_widget.setVisible(False)

    def position_log_drawer(self) -> None:
        if self.log_drawer.isHidden():
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
        status = self.current_file_label.text().strip() or self.text["initial_status"]
        prefix = "状态与日志" if self.language == "zh" else "Status and logs"
        self.log_shelf.setText(f"{prefix} · {self.progress_bar.value()}% · {status}")

    def update_audit_button_state(self) -> None:
        if not hasattr(self, "audit_button"):
            return
        if self.is_running:
            self.audit_button.setDisabled(True)
            return

        has_done = False
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            raw_path = item.data(Qt.ItemDataRole.UserRole)
            source = Path(raw_path) if raw_path else None
            output = self.output_paths.get(source) if source else None
            if (
                source
                and output
                and output != source
                and item.data(Qt.ItemDataRole.UserRole + 1) == "done"
            ):
                has_done = True
                break
        self.audit_button.setDisabled(not has_done)

    def update_remove_button_state(self) -> None:
        if not hasattr(self, "remove_button"):
            return
        has_selection = any(
            item.data(Qt.ItemDataRole.UserRole)
            for item in self.file_list.selectedItems()
        )
        self.remove_button.setDisabled(self.is_running or not has_selection)

    def run_audit(self, target_paths: set[Path] | None = None) -> None:
        self.failed_audits.clear()
        done_items = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            raw_path = item.data(Qt.ItemDataRole.UserRole)
            source = Path(raw_path) if raw_path else None
            output = self.output_paths.get(source) if source else None
            if (
                source
                and output
                and output != source
                and item.data(Qt.ItemDataRole.UserRole + 1) == "done"
            ):
                done_items.append(item)

        if not done_items:
            self.append_log("没有已完成的文件可供评估。")
            return

        self.audit_button.setDisabled(True)
        self.run_button.setDisabled(True)
        self.run_button.setToolTip(self.text.get("audit_button", "画质评估"))
        self.file_list.setDisabled(True)
        self.progress_bar.setValue(0)

        if target_paths is None:
            selected_done = [
                item for item in self.file_list.selectedItems() if item in done_items
            ]
            targets = selected_done if selected_done else done_items
        else:
            targets = [
                item
                for item in done_items
                if Path(item.data(Qt.ItemDataRole.UserRole)) in target_paths
            ]

        self.audit_queue = [
            Path(item.data(Qt.ItemDataRole.UserRole)) for item in targets
        ]
        self._audit_next()

    def on_optimize_clicked(self) -> None:
        if not hasattr(self, "failed_audits") or not self.failed_audits:
            return
        self.optimize_button.setDisabled(True)
        self.append_log(f"🚀 正在准备增量重试 ({len(self.failed_audits)} 个文件)...")
        self.recompress_incremental(self.failed_audits)

    def recompress_incremental(self, failed_audits: dict) -> None:
        from pathlib import Path

        current_video_profile = self.profile_select.currentData()
        current_image_profile = self.image_profile_select.currentData()

        def optimize_profile(prof: str) -> str:
            if prof == "none":
                return "none"
            if prof == "high":
                return "high"
            if prof in {"balanced", "aggressive"}:
                return "high"
            return "none"

        all_failed = []
        for assets in failed_audits.values():
            all_failed.extend(assets)

        has_failed_video = any(r.is_video for r in all_failed)
        has_failed_image = any(not r.is_video for r in all_failed)
        new_video_profile = (
            optimize_profile(current_video_profile)
            if has_failed_video
            else current_video_profile
        )
        new_image_profile = (
            optimize_profile(current_image_profile)
            if has_failed_image
            else current_image_profile
        )

        self._suppress_settings_reset = True
        try:
            self.profile_select.setCurrentIndex(PRESET_OPTIONS.index(new_video_profile))
            self.image_profile_select.setCurrentIndex(
                IMAGE_PRESET_OPTIONS.index(new_image_profile)
            )
            self.target_input.clear()
        finally:
            self._suppress_settings_reset = False

        self.append_log(
            f"▶ 增量提档策略：视频({new_video_profile}), 图片({new_image_profile})。"
            "低/中档失败项按高保真重压；高保真仍失败时保留压缩结果并停止。"
        )

        incremental_patch_map: dict[Path, dict[str, object]] = {}
        incremental_retry_context: dict[Path, dict[str, object]] = {}
        retry_paths: list[Path] = []

        for raw_input_pptx, failed_assets in failed_audits.items():
            source_key = Path(raw_input_pptx)
            input_pptx = source_key.expanduser().resolve()
            report_path = self.output_paths.get(input_pptx) or self.output_paths.get(
                source_key
            )
            if report_path:
                report_path = report_path.with_suffix(".report.json")
            if not report_path or not report_path.exists():
                self.append_log(
                    f"无法提取增量缓存 {input_pptx.name}：找不到 report.json"
                )
                continue

            try:
                report_data = load_json_file(report_path, source="Incremental report")
            except (OSError, ValueError, SystemExit):
                self.append_log(
                    f"无法提取增量缓存 {input_pptx.name}：report.json 不可读"
                )
                continue
            output_pptx_str = report_data.get("output_pptx")
            if not output_pptx_str or not Path(output_pptx_str).exists():
                self.append_log(
                    f"无法提取增量缓存 {input_pptx.name}：上一版输出文件不存在"
                )
                continue

            output_pptx = Path(output_pptx_str)
            failed_paths = {r.media_path for r in failed_assets}
            if not failed_paths:
                continue
            incremental_patch_map[input_pptx] = {
                "failed_media_paths": failed_paths,
                "base_output_pptx": output_pptx,
                "report_path": report_path,
            }
            incremental_retry_context[input_pptx] = {
                "video_profile": new_video_profile,
                "image_profile": new_image_profile,
            }
            retry_paths.append(input_pptx)
            video_count = sum(1 for r in failed_assets if r.is_video)
            image_count = sum(1 for r in failed_assets if not r.is_video)
            self.append_log(
                f"[{input_pptx.name}] 将只从源文档抽取并重写 {video_count} 个视频、{image_count} 张图片。"
            )

        if not retry_paths:
            self.append_log("没有可提档的文件：缺少上一版输出或 report.json。")
            self.cleanup_incremental_temp_dirs()
            self.update_audit_button_state()
            return

        self.incremental_patch_map = incremental_patch_map
        self.incremental_retry_context = incremental_retry_context
        self.input_paths = retry_paths
        self.file_statuses = {path: "pending" for path in retry_paths}
        self.last_settings_signature = self.current_settings_signature()
        self.refresh_file_list()
        self.refresh_info_text()
        self.auto_audit_after_compression = True
        self.start_job()

    def cleanup_incremental_temp_dirs(self) -> None:
        temp_dirs = getattr(self, "_incremental_temp_dirs", [])
        for temp_dir in temp_dirs:
            shutil.rmtree(temp_dir, ignore_errors=True)
        if hasattr(self, "_incremental_temp_dirs"):
            delattr(self, "_incremental_temp_dirs")

    def remember_report_path(self, output_path: Path | None) -> None:
        if output_path is None:
            return
        report_path = output_path.with_suffix(".report.json")
        if report_path.exists():
            self.session_report_paths.add(report_path.resolve())

    def cleanup_report_path(self, report_path: Path) -> None:
        try:
            resolved = report_path.resolve()
        except Exception:
            resolved = report_path
        try:
            resolved.unlink(missing_ok=True)
        except Exception:
            pass
        self.session_report_paths.discard(resolved)

    def cleanup_session_reports(self) -> None:
        for report_path in list(self.session_report_paths):
            self.cleanup_report_path(report_path)

    def _audit_next(self) -> None:
        if not hasattr(self, "audit_queue") or not self.audit_queue:
            self.update_audit_button_state()
            self.file_list.setDisabled(False)
            self.sync_run_button_state()
            self.progress_bar.setValue(100)
            self.append_log("=" * 40)
            self.append_log("全部评估完成！")
            if self.auto_optimize_checkbox.isChecked() and self.failed_audits:
                self.on_optimize_clicked()
            return

        input_path = self.audit_queue.pop(0)
        if input_path.suffix.lower() == ".pdf":
            self.append_log(
                f"PDF 不是 ZIP 结构，暂不支持画质评估，跳过: {input_path.name}"
            )
            self._audit_next()
            return
        output_path = self.output_paths.get(input_path)
        if output_path:
            report_path = output_path.with_suffix(".report.json")
        else:
            report_path = input_path.with_name(f"{input_path.stem}.report.json")

        if not report_path.exists():
            self.append_log(f"找不到对应的报告文件，跳过评估: {input_path.name}")
            self._audit_next()
            return

        self.append_log("=" * 40)
        self.append_log(f"正在启动画质评估 (SSIM) - {input_path.name}")
        self.append_log("=" * 40)

        self.audit_worker = QualityAuditWorker(input_path, Path("dummy"), report_path)

        self.audit_thread = QThread()
        self.audit_worker.moveToThread(self.audit_thread)
        self.audit_thread.started.connect(self.audit_worker.run)
        self.audit_worker.finished.connect(self.audit_thread.quit)
        self.audit_worker.finished.connect(self.audit_worker.deleteLater)
        self.audit_thread.finished.connect(self.audit_thread.deleteLater)
        self.audit_worker.progress.connect(self.on_progress)
        self.audit_worker.log.connect(self.append_log)
        self.audit_worker.finished.connect(self._on_single_audit_finished)
        self.audit_thread.start()

    def _on_single_audit_finished(self, results: list) -> None:
        if not results:
            self.append_log("评估未产生结果。")
        else:
            videos = [r for r in results if r.is_video and r.status == "success"]
            images = [r for r in results if not r.is_video and r.status == "success"]
            errors = [r for r in results if r.status == "error"]

            def score_to_grade(score: float) -> str:
                if score >= 0.98:
                    return "A+ (无损视觉)"
                if score >= 0.95:
                    return "A (极优)"
                if score >= 0.90:
                    return "B (良好)"
                return "C (明显有损)"

            if videos:
                avg_v = sum(v.ssim for v in videos) / len(videos)
                min_v = min(videos, key=lambda x: x.ssim)
                max_v = max(videos, key=lambda x: x.ssim)
                self.append_log(
                    f"▶ 综合视频画质: {score_to_grade(avg_v)} ({avg_v:.3f})"
                )
                self.append_log(
                    f"  - 最高: {max_v.ssim:.3f} ({Path(max_v.media_path).name})"
                )
                self.append_log(
                    f"  - 最低: {min_v.ssim:.3f} ({Path(min_v.media_path).name})"
                )

            if images:
                avg_i = sum(i.ssim for i in images) / len(images)
                self.append_log(
                    f"▶ 综合图片画质: {score_to_grade(avg_i)} ({avg_i:.3f})"
                )
                self.append_log(f"  - 共评估了 {len(images)} 张图片")

            if errors:
                self.append_log(f"▶ {len(errors)} 个素材评估失败：")
                for err in errors:
                    self.append_log(f"    - {err.media_path}: {err.error}")

            failed_assets = []
            for r in videos + images:
                threshold = (
                    self.video_threshold_spinbox.value()
                    if r.is_video
                    else self.image_threshold_spinbox.value()
                )
                if r.ssim is not None and r.ssim < threshold:
                    failed_assets.append(r)

            retry_context_map = getattr(self, "incremental_retry_context", {})
            retry_context = (
                retry_context_map.pop(self.audit_worker.input_pptx, None)
                if retry_context_map
                else None
            )
            can_incremental_optimize = (
                self.audit_worker.input_pptx.suffix.lower() == ".pptx"
            )

            if failed_assets:
                self.append_log(
                    f"💡 检测到 {len(failed_assets)} 个素材画质低于设定阈值"
                    f"（视频 {self.video_threshold_spinbox.value():.2f} / "
                    f"图片 {self.image_threshold_spinbox.value():.2f}）"
                )
                if self.target_input.text().strip():
                    self.append_log("  - 注意：强制指定目标大小可能会导致画质下降。")
                retryable_assets = failed_assets
                terminal_assets = []
                if retry_context:
                    retryable_assets = []
                    for asset in failed_assets:
                        current_profile = (
                            retry_context["video_profile"]
                            if asset.is_video
                            else retry_context["image_profile"]
                        )
                        if current_profile == "high":
                            terminal_assets.append(asset)
                        elif current_profile == "none":
                            terminal_assets.append(asset)
                        else:
                            retryable_assets.append(asset)
                    if retryable_assets:
                        self.append_log(
                            self.text["optimize_prompt_restore"].format(
                                count=len(retryable_assets)
                            )
                        )
                    if terminal_assets:
                        self.append_log(
                            self.text["optimize_terminal_mixed"].format(
                                count=len(terminal_assets)
                            )
                        )
                    if not retryable_assets and terminal_assets:
                        self.append_log(
                            self.text["optimize_terminal_original"].format(
                                count=len(terminal_assets)
                            )
                        )
                else:
                    if can_incremental_optimize:
                        self.append_log(
                            "💡 您可以点击右上角的【提档优化】按钮，仅对受损素材升档重压。"
                        )
                    else:
                        self.append_log(
                            "💡 独立图片/视频当前仅支持画质评估；提档优化仍只支持 PPTX。"
                        )

                if retryable_assets and can_incremental_optimize:
                    self.failed_audits[self.audit_worker.input_pptx] = retryable_assets
                    self.optimize_button.setDisabled(False)
                else:
                    self.failed_audits.pop(self.audit_worker.input_pptx, None)
                    self.optimize_button.setDisabled(True)
            elif retry_context:
                self.append_log("✅ 增量提档后的复评已全部达标。")

        self._audit_next()

    def run_or_stop(self) -> None:
        if self.is_running:
            self.stop_job()
        else:
            self.start_job()

    def run_forced_candidates(self) -> None:
        candidates = [path for path in self.forced_candidates if path.is_file()]
        if not candidates or self.is_running:
            return
        answer = self.show_dialog(
            "确认生成强制版" if self.language == "zh" else "Confirm forced version",
            (
                "强制版会使用更低但仍有绝对红线的质量阈值，另存新文件并保留安全版。"
                "仍可能无法达到目标容量。"
                if self.language == "zh"
                else "The forced version uses lower absolute quality floors, saves separately, "
                "and keeps the safe output. It may still miss the target."
            ),
            [
                (
                    "生成强制版" if self.language == "zh" else "Generate",
                    "forced",
                    False,
                ),
                ("取消" if self.language == "zh" else "Cancel", "cancel", True),
            ],
        )
        if answer != "forced":
            return
        for path in candidates:
            self.file_statuses[path] = "pending"
        self.forced_button.setVisible(False)
        self.update_responsive_layout(self.content_widget.width())
        self.start_job(quality_mode="forced", selected_paths=candidates)

    def stop_job(self) -> None:
        if not self.is_running or self.worker is None:
            return
        self.current_file_label.setText(self.text["stopping"])
        self.append_log(self.text["stopping"])
        self.run_button.setText(self.text["stopping"])
        self.run_button.setDisabled(True)
        self.run_button.setToolTip(self.text["stopping"])
        self.worker.cancel()

    def start_job(
        self,
        quality_mode: str = "safe",
        selected_paths: list[Path] | None = None,
    ) -> None:
        if self.is_running:
            self.stop_job()
            return
        if not self.input_paths:
            self.show_dialog(
                self.text["missing_file_title"], self.text["missing_file_body"]
            )
            return
        try:
            target_size_mb = self.current_target_size()
        except ValueError:
            self.show_dialog(
                self.text["invalid_size_title"], self.text["invalid_size_number"]
            )
            return
        if target_size_mb is not None and target_size_mb <= 0:
            self.show_dialog(
                self.text["invalid_size_title"], self.text["invalid_size_positive"]
            )
            return
        if (
            quality_mode == "safe"
            and target_size_mb is not None
            and len(self.input_paths) > 1
        ):
            answer = self.show_dialog(
                self.text["batch_size_title"],
                self.text["batch_size_body"],
                [
                    (self.text["preset_button"], "preset", True),
                    (self.text["continue_button"], "continue", False),
                ],
            )
            if answer == "preset":
                target_size_mb = None
                self.target_input.clear()

        profile = self.profile_select.currentData() or "high"
        image_profile = self.image_profile_select.currentData() or "high"
        if (
            profile == "none"
            and image_profile == "none"
            and getattr(self, "asset_pre_encoded_map", None) is None
            and getattr(self, "incremental_patch_map", None) is None
        ):
            self.show_dialog(
                self.text["no_media_profile_title"], self.text["no_media_profile_body"]
            )
            return
        pending_paths = selected_paths or self.pending_paths()
        if not pending_paths:
            self.show_dialog(
                self.text["no_pending_title"], self.text["no_pending_body"]
            )
            return
        if any(
            path.suffix.lower() in DOCUMENT_INPUT_EXTENSIONS for path in pending_paths
        ):
            if target_size_mb is None:
                self.show_dialog(
                    self.text["document_target_required_title"],
                    self.text["document_target_required_body"],
                )
                return
            if image_profile == "none":
                self.show_dialog(
                    self.text["document_image_profile_none_title"],
                    self.text["document_image_profile_none_body"],
                )
                return
        overwrite_original = (
            self.overwrite_checkbox.isChecked() and quality_mode == "safe"
        )
        if overwrite_original:
            overwrite_valid = (
                all(path.suffix.lower() == ".pptx" for path in pending_paths)
                and self.archive_mode_select.currentData() != "off"
                and profile != "none"
                and image_profile == "none"
                and not self.auto_optimize_checkbox.isChecked()
            )
            if not overwrite_valid:
                self.show_dialog(
                    self.text["overwrite_invalid_title"],
                    self.text["overwrite_invalid_body"],
                )
                return
            answer = self.show_dialog(
                self.text["overwrite_confirm_title"],
                self.text["overwrite_confirm_body"],
                [
                    (
                        self.text["overwrite_confirm_button"],
                        "overwrite",
                        False,
                    ),
                    (self.text["overwrite_cancel_button"], "cancel", True),
                ],
            )
            if answer != "overwrite":
                return
        video_library_root = None
        archive_category = self.archive_category_input.text().strip()
        if self.archive_mode_select.currentData() != "off" and any(
            path.suffix.lower() == ".pptx" for path in pending_paths
        ):
            try:
                normalize_library_category(archive_category)
            except ValueError as exc:
                self.show_dialog(self.text["archive_category_invalid"], str(exc))
                return
            video_library_root = self.configured_video_library_root()
            if video_library_root is None:
                video_library_root = self.choose_archive_library()
                if video_library_root is None:
                    self.append_log("[STOP] 已取消选择视频库，未开始压缩。")
                    return
        image_library_root = None
        image_archive_category = self.image_archive_category_input.text().strip()
        image_archive_mode = str(self.image_archive_mode_select.currentData())
        if image_archive_mode != "off" and any(
            path.suffix.lower() == ".pptx" or path.suffix.lower() in IMAGE_EXTENSIONS
            for path in pending_paths
        ):
            image_library_root = self.configured_image_library_root()
            if image_library_root is None:
                image_library_root = self.choose_image_library()
                if image_library_root is None:
                    self.append_log("[STOP] 已取消选择图片库，未开始压缩。")
                    return
        self.last_settings_signature = self.current_settings_signature()
        self.refresh_info_text()
        self.current_file_label.setText(self.text["running_status"])
        self.progress_bar.setValue(0)
        self.event_log.setPlainText(
            self.text["job_started"].format(count=len(pending_paths))
        )
        self.set_running(True)

        self.worker_thread = QThread(self)
        self.active_path = None
        for path in pending_paths:
            self.file_statuses[path] = "queued"
        self.refresh_file_list()
        self.worker = CompressionWorker(
            list(pending_paths),
            target_size_mb,
            profile,
            image_profile,
            self.language,
            self.text,
            video_library_root,
            str(self.archive_mode_select.currentData()),
            archive_category,
            image_library_root,
            image_archive_mode,
            image_archive_category,
            overwrite_original,
            self.target_gpu_checkbox.isChecked(),
            self.video_threshold_spinbox.value(),
            self.image_threshold_spinbox.value(),
            quality_mode,
            str(self.standard_encoder_select.currentData() or "auto"),
        )
        self.worker.asset_pre_encoded_map = getattr(self, "asset_pre_encoded_map", None)
        if hasattr(self, "asset_pre_encoded_map"):
            delattr(self, "asset_pre_encoded_map")
        self.worker.incremental_patch_map = getattr(self, "incremental_patch_map", None)
        if hasattr(self, "incremental_patch_map"):
            delattr(self, "incremental_patch_map")
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.file_started.connect(self.on_file_started)
        self.worker.status.connect(self.on_status)
        self.worker.progress.connect(self.on_progress)
        self.worker.log.connect(self.append_log)
        self.worker.file_completed.connect(self.on_file_completed)
        self.worker.file_failed.connect(self.on_file_failed)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def refresh_info_text(self, output_override: Path | None = None) -> None:
        if not self.input_paths:
            self.output_hint_label.setText(self.text["output_waiting_hint"])
            return

        if output_override is not None:
            self.output_hint_label.setText(self.text["output_done_hint"])
            return

        if self.overwrite_checkbox.isChecked():
            self.output_hint_label.setText(self.text["output_overwrite_hint"])
            return

        if len(self.input_paths) == 1:
            self.output_hint_label.setText(self.text["output_same_folder"])
        else:
            self.output_hint_label.setText(self.text["output_each_folder"])

    def on_progress(self, percent: int, label: str) -> None:
        self.progress_bar.setValue(percent)
        self.current_file_label.setText(label)
        if getattr(self, "last_progress_label", None) != label:
            self.last_progress_label = label
            self.append_log(f"[{percent:>3}%] {label}")
        self.refresh_log_shelf()

    def on_file_started(self, path: object) -> None:
        source_path = Path(path)
        self.active_path = source_path
        self.file_statuses[source_path] = "running"
        self.current_file_label.setText(
            f"{self.text['current_processing']}: {source_path.name}"
        )
        self.refresh_file_list()

    def on_file_completed(
        self,
        path: object,
        skipped: bool,
        output_path: object,
        size_bytes: int,
        reason: str,
        output_sha256: str,
    ) -> None:
        source_path = Path(path)
        output_pptx = Path(output_path)
        self.output_paths[source_path] = output_pptx
        self.remember_report_path(output_pptx)
        self.file_statuses[source_path] = "skipped" if skipped else "done"
        if self.active_path == source_path:
            self.active_path = None
        self.refresh_file_list()
        if skipped:
            self.append_log(f"[SKIP] {source_path.name}: {reason}")
        else:
            self.append_log(
                f"[DONE] {output_pptx} ({size_bytes / 1024 / 1024:.2f} MiB)"
            )

    def on_file_failed(self, path: object, message: str) -> None:
        source_path = Path(path)
        self.file_statuses[source_path] = "failed"
        if self.active_path == source_path:
            self.active_path = None
        self.refresh_file_list()
        self.append_log(f"[ERR] {source_path.name}: {message}")
        self.show_log_drawer(auto_hide=False)

    def on_status(self, message: str) -> None:
        self.current_file_label.setText(f"{self.text['current_processing']}: {message}")

    def append_log(self, message: str) -> None:
        LOGGER.info("%s", message)
        current = self.event_log.toPlainText().strip()
        if current == self.text["waiting_log"]:
            current = ""
        if current:
            current = current + "\n" + message
        else:
            current = message
        self.event_log.setPlainText(current)
        cursor = self.event_log.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.event_log.setTextCursor(cursor)
        self.refresh_log_shelf()
        if self.is_running:
            self.show_log_drawer(auto_hide=False)

    def on_finished(
        self,
        results: list,
        failures: list,
        cancelled: bool = False,
        stopped_path: object | None = None,
    ) -> None:
        completed = [item for item in results if not item[3]]
        skipped = [item for item in results if item[3]]
        for source_path, output_path, _, skipped_flag, _ in results:
            self.file_statuses[source_path] = "skipped" if skipped_flag else "done"
            if output_path:
                output_pptx = Path(output_path)
                self.output_paths[source_path] = output_pptx
                self.remember_report_path(output_pptx)
        for failed_path, _ in failures:
            self.file_statuses[failed_path] = "failed"
        self.forced_candidates = []
        for source_path, output_path, _, skipped_flag, _ in results:
            output = Path(output_path)
            if skipped_flag or "_forced" in output.stem:
                continue
            report_path = output.with_suffix(".report.json")
            try:
                report = load_json_file(report_path, source="Compression report")
            except (OSError, ValueError, SystemExit):
                continue
            if report.get("target", {}).get("status") == "quality_limited":
                self.forced_candidates.append(Path(source_path))
        self.forced_button.setVisible(bool(self.forced_candidates))
        self.forced_button.setEnabled(bool(self.forced_candidates))
        self.update_responsive_layout(self.content_widget.width())
        if cancelled:
            stopped = Path(stopped_path) if stopped_path else self.active_path
            for path, state in list(self.file_statuses.items()):
                if state == "queued":
                    self.file_statuses[path] = "pending"
                elif state == "running":
                    self.file_statuses[path] = "stopped"
            if stopped is not None and stopped in self.file_statuses:
                self.file_statuses[stopped] = "stopped"
            self.active_path = None
            self.refresh_file_list()
            self.current_file_label.setText(self.text["stopped_status"])
            self.append_log(self.text["cancelled_log"])
            self.set_running(False)
            if self.worker_thread is not None:
                self.worker_thread.quit()
                self.worker_thread = None
            self.worker = None
            self.cleanup_incremental_temp_dirs()
            return
        self.active_path = None
        self.refresh_file_list()
        output_pptx = (
            completed[-1][1] if completed else (results[-1][1] if results else None)
        )
        if output_pptx is not None:
            self.refresh_info_text(output_pptx)
        else:
            self.refresh_info_text()
        self.progress_bar.setValue(100)
        total = len(results) + len(failures)
        success = len(completed)
        failed = len(failures)
        summary = self.text["batch_summary"].format(success=success, total=total)
        if failed:
            summary += " " + self.text["batch_failed_summary"].format(
                failed=failed, total=total
            )
        if skipped:
            summary += " " + self.text["batch_skipped_summary"].format(
                skipped=len(skipped)
            )
        self.current_file_label.setText(summary)
        self.set_running(False)
        if failed:
            self.show_log_drawer(auto_hide=False)
        elif not self.log_drawer.isHidden():
            self.log_drawer_timer.start(2500)
        if self.worker_thread is not None:
            self.worker_thread.quit()
            self.worker_thread = None
        self.worker = None
        self.cleanup_incremental_temp_dirs()
        dialog_body = summary + "\n\n" + self.text["completion_hint"]
        if len(completed) == 1:
            _, output_path, size_bytes, _, _ = completed[0]
            dialog_body += "\n" + self.text["completion_single"].format(
                name=output_path.name,
                size=size_bytes / 1024 / 1024,
            )
        if failed and not success:
            dialog_title = self.text["failed_title"]
        else:
            dialog_title = (
                self.text["batch_done"] if total > 1 else self.text["done_title"]
            )
        auto_audit = getattr(self, "auto_audit_after_compression", False) or (
            bool(completed) and self.auto_optimize_checkbox.isChecked()
        )
        if not auto_audit:
            self.show_dialog(dialog_title, dialog_body)
        else:
            self.auto_audit_after_compression = False
            self.run_audit({item[0] for item in completed})

    def on_failed(self, message: str) -> None:
        self.auto_audit_after_compression = False
        self.current_file_label.setText(self.text["failed_status"])
        self.append_log(f"[ERR] {message}")
        self.set_running(False)
        self.show_log_drawer(auto_hide=False)
        if self.worker_thread is not None:
            self.worker_thread.quit()
            self.worker_thread = None
        self.worker = None
        self.cleanup_incremental_temp_dirs()
        self.show_dialog(self.text["failed_title"], message)

    def resizeEvent(self, event) -> None:  # noqa: N802, ANN001
        super().resizeEvent(event)
        if hasattr(self, "log_drawer"):
            self.position_log_drawer()

    def closeEvent(self, event) -> None:  # noqa: N802, ANN001
        if self.worker is not None and self.is_running:
            try:
                self.worker.cancel()
            except Exception:
                pass
        if hasattr(self, "audit_worker"):
            try:
                self.audit_worker.cancel()
            except Exception:
                pass
        for attr_name in ("worker_thread", "audit_thread"):
            thread = getattr(self, attr_name, None)
            if thread is None or not thread.isRunning():
                continue
            thread.quit()
            thread.wait()
        self.cleanup_incremental_temp_dirs()
        self.cleanup_session_reports()
        super().closeEvent(event)


def main() -> int:
    configure_app_logging()
    app = QApplication([])
    configure_ui_font(app)
    language = detect_language()
    experimental = is_experimental_runtime()
    app.setApplicationName(
        "Doc Media Toolkit Experimental" if experimental else "Doc Media Toolkit"
    )
    app.setApplicationDisplayName(
        (
            "文档媒体工具箱 实验版"
            if language == "zh"
            else "Doc Media Toolkit Experimental"
        )
        if experimental
        else ("文档媒体工具箱" if language == "zh" else "Doc Media Toolkit")
    )
    icon_path = resource_root().joinpath("assets", "app_icon.png")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
