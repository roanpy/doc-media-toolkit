from __future__ import annotations

import sys
import os
import tempfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import (
    Property,
    QEasingCurve,
    QObject,
    QPropertyAnimation,
    QRectF,
    QSettings,
    QSize,
    QThread,
    QTimer,
    Qt,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QColor,
    QFont,
    QIcon,
    QKeySequence,
    QPainter,
    QPixmap,
    QShortcut,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QAbstractButton,
    QAbstractSpinBox,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QTextBrowser,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pptx_output_watermark.process_utils import terminate_active_processes
from pptx_tools import __version__
from pptx_tools.app_logging import configure_app_logging
from pptx_tools.language import detect_language as detect_system_language
from pptx_tools.ui_theme import (
    SHARED_DIALOG_QSS,
    configure_ui_font,
    install_control_help,
)

APP_NAME = "Doc Media Toolkit"


def persistent_library_setting(
    settings: QSettings,
    key: str,
) -> str:
    value = settings.value(key, "", str).strip()
    if not value:
        return ""
    try:
        root = Path(value).expanduser().resolve()
        temp_root = Path(tempfile.gettempdir()).resolve()
    except (OSError, RuntimeError):
        settings.remove(key)
        return ""
    try:
        root.relative_to(temp_root)
    except ValueError:
        pass
    else:
        settings.remove(key)
        return ""
    # A saved cloud/network library can be temporarily unavailable at startup.
    # Keep its path so the next activation can recover without manual selection.
    return str(root)


STRINGS = {
    "zh": {
        "app_name": "文档媒体工具箱",
        "watermark_tab": "水印导出",
        "watermark_eyebrow": "DOC MEDIA TOOLKIT",
        "watermark_title": "文档及媒体水印导出",
        "watermark_subtitle": "批量处理 DOCX / PDF / PPTX / 图片 / 视频，支持水印导出",
        "video_tab": "动态压缩",
        "video_eyebrow": "DOC MEDIA TOOLKIT",
        "video_title": "文档及媒体动态压缩",
        "video_subtitle": "PPTX 媒体压缩，也支持图片、视频、DOCX、PDF 与 XLSX",
        "manager_tab": "视频库",
        "manager_eyebrow": "DOC MEDIA TOOLKIT",
        "manager_title": "PPTX 视频资产库",
        "manager_subtitle": "管理高清源、版本与 PPTX 关联，为压缩文档安全回填可播放视频",
        "image_manager_tab": "图片库",
        "image_manager_eyebrow": "DOC MEDIA TOOLKIT",
        "image_manager_title": "文档图片资产库",
        "image_manager_subtitle": "归档独立图片及 PPTX / DOCX / PDF 内嵌图片，精确去重并辅助归类",
        "help_button": "使用说明",
        "help_window_title": "文档媒体工具箱 使用说明",
        "help_close": "关闭",
        "watermark_help_title": "文档及媒体水印导出说明",
        "watermark_help_body": (
            "1. 拖入或添加一个或多个 DOCX / PDF / PPTX 文件。\n"
            "2. DOCX / PDF 只能导出为 PDF；混合队列里即使全局选择 PPTX，也会按当前形式导出为 PDF。\n"
            "3. 图片清晰度只作用于图片化 PDF / 图片化 PPTX；可编辑模式尽量保留原始图片与文档结构。\n"
            "4. 图片型 PPTX 可开启视频回贴；正式签名包将内置固定源码构建的 FFmpeg 8.1.2。\n"
            "5. Windows 按文件类型优先使用 PowerPoint/Word/WPS 导出 PDF，失败后再使用 LibreOffice；macOS 中 PPTX 可用 Keynote 兜底，DOCX 可用 Pages 兜底。\n"
            "6. 切换功能时会保留已选文件、当前设置和日志。"
        ),
        "video_help_title": "文档及媒体动态压缩说明",
        "video_help_body": (
            "1. 拖入或添加一个或多个 PPTX、图片、视频或 DOCX/PDF/XLSX 文件。\n"
            "2. 视频和图片各有独立预设，可分别选择高保真、平衡/低体积或不压缩。\n"
            "3. 图片保持原格式；JPEG/WebP 调整编码质量，PNG 只做无损优化，低体积目标模式下才允许轻微降采样。\n"
            "4. DOCX/PDF/XLSX 仅压缩内嵌图片，必须显式填写目标大小；PDF 需要 Poppler，DOCX/XLSX 渲染校验需要 LibreOffice。\n"
            "5. 画质评估会按“阈值”逐项执行 SSIM；提档优化只重写低分素材，并覆盖上一版压缩 PPTX。\n"
            "6. 输出文件默认保存在源文件同目录；完整路径见任务日志。\n"
            "7. 切换功能时会保留已选文件、当前设置和日志。"
        ),
        "load_failed": "{title} 加载失败：\n{error}",
    },
    "en": {
        "app_name": "Doc Media Toolkit",
        "watermark_tab": "Watermark",
        "watermark_eyebrow": "DOC MEDIA TOOLKIT",
        "watermark_title": "Document & Media Watermark Export",
        "watermark_subtitle": "Batch process DOCX / PDF / PPTX / images / videos with watermark export",
        "video_tab": "Compression",
        "video_eyebrow": "DOC MEDIA TOOLKIT",
        "video_title": "Document & Media Dynamic Compression",
        "video_subtitle": "PPTX media compression with supporting document, image, and video inputs",
        "manager_tab": "Video Library",
        "manager_eyebrow": "DOC MEDIA TOOLKIT",
        "manager_title": "PPTX Video Library",
        "manager_subtitle": "Archive one source per video and upgrade compact PPTX files by exact hash",
        "image_manager_tab": "Image Library",
        "image_manager_eyebrow": "DOC MEDIA TOOLKIT",
        "image_manager_title": "Document Image Library",
        "image_manager_subtitle": "Archive and organize images from files, PPTX, DOCX, and PDF",
        "help_button": "Help",
        "help_window_title": "Doc Media Toolkit Help",
        "help_close": "Close",
        "watermark_help_title": "Document & Media Watermark Export Help",
        "watermark_help_body": (
            "1. Drop or add one or more DOCX, PDF, or PPTX files.\n"
            "2. DOCX and PDF can only export to PDF. In mixed queues, they still export as PDF even if the global format is PPTX.\n"
            "3. Image quality only applies to image-based PDF / image-based PPTX; editable modes keep original structure and source images where possible.\n"
            "4. Image-based PPTX video reinsertion uses the source-pinned FFmpeg 8.1.2 bundled with signed release packages.\n"
            "5. Windows tries PowerPoint/Word/WPS by source type first, then LibreOffice; on macOS, LibreOffice stays primary while PPTX can fall back to Keynote and DOCX can fall back to Pages.\n"
            "6. If Keynote or Pages is installed but Automation permission is missing, the app opens System Settings instead of misreporting the engine as missing.\n"
            "7. Switching tools keeps selected files, current settings, and logs."
        ),
        "video_help_title": "Document & Media Dynamic Compression Help",
        "video_help_body": (
            "1. Drop or add one or more PPTX, image, video, or DOCX/PDF/XLSX files.\n"
            "2. Video and image presets are independent and can each be set to High, Balanced/Low, or Off.\n"
            "3. Images keep their original format; JPEG/WebP use encoder quality, PNG is losslessly optimized, and low target-size mode may lightly downsample.\n"
            "4. DOCX/PDF/XLSX compression only re-encodes embedded images and requires an explicit target size; PDF needs Poppler, and DOCX/XLSX render validation needs LibreOffice.\n"
            "5. Quality audit uses the Limit field as the SSIM threshold, and optimize rewrites only failed assets into the previous compressed PPTX.\n"
            "6. Output is saved beside the source file by default; full paths are in the log.\n"
            "7. Switching tools keeps selected files, current settings, and logs."
        ),
        "load_failed": "{title} failed to load:\n{error}",
    },
}

HELP_SECTIONS = {
    "zh": [
        (
            "文档及媒体水印导出",
            """
            <h2>文档及媒体水印导出</h2>
            <p>用于把 DOCX / PDF / PPTX 批量导出为 PDF，或把 PPTX 导出为 PDF / PPTX，并按需要添加文字/图片水印。</p>
            <h3>常用流程</h3>
            <ol>
              <li>拖入或点击添加 DOCX / PDF / PPTX，可一次处理多个文件。</li>
              <li>选择输出格式：只有 PPTX 支持导出为 PPTX；DOCX / PDF 只会导出为 PDF。</li>
              <li>选择输出形式：可编辑尽量保留文本/矢量结构，图片化会把页面扁平为整页图片。</li>
              <li>设置水印内容、颜色、透明度、字号、间距、角度，然后点击开始导出。</li>
            </ol>
            <h3>隐藏能力</h3>
            <ul>
              <li><b>图片水印</b>：切换到“图片”水印后可选择 PNG，适合 Logo、印章、保密标识。</li>
              <li><b>图片型 PPTX 保留视频</b>：在 PPTX + 图片化模式下开启“视频保留/回贴”，会提取内嵌视频，按媒体压缩的视频高保真预设给视频画面加同款水印，再按原位置贴回输出 PPTX。</li>
              <li><b>视频回贴失败兜底</b>：如果某些特殊 PPTX 无法回贴，程序会在输出旁生成视频目录和 manifest，避免处理结果丢失。</li>
              <li><b>字体补齐</b>：检测到源 PPTX 显式字体缺失时，可用当前平台的系统字体临时替换副本后导出，不会修改原文件。</li>
              <li><b>图片清晰度</b>：只作用于图片化 PDF / 图片化 PPTX，档位统一为原 / 高 / 平衡 / 低；可编辑模式按原文档保留。</li>
              <li><b>预览</b>：选择文件后会渲染前几页做文字或图片水印预览，大文件可手动刷新。</li>
            </ul>
            """,
        ),
        (
            "文档及媒体动态压缩",
            """
            <h2>文档及媒体动态压缩</h2>
            <p>核心场景是压缩 PPTX 内嵌视频和图片；也接受独立图片、视频及 DOCX/PDF/XLSX，并共用同一队列和设置。</p>
            <h3>常用流程</h3>
            <ol>
              <li>拖入一个或多个 PPTX、图片、视频或 DOCX/PDF/XLSX 文件。</li>
              <li>目标大小留空时按视频预设和图片预设分别处理；填写 MB 时按目标体积压缩。</li>
              <li>视频预设和图片预设可以分别选择不压缩。</li>
              <li>DOCX/PDF/XLSX 只压缩内嵌图片，必须填写目标大小且图片预设不能为“不压缩”。</li>
              <li>点击开始压缩，输出默认保存在源文件同目录。</li>
            </ol>
            <h3>压缩策略</h3>
            <ul>
              <li>会读取视频在每页里的显示面积，面积越大优先保留更高清晰度。</li>
              <li>全屏或接近全屏视频默认不低于 720p，小画面可自动降到 480p。</li>
              <li>默认优先尝试 GPU H.264 编码，不可用或失败时自动切回 CPU libx264。</li>
              <li>旧格式视频会转为 H.264/AAC MP4，并自动重写 PPTX 内部引用。</li>
              <li>图片保持原格式：JPEG/WebP 使用质量参数，PNG 只做无损优化；不支持或压缩后变大的图片会自动保留原图。</li>
              <li>填写目标体积时优先压缩视频，仍不够时再降低图片质量；低体积图片档才允许最多缩到 80% 像素尺寸。</li>
              <li>目标体积安全模式默认使用 CPU 两遍编码；只有手动开启“目标容量用 GPU”才尝试 GPU。</li>
              <li>“阈值”是 SSIM 画质门槛，数值越高越严格；鼠标悬停可查看推荐区间。</li>
              <li>如果视频已经足够小，会直接保留原文件，避免反向变大。</li>
              <li>过程视频默认只放在临时目录；成功或手动停止会清理，普通处理错误时才会在源文件旁保留诊断目录。</li>
            </ul>
            """,
        ),
        (
            "PPTX 视频资产库",
            """
            <h2>PPTX 视频资产库</h2>
            <p>视频库保存高质量视频实体和 PPTX 形状关联，不复制 PPTX。可按客户或主题建立多个独立视频库并随时切换。</p>
            <h3>首次建库与回填</h3>
            <ol>
              <li>新建或打开一个视频库目录；最近使用的库会在下次启动时恢复。</li>
              <li>选择 PPTX，按需要选择“1080p 高质量”或“保留原片”，可填写客户/年份/项目分类，再归档高清视频。</li>
              <li>外部视频可通过按钮或拖入列表批量导入：完全相同则复用，唯一内容指纹命中则加入已有视频族；无可靠唯一结果时显示双方封面、播放入口和相似度候选，由你选择已有视频族、新建或跳过。</li>
              <li>压缩页开启入库后，会先归档高清源，再登记压缩后媒体哈希；低清压缩文件不会在库内重复保存。</li>
              <li>源视频归档默认关闭，并非普通压缩的前置条件；只有“覆盖原 PPTX”才要求先归档视频源。</li>
              <li>选择压缩版或改名后的 PPTX，点击“高清回填 PPTX（另存）”并指定输出位置。未匹配媒体可人工选择候选并决定是否记住低清哈希；输入文件不会被覆盖。</li>
            </ol>
            <h3>人工管理</h3>
            <ul>
              <li>顶部数字展示视频族、版本、待核对、无关联、多版本和文件异常数量，点击数字可直接筛选。“待核对”是后三类的去重并集；无关联和多版本通常不是错误，文件异常才表示库内文件不可用。</li>
              <li><b>库体检</b>：默认只读检查视频实体、哈希唯一归属、PPTX 来源与形状关联、历史输出和待清理索引；完整复核会逐个重算 SHA-256，并区分内容变化与仅时间戳变化。</li>
              <li>可查找名称、路径和哈希，按状态筛选，并点击任一表头排序。无结果时会提示清除查找或恢复“全部视频”。</li>
              <li>选择条目后可播放、重命名、移动、添加版本或设为高清源；“查找丢失”可恢复 Finder 中改名或移动后的文件。</li>
              <li>筛出无关联视频后可点击“归并视频”，按候选封面和相似度人工归族；文件异常版本只有在非高清源、非当前版本且零引用时才能隔离。</li>
              <li>“整理视频库”只生成候选并要求确认；落选文件先移入 <code>_cleanup/</code> 隔离区，可还原，不会直接删除。</li>
            </ul>
            <h3>匹配与安全边界</h3>
            <ul>
              <li>文件名和目录不是身份标识。先用 SHA-256 精确匹配；不同压缩或分辨率版本只有在时长、宽高比、五帧画面和音频指纹均严格且唯一匹配时才归为同族。</li>
              <li>不同音轨、被裁剪、多个候选或置信度不足时不会自动归并或回填；单独的时长、分辨率和文件名都不能证明同源。</li>
              <li>高清优化保留原视频形状、海报、尺寸和播放时间线，只替换包内媒体。</li>
              <li>WMV、AVI 等原片可按原格式保存；回填时按需临时生成不放大的 PowerPoint 兼容 MP4。</li>
            </ul>
            <h3>文件与日志</h3>
            <ul>
              <li>每个视频库目录包含 <code>video-project.json</code>、<code>media/</code>、<code>_cleanup/</code> 和 <code>reports/</code>。</li>
              <li>最近目录等偏好由系统 QSettings 保存；“操作记录”只显示本次运行，长期滚动日志位于“日志目录”。</li>
            </ul>
            """,
        ),
        (
            "文档图片资产库",
            """
            <h2>文档图片资产库</h2>
            <p>保存独立图片以及 PPTX、DOCX、数字版 PDF 中实际嵌入的图片。图片库不修改原文档，也不负责图片回填。</p>
            <h3>归档与判重</h3>
            <ol>
              <li>新建或打开图片库，导入图片或文档，也可直接拖入。</li>
              <li>先按 SHA-256 精确复用相同文件；不同编码的图片只生成代码指纹相似候选，不会自动合并。</li>
              <li>PPTX / DOCX 只提取关系文件实际引用的图片；PDF 只提取内嵌图像，不整页渲染、不做 OCR。</li>
              <li>图片和视频使用相同的跨平台分类校验与保守名称清洗；原文件名和来源路径仍保留在清单中。</li>
              <li>名称、分类、标签和说明可以人工维护；配置视觉 AI 后可生成整理建议，但应用前仍需确认。</li>
            </ol>
            <h3>空间控制</h3>
            <ul>
              <li>相同内容在库中只保存一份，多个来源记录到同一图片条目。</li>
              <li>图片库不生成永久缩略图缓存；“清理未引用文件”只处理清单未引用的库内文件。</li>
              <li>移除、合并和清理会先把库内副本移入待清理区，可还原；再次确认永久清空后才会删除，不影响原始图片或文档。</li>
            </ul>
            """,
        ),
        (
            "依赖与平台",
            """
            <h2>依赖与平台说明</h2>
            <h3>macOS</h3>
            <ul>
              <li>PPTX 的 PDF/图片化导出优先使用 LibreOffice，缺失或短时不可用时回退 Keynote；DOCX 优先使用 LibreOffice，缺失或短时不可用时回退 Pages。这些外部应用都不内置。</li>
              <li>如果 Keynote 或 Pages 已安装但没有 Automation 权限，界面会提示打开系统设置，不再混淆成“未安装”。</li>
              <li>正式签名包将内置由固定 FFmpeg 8.1.2、x264 与 zlib 源码构建的 FFmpeg/FFprobe；同一 Release 提供对应源码、构建记录和哈希。</li>
              <li>水印导出的临时目录也会在下次启动时回收异常退出残留。</li>
            </ul>
            <h3>Windows</h3>
            <ul>
              <li>PPTX 水印导出优先使用 Microsoft PowerPoint/WPS COM；DOCX 优先使用 Microsoft Word/WPS COM；最后回退 LibreOffice。</li>
              <li>WPS 如果未注册 COM，可先打开一次 WPS 或修复安装；也可设置 <code>PPTX_TOOLS_WPP</code> 指向 <code>wpp.exe</code>。</li>
              <li>Windows 正式包使用可替换的 portable onedir 结构，内置 FFmpeg 8.1.2，并保留 Media Foundation 硬件编码和 libx264 CPU 回退。</li>
            </ul>
            <h3>CLI</h3>
            <ul>
              <li><code>pptx-tools watermark ...</code> 使用水印导出子命令。</li>
              <li><code>pptx-tools compact ...</code> 使用媒体压缩子命令。</li>
              <li>统一环境变量优先使用 <code>PPTX_TOOLS_FFMPEG</code>、<code>PPTX_TOOLS_FFPROBE</code>、<code>PPTX_TOOLS_SOFFICE</code>、<code>PPTX_TOOLS_WPP</code>。</li>
            </ul>
            """,
        ),
    ],
    "en": [
        (
            "Document & Media Watermark Export",
            """
            <h2>Document &amp; Media Watermark Export</h2>
            <p>Batch export DOCX / PDF / PPTX to PDF, or export PPTX to PDF / PPTX with optional text or image watermarks.</p>
            <h3>Workflow</h3>
            <ol>
              <li>Drop or add one or more DOCX, PDF, or PPTX files.</li>
              <li>Choose the output format. Only PPTX sources can export as PPTX; DOCX and PDF always export as PDF.</li>
              <li>Choose editable or image mode.</li>
              <li>Set watermark content, color, opacity, size, spacing, and angle.</li>
              <li>Click export. Outputs are saved beside the source file by default.</li>
            </ol>
            <h3>Advanced Features</h3>
            <ul>
              <li><b>Image watermark</b>: use a PNG logo, stamp, or confidential mark.</li>
              <li><b>Keep videos in image PPTX</b>: extracts embedded videos, uses the media compression video High preset while applying the same watermark to video frames, and reinserts them at the original positions.</li>
              <li><b>Video fallback</b>: if reinsertion fails, watermarked videos and a manifest are saved beside the output.</li>
              <li><b>Font replacement</b>: missing source fonts can be replaced with platform system fonts in a temporary copy without modifying the original PPTX.</li>
              <li><b>Image quality</b>: only affects image-based PDF / image-based PPTX. Editable modes preserve the original structure and source images where possible.</li>
            </ul>
            """,
        ),
        (
            "Document & Media Dynamic Compression",
            """
            <h2>Document &amp; Media Dynamic Compression</h2>
            <p>The core workflow compresses embedded PPTX video and images. Standalone images/videos and DOCX/PDF/XLSX inputs share the same queue and settings.</p>
            <h3>Workflow</h3>
            <ol>
              <li>Drop one or more PPTX, image, video, or DOCX/PDF/XLSX files.</li>
              <li>Leave target size blank for video and image presets, or enter MB for target-size mode.</li>
              <li>Video and image presets can each be set to Off.</li>
              <li>DOCX/PDF/XLSX only re-encode embedded images and require a target size plus an enabled image preset.</li>
              <li>Click compress. Output is saved beside the source file.</li>
            </ol>
            <h3>Strategy</h3>
            <ul>
              <li>Reads each video's on-slide display size to allocate quality.</li>
              <li>Large videos stay at least 720p; small videos may use 480p.</li>
              <li>Auto encoding tries GPU H.264 first and falls back to CPU libx264.</li>
              <li>Legacy formats are converted to H.264/AAC MP4 and internal PPTX references are updated.</li>
              <li>Images keep their original format: JPEG/WebP use encoder quality, PNG is optimized losslessly, and unsupported or larger results are copied unchanged.</li>
              <li>Target-size mode compresses videos first, then lowers image quality only if needed; Image Low may downsample to 80% pixel dimensions.</li>
              <li>Safe target-size mode uses CPU two-pass encoding unless “GPU for target size” is explicitly enabled.</li>
              <li>The Limit field is the SSIM quality threshold; hover it to see the suggested ranges.</li>
              <li>Process videos stay in a temp folder by default; success and manual stops clean them, while ordinary failures preserve a diagnostic folder beside the source.</li>
            </ul>
            """,
        ),
        (
            "PPTX Video Library",
            """
            <h2>PPTX Video Library</h2>
            <p>Store high-quality video files and PPTX shape associations without copying the PPTX files. Separate libraries can be created for different clients or topics.</p>
            <h3>Archive and restore</h3>
            <ol>
              <li>Create or switch to a library folder. The most recent library is reopened on the next launch.</li>
              <li>Select PPTX files, choose high-quality 1080p or original preservation, optionally enter a category, then archive their sources.</li>
              <li>External videos can be selected or dropped onto the list. Exact and unique strict matches are automatic; otherwise a cover-frame, playback, and ranked-candidate review lets you link, create, or skip.</li>
              <li>Compression can archive the source first and register compact media hashes as aliases without keeping duplicate low-quality files.</li>
              <li>Use “Upgrade PPTX” on a compact or renamed deck and choose a separate output path. Unresolved media can be linked manually with optional hash remembering. The input is not overwritten.</li>
            </ol>
            <h3>Review and safety</h3>
            <ul>
              <li>Review counts combine unlinked, multi-version, and abnormal families; they are review hints, not proof of an error. Search, filter, and click any table header to sort.</li>
              <li>Identity does not depend on filenames or folders. Matching starts with SHA-256, then requires strict duration, aspect-ratio, five-frame visual, and audio fingerprints to produce one candidate.</li>
              <li>Different audio, trimming, ambiguity, or low confidence prevents automatic merging and replacement. Duration or resolution alone never establishes identity.</li>
              <li>Upgraded decks keep original shapes, poster frames, geometry, and playback timing.</li>
              <li>Cleanup first moves files to the library <code>_cleanup/</code> quarantine and supports restore; it never deletes candidates immediately.</li>
              <li>Manual family merge uses the same visual candidate review. An unreadable version can only be isolated when it is unused and is neither the source nor active version.</li>
              <li>Each library owns its manifest, media, quarantine, and reports. Preferences and rotating application logs remain in the operating system application-data locations.</li>
            </ul>
            """,
        ),
        (
            "Document Image Library",
            """
            <h2>Document Image Library</h2>
            <p>Archive standalone images and embedded images referenced by PPTX, DOCX, and digital PDF files. The library does not modify or restore source documents.</p>
            <h3>Archive and deduplicate</h3>
            <ol>
              <li>Create or open an image library, then select or drop images and documents.</li>
              <li>Exact SHA-256 matches reuse one stored file. Visual fingerprints only create review candidates and never merge automatically.</li>
              <li>Office packages import referenced media only. PDFs import embedded images without page rendering or OCR.</li>
              <li>Image and video imports share portable category validation and conservative name cleanup; original names and source paths remain in the manifest.</li>
              <li>Names, categories, tags, and summaries remain editable. Optional vision AI produces reviewable suggestions only.</li>
            </ol>
            <h3>Storage control</h3>
            <ul>
              <li>Identical bytes are stored once with multiple source records.</li>
              <li>No permanent thumbnail cache is created. Cleanup removes only files absent from the manifest.</li>
              <li>Removing an item deletes the library copy without touching source files or documents.</li>
            </ul>
            """,
        ),
        (
            "Platforms",
            """
            <h2>Dependencies and Platforms</h2>
            <h3>macOS</h3>
            <ul>
              <li>PPTX PDF/image export prefers LibreOffice and falls back to Keynote when needed. DOCX prefers LibreOffice and falls back to Pages. External engines are not bundled.</li>
              <li>If Keynote or Pages is installed but Automation permission is missing, the UI should offer a System Settings action instead of a download link.</li>
              <li>Signed release packages bundle FFmpeg/FFprobe built from pinned FFmpeg 8.1.2, x264, and zlib source. The same Release carries corresponding source, build records, and hashes.</li>
              <li>Watermark temp roots are also reaped on the next launch after abnormal termination.</li>
            </ul>
            <h3>Windows</h3>
            <ul>
              <li>Watermark export tries PowerPoint COM, then WPS COM, then LibreOffice fallback.</li>
              <li>If WPS COM is not registered, open/repair WPS or set <code>PPTX_TOOLS_WPP</code> to <code>wpp.exe</code>.</li>
              <li>The public Windows package uses a replaceable portable onedir layout with Media Foundation hardware encoding and libx264 CPU fallback.</li>
            </ul>
            <h3>CLI</h3>
            <ul>
              <li><code>pptx-tools watermark ...</code> runs watermark export.</li>
              <li><code>pptx-tools compact ...</code> runs media compression.</li>
              <li>Preferred environment variables: <code>PPTX_TOOLS_FFMPEG</code>, <code>PPTX_TOOLS_FFPROBE</code>, <code>PPTX_TOOLS_SOFFICE</code>, <code>PPTX_TOOLS_WPP</code>.</li>
            </ul>
            """,
        ),
    ],
}

HELP_EXTRA_SECTIONS = {
    "zh": {
        "快速开始": """
            <h2>快速开始</h2>
            <p>先按目标选择工具，再添加文件。切换工具时会保留当前文件、设置和本次日志。</p>
            <h3>我应该用哪个工具？</h3>
            <table cellspacing="0" cellpadding="6">
              <tr><td><b>水印导出</b></td><td>给 DOCX、PDF、PPTX、图片或视频添加水印并导出。</td></tr>
              <tr><td><b>动态压缩</b></td><td>压缩 PPTX 内嵌媒体，也接受独立媒体和 DOCX/PDF/XLSX 内嵌图片。</td></tr>
              <tr><td><b>视频资产库</b></td><td>管理视频版本、哈希与 PPTX 关联，并进行高清回填。</td></tr>
              <tr><td><b>图片资产库</b></td><td>归档独立图片及文档内实际引用的图片，不做图片回填。</td></tr>
            </table>
            <h3>推荐顺序</h3>
            <ol>
              <li>从顶部选择工具，拖入文件或点击选择按钮。</li>
              <li>先查看默认设置和预计输出，再启动处理。</li>
              <li>运行时查看底部“状态与日志”；失败时日志会自动展开。</li>
            </ol>
            <h3>安全原则</h3>
            <ul>
              <li>默认另存输出，不覆盖源文件；明确勾选覆盖时会再次确认。</li>
              <li>相似候选和 AI 结果都只是建议，归并、删除和回填前仍需人工确认。</li>
              <li>库维护先移入待清理区，避免直接删除仍有关联的资源。</li>
            </ul>
        """,
        "AI 辅助": """
            <h2>AI 辅助</h2>
            <p>AI 用于整理建议，不替代代码判重，也不会自动修改视频库或图片库。</p>
            <h3>配置与验证</h3>
            <ol>
              <li>点击顶栏齿轮，填写 OpenAI 兼容的 Base URL、模型和 API Key。</li>
              <li>点击“测试 AI 连接”；勾选图片输入时会额外验证视觉能力。</li>
              <li>验证结果会区分文本可用、视觉可用、视觉不支持和暂时无法判断。</li>
            </ol>
            <h3>在资源库中能做什么？</h3>
            <ul>
              <li>视频库：建议同源归并、主视频、名称和分类；视觉模式可参考三帧联系图。</li>
              <li>图片库：建议命名、分类、标签、说明和可能的合并组。</li>
              <li>关闭图片输入后只发送名称、分类、标签、说明、规格、大小和代码相似度，不发送预览图片。</li>
            </ul>
            <h3>隐私与安全</h3>
            <ul>
              <li>API Key 只保留在本次运行内，不写入磁盘或日志。</li>
              <li>不会发送本地路径、原始 PPTX/DOCX/PDF 或完整视频。</li>
              <li>AI 建议必须在确认窗口中选择后才会应用。</li>
            </ul>
        """,
        "日志、安全与恢复": """
            <h2>日志、安全与恢复</h2>
            <p>底部状态条显示当前结果；完整日志按需展开，不长期挤占主界面。</p>
            <h3>查看日志</h3>
            <ul>
              <li>鼠标停留状态条约 1 秒后展开；移开后自动收起。</li>
              <li>任务运行或出现错误时会自动展开；运行中的日志不会自动隐藏。</li>
              <li>界面显示本次运行摘要，长期滚动日志可从“日志目录”打开。</li>
            </ul>
            <h3>失败与恢复</h3>
            <ul>
              <li>普通导出和压缩默认另存，源文件保持不变。</li>
              <li>视频库和图片库整理先移动到 <code>_cleanup/</code>，可在“待清理”中恢复。</li>
              <li>清单写入使用版本校验与失败回滚，避免多窗口同时覆盖。</li>
              <li>异常退出留下的运行时临时目录会在下次启动时清理。</li>
            </ul>
            <h3>需要人工确认的操作</h3>
            <ul>
              <li>覆盖源 PPTX、移除图片、清空待清理文件、跨族归并和应用 AI 建议。</li>
              <li>文件名、目录、时长或分辨率单独都不能证明两个资源相同。</li>
            </ul>
        """,
        "开源、语言与隐私": """
            <h2>开源、语言与隐私</h2>
            <ul>
              <li>当前版本：<code>{version}</code>。</li>
              <li>项目源码采用 MIT 许可证，仓库地址为 <code>github.com/roanpy/doc-media-toolkit</code>；第三方组件仍适用各自许可证。</li>
              <li>项目由所有者主导，主要使用 OpenAI Codex 协助开发，并由 Google Gemini 与 Anthropic Claude Code 辅助和交叉核验；所有产出均由项目所有者审核、测试和维护。</li>
              <li>正式包会在 <code>licenses/</code> 内携带项目、Python、Qt、资源和 Python 依赖许可；内置 FFmpeg/x264 适用 GPL，并在同一 Release 提供对应源码和构建证据。</li>
              <li>当前公开源码可直接使用；预构建候选在完成平台签名和公证前保持 Draft，不应绕过系统安全检查。</li>
              <li>主窗口、水印、压缩和帮助中心支持简体中文与英文；视频库和图片库业务界面当前仍以中文为主。</li>
              <li>支持双语的工作区首次启动跟随系统界面语言；可在启动前设置 <code>PPTX_TOOLS_LANG=zh</code> 或 <code>PPTX_TOOLS_LANG=en</code> 覆盖。</li>
              <li>API Key 只保留在本次运行的内存中。不开启图片输入时不会向 AI 发送预览图；不会发送完整文档、完整视频或本机路径。</li>
              <li>普通问题可提交 GitHub Issue；安全漏洞请使用仓库 Security 页的私密报告入口，不要公开披露。</li>
            </ul>
        """,
    },
    "en": {
        "Quick Start": """
            <h2>Quick Start</h2>
            <p>Choose a tool by outcome, then add files. Switching tools keeps the current files, settings, and session log.</p>
            <h3>Choose a tool</h3>
            <table cellspacing="0" cellpadding="6">
              <tr><td><b>Watermark</b></td><td>Add watermarks to documents and media, then export.</td></tr>
              <tr><td><b>Compression</b></td><td>Compress PPTX media plus standalone media and embedded images in DOCX/PDF/XLSX.</td></tr>
              <tr><td><b>Video Library</b></td><td>Manage video versions, hashes, PPTX links, and high-quality restoration.</td></tr>
              <tr><td><b>Image Library</b></td><td>Archive standalone and referenced document images without image restoration.</td></tr>
            </table>
            <h3>Recommended order</h3>
            <ol>
              <li>Choose a tool and add or drop files.</li>
              <li>Review defaults and expected output before starting.</li>
              <li>Use the bottom status shelf while running; failures open the log automatically.</li>
            </ol>
            <h3>Safety defaults</h3>
            <ul>
              <li>Outputs are saved separately unless overwrite is explicitly selected and confirmed.</li>
              <li>Similarity and AI results are suggestions; merge, delete, and restore actions require review.</li>
              <li>Library cleanup quarantines files before permanent removal.</li>
            </ul>
        """,
        "AI Assistance": """
            <h2>AI Assistance</h2>
            <p>AI provides organization suggestions. It does not replace deterministic matching or modify libraries automatically.</p>
            <h3>Configure and verify</h3>
            <ol>
              <li>Open the gear menu and enter an OpenAI-compatible Base URL, model, and API key.</li>
              <li>Test the connection. Enabling image input also probes vision support.</li>
              <li>Results distinguish text support, vision support, unsupported vision, and unknown availability.</li>
            </ol>
            <h3>Library assistance</h3>
            <ul>
              <li>Video: merge candidates, primary source, naming, and category suggestions.</li>
              <li>Image: naming, category, tags, summary, and possible merge groups.</li>
              <li>With image input off, only names, categories, tags, summaries, specifications, sizes, and code similarity are sent; previews are not sent.</li>
            </ul>
            <h3>Privacy</h3>
            <ul>
              <li>The API key stays in memory for the current run and is not written to disk or logs.</li>
              <li>Local paths, source documents, and full videos are not sent.</li>
              <li>Suggestions apply only after explicit confirmation.</li>
            </ul>
        """,
        "Logs, Safety, and Recovery": """
            <h2>Logs, Safety, and Recovery</h2>
            <p>The bottom shelf shows current status. Full logs expand only when needed.</p>
            <h3>Logs</h3>
            <ul>
              <li>Hover the status shelf for about one second to expand it.</li>
              <li>Running tasks and errors open it automatically; active logs stay visible.</li>
              <li>The UI shows the current session. Rotating logs are available from the log folder.</li>
            </ul>
            <h3>Recovery</h3>
            <ul>
              <li>Export and compression save separate outputs by default.</li>
              <li>Video cleanup moves files to <code>_cleanup/</code> for review and restore.</li>
              <li>Manifest revision checks and rollback protect concurrent updates.</li>
              <li>Stale runtime folders are cleaned on the next launch.</li>
            </ul>
        """,
        "Open Source, Language, and Privacy": """
            <h2>Open Source, Language, and Privacy</h2>
            <ul>
              <li>Current version: <code>{version}</code>.</li>
              <li>The project source is MIT-licensed at <code>github.com/roanpy/doc-media-toolkit</code>. Third-party components retain their own licenses.</li>
              <li>The owner leads the project with primary development assistance from OpenAI Codex and additional generation and cross-checking from Google Gemini and Anthropic Claude Code. The owner reviews, tests, and maintains all output.</li>
              <li>Release packages carry project, Python, Qt, asset, and Python dependency notices under <code>licenses/</code>. Bundled FFmpeg/x264 is GPL-covered, with corresponding source and build evidence on the same Release.</li>
              <li>The public source is usable now. Binary candidates remain Draft until platform signing and notarization are complete; do not bypass operating-system security checks.</li>
              <li>The application shell, watermark, compression, and help center support Simplified Chinese and English. Video and image library operations are currently Chinese-first.</li>
              <li>Bilingual workspaces follow the system UI language on first launch. Set <code>PPTX_TOOLS_LANG=zh</code> or <code>PPTX_TOOLS_LANG=en</code> before launch to override it.</li>
              <li>API keys stay in process memory. AI previews are sent only when image input is enabled; complete documents, complete videos, and local paths are not sent.</li>
              <li>Use GitHub Issues for ordinary problems. Report vulnerabilities privately from the repository Security page.</li>
            </ul>
        """,
    },
}

HELP_TOPIC_GROUPS = {
    "zh": [
        ("开始使用", ["快速开始"]),
        (
            "四个工具",
            [
                "文档及媒体水印导出",
                "文档及媒体动态压缩",
                "PPTX 视频资产库",
                "文档图片资产库",
            ],
        ),
        (
            "通用能力",
            ["AI 辅助", "日志、安全与恢复", "依赖与平台", "开源、语言与隐私"],
        ),
    ],
    "en": [
        ("Get Started", ["Quick Start"]),
        (
            "Tools",
            [
                "Document & Media Watermark Export",
                "Document & Media Dynamic Compression",
                "PPTX Video Library",
                "Document Image Library",
            ],
        ),
        (
            "Common",
            [
                "AI Assistance",
                "Logs, Safety, and Recovery",
                "Platforms",
                "Open Source, Language, and Privacy",
            ],
        ),
    ],
}


def help_topics(language: str) -> dict[str, str]:
    topics = dict(HELP_SECTIONS[language])
    extras = HELP_EXTRA_SECTIONS[language]
    return {
        title: extras.get(title, topics.get(title, "")).replace(
            "{version}", __version__
        )
        for _group, titles in HELP_TOPIC_GROUPS[language]
        for title in titles
    }


def detect_language() -> str:
    return detect_system_language("PPTX_TOOLS_LANG")


def resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    source_root = Path(__file__).resolve().parents[2]
    return source_root if (source_root / "assets").is_dir() else Path(sys.prefix)


def resource_path(*parts: str) -> Path:
    return resource_root().joinpath(*parts)


@dataclass(slots=True)
class EmbeddedTool:
    title: str
    factory: Callable[[], QMainWindow] | None
    window: QMainWindow | None
    widget: QWidget
    error: Exception | None = None


class HelpDialog(QDialog):
    def __init__(
        self, language: str, current_index: int, parent: QWidget | None = None
    ) -> None:
        super().__init__(parent)
        text = STRINGS[language]
        self.language = language
        self.topics = help_topics(language)
        self.topic_items: dict[str, QTreeWidgetItem] = {}
        self.topic_search: dict[str, str] = {}
        self.setWindowTitle(text["help_window_title"])
        self.setMinimumSize(760, 540)
        self.resize(980, 680)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 16)
        root.setSpacing(14)

        eyebrow = QLabel("DOC MEDIA TOOLKIT · HELP")
        eyebrow.setObjectName("helpEyebrow")
        root.addWidget(eyebrow)
        title = QLabel("帮助中心" if language == "zh" else "Help Center")
        title.setObjectName("helpTitle")
        root.addWidget(title)
        subtitle = QLabel(
            "按主题浏览四个工具、AI、安全与排错；也可以直接搜索。"
            if language == "zh"
            else "Browse tools, AI, safety, and troubleshooting by topic or search."
        )
        subtitle.setObjectName("helpSubtitle")
        root.addWidget(subtitle)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("helpSplitter")
        splitter.setChildrenCollapsible(False)

        navigation = QFrame()
        navigation.setObjectName("helpNavigation")
        navigation_layout = QVBoxLayout(navigation)
        navigation_layout.setContentsMargins(10, 10, 10, 10)
        navigation_layout.setSpacing(8)
        self.search_input = QLineEdit()
        self.search_input.setObjectName("helpSearch")
        self.search_input.setPlaceholderText(
            "搜索主题或内容" if language == "zh" else "Search topics or content"
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setAccessibleName(
            "搜索帮助" if language == "zh" else "Search help"
        )
        self.search_input.textChanged.connect(self._filter_topics)
        navigation_layout.addWidget(self.search_input)

        self.topic_tree = QTreeWidget()
        self.topic_tree.setObjectName("helpTopicTree")
        self.topic_tree.setHeaderHidden(True)
        self.topic_tree.setIndentation(10)
        self.topic_tree.setRootIsDecorated(False)
        self.topic_tree.setItemsExpandable(False)
        self.topic_tree.setAccessibleName(
            "帮助主题" if language == "zh" else "Help topics"
        )
        navigation_layout.addWidget(self.topic_tree, 1)
        navigation_hint = QLabel(
            "↑ ↓ 选择主题　Enter 打开"
            if language == "zh"
            else "Use ↑ ↓ and Enter to navigate"
        )
        navigation_hint.setObjectName("helpNavigationHint")
        navigation_layout.addWidget(navigation_hint)
        splitter.addWidget(navigation)

        content = QFrame()
        content.setObjectName("helpContent")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 12, 10, 10)
        content_layout.setSpacing(8)
        self.context_label = QLabel()
        self.context_label.setObjectName("helpContext")
        content_layout.addWidget(self.context_label)
        self.body = QTextBrowser()
        self.body.setObjectName("helpBody")
        self.body.setOpenExternalLinks(False)
        self.body.setAccessibleName("帮助正文" if language == "zh" else "Help content")
        self.body.document().setDefaultStyleSheet(
            """
            body { color: #cbd5e1; font-size: 13px; line-height: 1.5; }
            h2 { color: #f8fafc; font-size: 24px; margin: 4px 0 10px; }
            h3 { color: #fb923c; font-size: 15px; margin: 18px 0 6px; }
            p { margin: 0 0 10px; }
            ul, ol { margin: 4px 0 12px 22px; }
            li { margin-bottom: 6px; }
            table { width: 100%; border-collapse: collapse; margin: 6px 0 14px; }
            td { border-bottom: 1px solid #26394d; padding: 8px; }
            code { color: #dbeafe; background-color: #16283a; }
            """
        )
        content_layout.addWidget(self.body, 1)
        splitter.addWidget(content)
        splitter.setSizes([220, 700])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        root.addWidget(splitter, 1)

        for group_name, topic_titles in HELP_TOPIC_GROUPS[language]:
            group_item = QTreeWidgetItem([group_name])
            group_item.setFlags(group_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            group_font = group_item.font(0)
            group_font.setBold(True)
            group_item.setFont(0, group_font)
            group_item.setForeground(0, QColor("#71849a"))
            self.topic_tree.addTopLevelItem(group_item)
            group_item.setExpanded(True)
            for topic_title in topic_titles:
                item = QTreeWidgetItem([topic_title])
                item.setData(0, Qt.ItemDataRole.UserRole, topic_title)
                group_item.addChild(item)
                self.topic_items[topic_title] = item
                document = QTextDocument()
                document.setHtml(self.topics[topic_title])
                self.topic_search[topic_title] = (
                    topic_title + " " + document.toPlainText()
                ).casefold()

        self.topic_tree.currentItemChanged.connect(self._topic_changed)
        selected_title = (
            HELP_SECTIONS[language][current_index][0]
            if 0 <= current_index < len(HELP_SECTIONS[language])
            else next(iter(self.topics))
        )
        self._select_topic(selected_title)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_button = buttons.button(QDialogButtonBox.StandardButton.Close)
        if close_button is not None:
            close_button.setText(text["help_close"])
            close_button.setObjectName("helpCloseButton")
            close_button.setFixedSize(96, 36)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self.setStyleSheet(
            """
            QDialog {
                background: #07101a;
            }
            QLabel#helpEyebrow {
                color: #f97316;
                font-size: 11px;
                font-weight: 700;
                letter-spacing: 1px;
            }
            QLabel#helpTitle {
                color: #f8fafc;
                font-size: 22px;
                font-weight: 700;
            }
            QLabel#helpSubtitle {
                color: #8fa2b8;
                font-size: 13px;
            }
            QFrame#helpNavigation, QFrame#helpContent {
                background: #0d1926;
                border: 1px solid #294057;
                border-radius: 10px;
            }
            QLineEdit#helpSearch {
                min-height: 34px;
                padding: 0 10px;
                color: #dce7f3;
                background: #091522;
                border: 1px solid #294057;
                border-radius: 8px;
                font-size: 13px;
            }
            QLineEdit#helpSearch:focus {
                border-color: #f97316;
            }
            QTreeWidget#helpTopicTree {
                background: transparent;
                border: 0;
                color: #c6d3e2;
                font-size: 13px;
                outline: 0;
            }
            QTreeWidget#helpTopicTree::item {
                min-height: 32px;
                padding: 0 8px;
                border-radius: 6px;
            }
            QTreeWidget#helpTopicTree::item:hover {
                background: #122235;
            }
            QTreeWidget#helpTopicTree::item:selected {
                background: #193047;
                color: #ffffff;
                border-left: 3px solid #f97316;
            }
            QLabel#helpNavigationHint {
                color: #63778d;
                font-size: 11px;
            }
            QLabel#helpContext {
                color: #71849a;
                font-size: 12px;
                font-weight: 600;
            }
            QTextBrowser#helpBody {
                background: transparent;
                color: #cbd5e1;
                border: 0;
                padding: 2px 8px 8px 2px;
                font-size: 13px;
            }
            QPushButton#helpCloseButton {
                background: #122235;
                color: #f8fafc;
                border: 1px solid #35506a;
                border-radius: 8px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton#helpCloseButton:hover {
                background: #193047;
            }
            QScrollBar:vertical, QScrollBar:horizontal {
                background: transparent;
                border: 0;
                border-radius: 5px;
            }
            QScrollBar:vertical { width: 10px; margin: 4px 2px; }
            QScrollBar:horizontal { height: 10px; margin: 2px 4px; }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: #334155;
                border-radius: 5px;
                min-height: 28px;
                min-width: 28px;
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
            """
            + SHARED_DIALOG_QSS
        )

    def _select_topic(self, title: str) -> None:
        item = self.topic_items.get(title) or next(iter(self.topic_items.values()))
        self.topic_tree.setCurrentItem(item)

    def _topic_changed(
        self,
        current: QTreeWidgetItem | None,
        _previous: QTreeWidgetItem | None,
    ) -> None:
        if current is None:
            return
        title = current.data(0, Qt.ItemDataRole.UserRole)
        if not title:
            return
        parent = current.parent()
        group = parent.text(0) if parent is not None else ""
        self.context_label.setText(f"{group}　/　{title}")
        self.body.setHtml(self.topics[title])
        self.body.verticalScrollBar().setValue(0)

    def _filter_topics(self, query: str) -> None:
        needle = query.strip().casefold()
        first_visible: QTreeWidgetItem | None = None
        for group_index in range(self.topic_tree.topLevelItemCount()):
            group = self.topic_tree.topLevelItem(group_index)
            group_visible = False
            for child_index in range(group.childCount()):
                child = group.child(child_index)
                title = child.data(0, Qt.ItemDataRole.UserRole)
                visible = not needle or needle in self.topic_search[title]
                child.setHidden(not visible)
                group_visible |= visible
                if visible and first_visible is None:
                    first_visible = child
            group.setHidden(not group_visible)
        current = self.topic_tree.currentItem()
        if first_visible is not None and (current is None or current.isHidden()):
            self.topic_tree.setCurrentItem(first_visible)


class AIConnectionWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, config: object) -> None:
        super().__init__()
        self.config = config

    @Slot()
    def run(self) -> None:
        from pptx_tools.ai_client import AIClientError, OpenAICompatibleClient

        try:
            client = OpenAICompatibleClient(self.config)
            message = client.test_connection()
            if self.config.vision_enabled:
                vision, vision_detail = client.probe_vision_support()
            else:
                vision, vision_detail = (
                    None,
                    "未发送测试图片；勾选图片预览后可验证视觉能力",
                )
        except AIClientError as exc:
            self.failed.emit(str(exc))
        else:
            self.finished.emit(
                {
                    "message": message,
                    "vision": vision,
                    "vision_detail": vision_detail,
                }
            )


class SettingsDialog(QDialog):
    def __init__(
        self,
        language: str,
        session_api_key: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.language = language
        self.settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        self.test_thread: QThread | None = None
        self.test_worker: AIConnectionWorker | None = None
        self.setWindowTitle("全局设置" if language == "zh" else "Settings")
        self.setMinimumWidth(760)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(12)

        title = QLabel("AI 与资源库设置" if language == "zh" else "AI and libraries")
        title.setObjectName("dialogTitle")
        root.addWidget(title)

        ai_box = QFrame()
        ai_box.setObjectName("settingsSection")
        ai_form = QFormLayout(ai_box)
        ai_form.setContentsMargins(12, 12, 12, 12)
        ai_form.setFieldGrowthPolicy(
            QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow
        )
        self.base_url_input = QLineEdit(
            self.settings.value("ai/base_url", "https://api.openai.com/v1", str)
        )
        self.model_input = QLineEdit(self.settings.value("ai/model", "", str))
        self.api_key_input = QLineEdit(session_api_key)
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_input.setPlaceholderText(
            "仅本次运行；也可设置 DOC_MEDIA_AI_API_KEY"
            if language == "zh"
            else "Session only, or use DOC_MEDIA_AI_API_KEY"
        )
        self.context_input = QTextEdit(self.settings.value("ai/context", "", str))
        self.context_input.setPlaceholderText(
            "可选，例如：示例项目素材；分类优先按项目/年份"
            if language == "zh"
            else "Optional, e.g. project and classification conventions"
        )
        self.context_input.setMinimumHeight(64)
        self.context_input.setMaximumHeight(96)
        for field in (
            self.base_url_input,
            self.model_input,
            self.api_key_input,
        ):
            field.setMinimumWidth(440)
        self.vision_checkbox = QCheckBox(
            "发送图片预览（仅在模型支持视觉时开启）"
            if language == "zh"
            else "Send image previews (enable only for vision models)"
        )
        self.vision_checkbox.setChecked(
            self.settings.value("ai/vision_enabled", False, bool)
        )
        self.vision_checkbox.setToolTip(
            "模型接口通常不会可靠声明视觉能力，因此不按模型名称自动猜测。"
            if language == "zh"
            else "Model APIs do not reliably declare vision support, so this is not guessed from the model name."
        )
        ai_form.addRow("Base URL", self.base_url_input)
        ai_form.addRow("模型" if language == "zh" else "Model", self.model_input)
        ai_form.addRow("API Key", self.api_key_input)
        ai_form.addRow(
            "业务分析上下文" if language == "zh" else "Business context",
            self.context_input,
        )
        ai_form.addRow("", self.vision_checkbox)
        self.capability_label = QLabel(
            "尚未验证模型能力" if language == "zh" else "Model capability not verified"
        )
        self.capability_label.setObjectName("dialogSubtitle")
        ai_form.addRow(
            "验证状态" if language == "zh" else "Validation", self.capability_label
        )
        root.addWidget(ai_box)

        library_box = QFrame()
        library_box.setObjectName("settingsSection")
        library_form = QFormLayout(library_box)
        library_form.setContentsMargins(12, 12, 12, 12)
        library_form.setFieldGrowthPolicy(
            QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow
        )
        self.video_library_input = QLineEdit(
            persistent_library_setting(
                self.settings,
                "video_manager/last_project",
            )
        )
        self.image_library_input = QLineEdit(
            persistent_library_setting(
                self.settings,
                "image_manager/last_project",
            )
        )
        self.video_library_input.setMinimumWidth(440)
        self.image_library_input.setMinimumWidth(440)
        video_library_row = QWidget()
        video_library_layout = QHBoxLayout(video_library_row)
        video_library_layout.setContentsMargins(0, 0, 0, 0)
        video_library_layout.setSpacing(8)
        video_library_layout.addWidget(self.video_library_input, 1)
        video_browse = QPushButton("选择…" if language == "zh" else "Browse…")
        video_browse.clicked.connect(
            lambda: self._choose_library(self.video_library_input)
        )
        video_library_layout.addWidget(video_browse)
        image_library_row = QWidget()
        image_library_layout = QHBoxLayout(image_library_row)
        image_library_layout.setContentsMargins(0, 0, 0, 0)
        image_library_layout.setSpacing(8)
        image_library_layout.addWidget(self.image_library_input, 1)
        image_browse = QPushButton("选择…" if language == "zh" else "Browse…")
        image_browse.clicked.connect(
            lambda: self._choose_library(self.image_library_input)
        )
        image_library_layout.addWidget(image_browse)
        library_form.addRow(
            "默认视频库" if language == "zh" else "Default video library",
            video_library_row,
        )
        library_form.addRow(
            "默认图片库" if language == "zh" else "Default image library",
            image_library_row,
        )
        root.addWidget(library_box)

        hint = QLabel(
            "API Key 不写入磁盘。AI 只生成整理、合并和主资源建议，应用前仍需人工确认。"
            if language == "zh"
            else (
                "The API key is not written to disk. AI suggestions require "
                "review before any library change."
            )
        )
        hint.setObjectName("dialogSubtitle")
        hint.setWordWrap(True)
        root.addWidget(hint)

        actions = QHBoxLayout()
        actions.setContentsMargins(8, 4, 8, 0)
        actions.setSpacing(12)
        self.test_button = QPushButton(
            "测试 AI 连接" if language == "zh" else "Test AI connection"
        )
        self.test_button.clicked.connect(self.test_connection)
        actions.addWidget(self.test_button)
        actions.addStretch(1)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.layout().setSpacing(12)
        save_button = buttons.button(QDialogButtonBox.StandardButton.Save)
        cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if save_button is not None:
            save_button.setText("保存" if language == "zh" else "Save")
        if cancel_button is not None:
            cancel_button.setText("取消" if language == "zh" else "Cancel")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        actions.addWidget(buttons)
        root.addLayout(actions)
        self.setStyleSheet(
            SHARED_DIALOG_QSS
            + """
QFrame#settingsSection {
    background: #0d1926;
    border: 1px solid #294057;
    border-radius: 8px;
}
QDialog QLabel { color: #cbd5e1; }
QDialog QCheckBox { color: #cbd5e1; }
QDialog QTextEdit {
    background: #0b1723;
    color: #e2e8f0;
    border: 1px solid #294057;
    border-radius: 7px;
    padding: 6px;
}
"""
        )

    def test_connection(self) -> None:
        from pptx_tools.ai_client import AIConfig

        if self.test_thread is not None:
            return
        config = AIConfig(
            self.base_url_input.text(),
            self.model_input.text(),
            self.api_key_input.text(),
            vision_enabled=self.vision_checkbox.isChecked(),
            timeout_seconds=20,
            context=self.context_input.toPlainText(),
        )
        self.test_button.setEnabled(False)
        self.capability_label.setText(
            "正在验证文本连接与图片输入能力…"
            if self.language == "zh"
            else "Testing text connection and image-input capability…"
        )
        self.test_thread = QThread(self)
        self.test_worker = AIConnectionWorker(config)
        self.test_worker.moveToThread(self.test_thread)
        self.test_thread.started.connect(self.test_worker.run)
        self.test_worker.finished.connect(self._connection_test_finished)
        self.test_worker.failed.connect(self._connection_test_failed)
        self.test_worker.finished.connect(self.test_thread.quit)
        self.test_worker.failed.connect(self.test_thread.quit)
        self.test_worker.finished.connect(self.test_worker.deleteLater)
        self.test_worker.failed.connect(self.test_worker.deleteLater)
        self.test_thread.finished.connect(self._connection_thread_finished)
        self.test_thread.finished.connect(self.test_thread.deleteLater)
        self.test_thread.start()

    def _connection_test_finished(self, result: object) -> None:
        if not isinstance(result, dict):
            self._connection_test_failed(
                "AI 验证结果无效。"
                if self.language == "zh"
                else "Invalid AI validation result."
            )
            return
        vision = result.get("vision")
        if vision is not None:
            self.vision_checkbox.setChecked(bool(vision))
        if self.language == "zh":
            vision_text = (
                "图片输入可用，已开启"
                if vision is True
                else (
                    "图片输入不可用，已关闭"
                    if vision is False
                    else "图片输入能力暂时无法确认，保留当前选择"
                )
            )
            capability_text = (
                f"文本连接正常；{vision_text}；JSON 兼容模式将在正式请求时自动降级。"
            )
            success_text = f"连接成功：{result.get('message', '')}\n{vision_text}"
        else:
            vision_text = (
                "Image input is available and enabled"
                if vision is True
                else (
                    "Image input is unavailable and disabled"
                    if vision is False
                    else "Image-input capability is unknown; the current choice was kept"
                )
            )
            capability_text = (
                f"Text connection succeeded; {vision_text}. JSON compatibility "
                "will fall back automatically when needed."
            )
            success_text = (
                f"Connection succeeded: {result.get('message', '')}\n{vision_text}"
            )
        self.capability_label.setText(capability_text)
        QMessageBox.information(
            self,
            "AI",
            success_text,
        )

    def _connection_test_failed(self, message: str) -> None:
        prefix = "验证失败" if self.language == "zh" else "Validation failed"
        self.capability_label.setText(f"{prefix}: {message}")
        QMessageBox.warning(self, "AI", message)

    def _connection_thread_finished(self) -> None:
        self.test_worker = None
        self.test_thread = None
        self.test_button.setEnabled(True)

    def _choose_library(self, target: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "选择资源库目录" if self.language == "zh" else "Choose library folder",
            target.text().strip() or str(Path.home()),
        )
        if selected:
            target.setText(selected)

    def closeEvent(self, event) -> None:  # noqa: N802, ANN001
        if self.test_thread is not None:
            self.capability_label.setText(
                "请等待当前 AI 验证完成后再关闭。"
                if self.language == "zh"
                else "Wait for the current AI validation before closing."
            )
            event.ignore()
            return
        super().closeEvent(event)

    def save_values(self) -> dict[str, object]:
        from pptx_tools.ai_client import normalize_model_name

        values: dict[str, object] = {
            "base_url": self.base_url_input.text().strip(),
            "model": normalize_model_name(self.model_input.text()),
            "api_key": self.api_key_input.text().strip(),
            "vision_enabled": self.vision_checkbox.isChecked(),
            "context": self.context_input.toPlainText().strip(),
            "video_library": self.video_library_input.text().strip(),
            "image_library": self.image_library_input.text().strip(),
        }
        self.settings.setValue("ai/base_url", values["base_url"])
        self.settings.setValue("ai/model", values["model"])
        self.settings.setValue("ai/vision_enabled", values["vision_enabled"])
        self.settings.setValue("ai/context", values["context"])
        self.settings.setValue("video_manager/last_project", values["video_library"])
        self.settings.setValue("image_manager/last_project", values["image_library"])
        return values


class ToolSwitch(QWidget):
    currentChanged = Signal(int)
    SEGMENT_WIDTH = 100
    HEIGHT = 42
    TRACK_INSET = 0.0

    def __init__(
        self, labels: list[str], language: str, parent: QWidget | None = None
    ) -> None:
        super().__init__(parent)
        self.labels = labels[:]
        self.language = language
        self._current_index = 0
        self._hover_index = -1
        self._knob_progress = 0.0
        self.animation = QPropertyAnimation(self, b"knobProgress", self)
        self.animation.setDuration(180)
        self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAccessibleName("功能切换" if language == "zh" else "Tool switch")
        self._update_accessible_description()
        self.setMouseTracking(True)
        self.setFixedSize(self.sizeHint())

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(
            max(1, len(self.labels)) * self.SEGMENT_WIDTH + int(self.TRACK_INSET * 2),
            self.HEIGHT,
        )

    def current_index(self) -> int:
        return self._current_index

    def set_current_index(self, index: int) -> None:
        if index < 0 or index >= len(self.labels) or index == self._current_index:
            return
        self._current_index = index
        self.animation.stop()
        self.animation.setStartValue(self._knob_progress)
        self.animation.setEndValue(float(index))
        self.animation.start()
        self._update_accessible_description()
        self.currentChanged.emit(index)

    def force_current_index(self, index: int) -> None:
        if index < 0 or index >= len(self.labels):
            return
        self.animation.stop()
        self._current_index = index
        self._knob_progress = float(index)
        self._update_accessible_description()
        self.update()

    def _update_accessible_description(self) -> None:
        if not self.labels:
            self.setAccessibleDescription("")
            return
        current = self.labels[self._current_index]
        self.setAccessibleDescription(
            f"当前：{current}。使用方向键、Home、End 或数字键切换。"
            if self.language == "zh"
            else (
                f"Current: {current}. Use arrow keys, Home, End, or number keys "
                "to switch tools."
            )
        )

    def knob_progress(self) -> float:
        return self._knob_progress

    def set_knob_progress(self, value: float) -> None:
        self._knob_progress = value
        self.update()

    knobProgress = Property(float, knob_progress, set_knob_progress)

    def segment_index_at(self, x: float) -> int:
        if self.width() <= 0 or not self.labels:
            return -1
        content_width = max(1.0, self.width() - self.TRACK_INSET * 2)
        segment_width = content_width / len(self.labels)
        index = int((x - self.TRACK_INSET) / segment_width)
        return max(0, min(len(self.labels) - 1, index))

    def mousePressEvent(self, event) -> None:  # noqa: N802, ANN001
        if (
            event.button() == Qt.MouseButton.LeftButton
            and self.width() > 0
            and self.labels
        ):
            self.set_current_index(self.segment_index_at(event.position().x()))
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # noqa: N802, ANN001
        next_hover = self.segment_index_at(event.position().x())
        if next_hover != self._hover_index:
            self._hover_index = next_hover
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802, ANN001
        self._hover_index = -1
        self.update()
        super().leaveEvent(event)

    def keyPressEvent(self, event) -> None:  # noqa: N802, ANN001
        if event.key() in (Qt.Key.Key_Right, Qt.Key.Key_Down):
            self.set_current_index(min(len(self.labels) - 1, self._current_index + 1))
            return
        if event.key() in (Qt.Key.Key_Left, Qt.Key.Key_Up):
            self.set_current_index(max(0, self._current_index - 1))
            return
        if event.key() == Qt.Key.Key_Home:
            self.set_current_index(0)
            return
        if event.key() == Qt.Key.Key_End:
            self.set_current_index(len(self.labels) - 1)
            return
        number_index = event.key() - Qt.Key.Key_1
        if 0 <= number_index < len(self.labels):
            self.set_current_index(number_index)
            return
        super().keyPressEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802, ANN001
        super().paintEvent(event)
        if not self.labels:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect())
        segment_width = rect.width() / len(self.labels)
        selected_rect = QRectF(
            rect.left() + self._knob_progress * segment_width + 6,
            rect.top() + 4,
            segment_width - 12,
            rect.height() - 10,
        )
        painter.setPen(QColor("#34506b"))
        painter.setBrush(QColor("#102032"))
        painter.drawRoundedRect(selected_rect, 9, 9)

        font = QFont(self.font())
        font.setPixelSize(13)
        font.setWeight(QFont.Weight.DemiBold)
        painter.setFont(font)
        for index, label in enumerate(self.labels):
            text_rect = QRectF(
                rect.left() + index * segment_width,
                rect.top(),
                segment_width,
                rect.height() - 4,
            )
            distance = abs(index - self._knob_progress)
            if index == self._hover_index and distance >= 0.42:
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(QColor("#122235"))
                painter.drawRoundedRect(text_rect.adjusted(8, 5, -8, -5), 8, 8)
            if distance < 0.42:
                text_color = QColor("#ff720d")
            elif index == self._hover_index:
                text_color = QColor("#cbd5e1")
            else:
                text_color = QColor("#94a3b8")
            painter.setPen(text_color)
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, label)

        underline_width = min(84.0, segment_width - 32.0)
        underline_x = (
            rect.left()
            + self._knob_progress * segment_width
            + (segment_width - underline_width) / 2
        )
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#ff720d"))
        underline_height = 3.0
        painter.drawRoundedRect(
            QRectF(
                underline_x,
                rect.bottom() - underline_height,
                underline_width,
                underline_height,
            ),
            underline_height / 2,
            underline_height / 2,
        )


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        app = QApplication.instance()
        if app is not None:
            configure_ui_font(app)
        super().__init__()
        self.language = detect_language()
        self.text = STRINGS[self.language]
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        persistent_library_setting(
            settings,
            "video_manager/last_project",
        )
        persistent_library_setting(
            settings,
            "image_manager/last_project",
        )
        self.session_ai_api_key = os.environ.get("DOC_MEDIA_AI_API_KEY", "")
        self._publish_ai_settings()
        self.setObjectName("toolboxWindow")
        self.setWindowTitle(self.text["app_name"])
        self.setMinimumSize(880, 560)
        self.resize(960, 620)
        self.setUnifiedTitleAndToolBarOnMac(True)

        icon_path = resource_path("assets", "app_icon.png")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.tabs = QTabWidget()
        self.tabs.setObjectName("toolTabs")
        self.tabs.tabBar().setObjectName("toolTabBar")
        self.tabs.setDocumentMode(False)
        self.tabs.setMovable(False)
        self.tabs.setUsesScrollButtons(False)
        self.tabs.setElideMode(Qt.TextElideMode.ElideRight)
        self.tabs.tabBar().hide()
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.embedded_tools: list[EmbeddedTool] = []
        self.tab_shortcuts: list[QShortcut] = []
        self.switcher: ToolSwitch | None = None
        self.switch_host: QWidget | None = None
        self.page_animation: QPropertyAnimation | None = None
        self.current_tab_index = -1
        self.shutting_down = False
        app = QApplication.instance()
        if app is not None:
            app.aboutToQuit.connect(self.shutdown_embedded_tools)

        central = QWidget()
        central.setObjectName("shellCentral")
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 6, 10, 6)
        root.setSpacing(6)

        self.header_card = QFrame()
        self.header_card.setObjectName("shellHeaderCard")
        self.header_card.setFixedHeight(58)
        header_layout = QHBoxLayout(self.header_card)
        header_layout.setContentsMargins(12, 4, 12, 4)
        header_layout.setSpacing(10)

        title_stack_widget = QWidget(self.header_card)
        title_stack_widget.setMinimumWidth(280)
        title_stack = QVBoxLayout(title_stack_widget)
        title_stack.setContentsMargins(0, 0, 0, 0)
        title_stack.setSpacing(3)
        title_stack.addStretch(1)

        self.header_eyebrow = QLabel()
        self.header_eyebrow.setObjectName("shellEyebrow")
        self.header_eyebrow.hide()

        self.header_title = QLabel()
        self.header_title.setObjectName("shellTitle")
        self.header_title.setFixedHeight(26)
        title_stack.addWidget(self.header_title)

        self.header_subtitle = QLabel()
        self.header_subtitle.setObjectName("shellSubtitle")
        self.header_subtitle.setFixedHeight(16)
        title_stack.addWidget(self.header_subtitle)
        title_stack.addStretch(1)

        self.header_icon = QLabel(self.header_card)
        self.header_icon.setObjectName("shellAppIcon")
        self.header_icon.setFixedSize(34, 34)
        self.header_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_icon_path = resource_path("assets", "app_icon_v2.png")
        if header_icon_path.exists():
            self.header_icon.setPixmap(
                QPixmap(str(header_icon_path)).scaled(
                    32,
                    32,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )

        self.switch_host = QWidget(self.header_card)
        self.switch_host.setFixedSize(
            ToolSwitch.SEGMENT_WIDTH * 4 + int(ToolSwitch.TRACK_INSET * 2),
            ToolSwitch.HEIGHT,
        )
        switch_host_layout = QHBoxLayout(self.switch_host)
        switch_host_layout.setContentsMargins(0, 0, 0, 0)
        switch_host_layout.setSpacing(0)
        switch_host_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        right_slot = QWidget(self.header_card)
        right_slot.setFixedWidth(88)
        right_slot_layout = QHBoxLayout(right_slot)
        right_slot_layout.setContentsMargins(0, 0, 0, 0)
        right_slot_layout.setSpacing(6)
        right_slot_layout.addStretch(1)

        self.settings_button = QPushButton()
        self.settings_button.setObjectName("shellSettingsButton")
        self.settings_button.setAccessibleName(
            "设置" if self.language == "zh" else "Settings"
        )
        self.settings_button.setToolTip(
            "AI、视频库和图片库设置"
            if self.language == "zh"
            else "AI, video, and image library settings"
        )
        self.settings_button.setIcon(
            QIcon(str(resource_path("assets", "icons", "settings.svg")))
        )
        self.settings_button.setIconSize(QSize(24, 24))
        self.settings_button.setFixedSize(40, 40)
        self.settings_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.settings_button.clicked.connect(self.show_settings)
        right_slot_layout.addWidget(
            self.settings_button,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )

        self.help_button = QPushButton()
        self.help_button.setObjectName("shellHelpButton")
        self.help_button.setAccessibleName(self.text["help_button"])
        self.help_button.setToolTip(self.text["help_button"])
        self.help_button.setIcon(
            QIcon(str(resource_path("assets", "icons", "help.svg")))
        )
        self.help_button.setIconSize(QSize(24, 24))
        self.help_button.setFixedSize(40, 40)
        self.help_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.help_button.clicked.connect(self.show_help)
        right_slot_layout.addWidget(
            self.help_button,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )

        header_layout.addWidget(self.header_icon, 0, Qt.AlignmentFlag.AlignVCenter)
        header_layout.addWidget(title_stack_widget, 1)
        header_layout.addWidget(self.switch_host, 0, Qt.AlignmentFlag.AlignRight)
        header_layout.addWidget(right_slot, 0)
        root.addWidget(self.header_card)
        root.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

        self.add_tool_tab(self.text["watermark_tab"], self.create_watermark_window)
        self.add_tool_tab(self.text["video_tab"], self.create_video_window)
        self.add_tool_tab(self.text["manager_tab"], self.create_manager_window)
        self.add_tool_tab(
            self.text["image_manager_tab"], self.create_image_manager_window
        )
        self.install_tool_switch()
        self.install_tab_shortcuts()
        self.apply_shell_style()
        install_control_help(self)
        self.on_tab_changed(self.tabs.currentIndex())

    def create_embedded_tool(
        self, title: str, factory: Callable[[], QMainWindow]
    ) -> EmbeddedTool:
        language_variables = (
            "PPTX_OUTPUT_WATERMARK_LANG",
            "PPTX_VIDEO_COMPACTOR_LANG",
        )
        previous_languages = {
            variable: os.environ.get(variable) for variable in language_variables
        }
        try:
            for variable in language_variables:
                os.environ[variable] = self.language
            window = factory()
        finally:
            for variable, previous in previous_languages.items():
                if previous is None:
                    os.environ.pop(variable, None)
                else:
                    os.environ[variable] = previous
        widget = window.takeCentralWidget()
        if widget is None:
            raise RuntimeError("tool window did not provide a central widget")
        widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.prepare_embedded_widget(widget)
        stylesheet = window.styleSheet()
        if stylesheet:
            widget.setStyleSheet(stylesheet)
        for control_type in (
            QAbstractButton,
            QAbstractSpinBox,
            QComboBox,
            QLineEdit,
        ):
            for control in widget.findChildren(control_type):
                control.ensurePolished()
                control.setMinimumHeight(
                    max(control.minimumHeight(), control.sizeHint().height())
                )
        QTimer.singleShot(
            0, lambda widget=widget: self._normalize_embedded_control_heights(widget)
        )
        return EmbeddedTool(title=title, factory=factory, window=window, widget=widget)

    @staticmethod
    def _normalize_embedded_control_heights(widget: QWidget) -> None:
        for control_type in (
            QAbstractButton,
            QAbstractSpinBox,
            QComboBox,
            QLineEdit,
        ):
            for control in widget.findChildren(control_type):
                control.setMinimumHeight(
                    max(control.minimumHeight(), control.sizeHint().height())
                )

    @staticmethod
    def prepare_embedded_widget(widget: QWidget) -> None:
        header = widget.findChild(QFrame, "headerCard")
        if header is not None:
            header.hide()
            header.setMaximumHeight(0)
        layout = widget.layout()
        if layout is not None:
            # Keep a small safe inset so the embedded page's top border is not
            # hidden by the tab pane edge on either platform.
            layout.setContentsMargins(0, 6, 0, 0)

    def add_tool_tab(self, title: str, factory: Callable[[], QMainWindow]) -> None:
        try:
            tool = self.create_embedded_tool(title, factory)
            self.embedded_tools.append(tool)
            index = self.tabs.addTab(tool.widget, title)
            self.tabs.setTabToolTip(index, title)
        except Exception as exc:
            widget = self.error_widget(title, exc)
            self.embedded_tools.append(
                EmbeddedTool(
                    title=title, factory=factory, window=None, widget=widget, error=exc
                )
            )
            index = self.tabs.addTab(widget, title)
            self.tabs.setTabToolTip(index, title)

    def install_tool_switch(self) -> None:
        if self.switch_host is None:
            return
        self.switcher = ToolSwitch(
            [self.tabs.tabText(index) for index in range(self.tabs.count())],
            self.language,
            self.switch_host,
        )
        self.switcher.currentChanged.connect(self.tabs.setCurrentIndex)
        layout = self.switch_host.layout()
        if layout is not None:
            layout.addWidget(self.switcher, 0, Qt.AlignmentFlag.AlignCenter)
        self.switcher.force_current_index(self.tabs.currentIndex())

    def install_tab_shortcuts(self) -> None:
        for index in range(self.tabs.count()):
            shortcut = QShortcut(QKeySequence(f"Ctrl+{index + 1}"), self)
            shortcut.activated.connect(
                lambda checked=False, tab_index=index: self.tabs.setCurrentIndex(
                    tab_index
                )
            )
            self.tab_shortcuts.append(shortcut)
            self.tabs.setTabToolTip(
                index, f"{self.tabs.tabText(index)}  Ctrl+{index + 1}"
            )

    def on_tab_changed(self, index: int) -> None:
        title = self.tabs.tabText(index) if index >= 0 else ""
        previous_index = self.current_tab_index
        selected_files: list[Path] = []
        if (
            previous_index >= 0
            and previous_index != index
            and previous_index < len(self.embedded_tools)
        ):
            selected_files = self.selected_files_for_tool(
                self.embedded_tools[previous_index]
            )
        self.current_tab_index = index
        if previous_index >= 0 and previous_index != index:
            self.restore_selected_files_to_tab(index, selected_files)
        if 0 <= index < len(self.embedded_tools):
            window = self.embedded_tools[index].window
            if window is not None and hasattr(window, "on_activated"):
                try:
                    window.on_activated()  # type: ignore[attr-defined]
                except Exception:
                    pass
        app_name = self.text["app_name"]
        self.setWindowTitle(f"{title} - {app_name}" if title else app_name)
        self.update_shell_header(index)
        self.update_switch_buttons(index)
        self.animate_current_tab(index)

    def update_shell_header(self, index: int) -> None:
        key = (
            ("watermark", "video", "manager", "image_manager")[index]
            if 0 <= index < 4
            else "watermark"
        )
        self.header_eyebrow.setText(self.text[f"{key}_eyebrow"])
        self.header_title.setText(self.text[f"{key}_title"])
        self.header_subtitle.setText(self.text[f"{key}_subtitle"])

    def update_switch_buttons(self, index: int) -> None:
        if self.switcher is None:
            return
        if self.switcher.current_index() != index:
            self.switcher.force_current_index(index)

    def animate_current_tab(self, index: int) -> None:
        if index < 0:
            return
        widget = self.tabs.widget(index)
        if widget is None:
            return
        if self.page_animation is not None:
            self.page_animation.stop()
        effect = QGraphicsOpacityEffect(widget)
        widget.setGraphicsEffect(effect)
        animation = QPropertyAnimation(effect, b"opacity", self)
        animation.setDuration(220)
        animation.setStartValue(0.25)
        animation.setEndValue(1.0)
        animation.setEasingCurve(QEasingCurve.Type.OutCubic)
        animation.finished.connect(lambda target=widget: target.setGraphicsEffect(None))
        self.page_animation = animation
        animation.start()

    def _publish_ai_settings(self, values: dict[str, object] | None = None) -> None:
        app = QApplication.instance()
        if app is None:
            return
        settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")
        current = values or {
            "base_url": settings.value("ai/base_url", "https://api.openai.com/v1", str),
            "model": settings.value("ai/model", "", str),
            "api_key": self.session_ai_api_key,
            "vision_enabled": settings.value("ai/vision_enabled", False, bool),
            "context": settings.value("ai/context", "", str),
        }
        app.setProperty("doc_media_ai_base_url", current.get("base_url", ""))
        app.setProperty("doc_media_ai_model", current.get("model", ""))
        app.setProperty("doc_media_ai_api_key", current.get("api_key", ""))
        app.setProperty(
            "doc_media_ai_vision_enabled", bool(current.get("vision_enabled", False))
        )
        app.setProperty("doc_media_ai_context", current.get("context", ""))

    def show_settings(self) -> None:
        dialog = SettingsDialog(self.language, self.session_ai_api_key, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        values = dialog.save_values()
        self.session_ai_api_key = str(values["api_key"])
        self._publish_ai_settings(values)
        for tool in self.embedded_tools:
            if tool.window is not None and hasattr(tool.window, "on_settings_changed"):
                tool.window.on_settings_changed()  # type: ignore[attr-defined]

    def show_help(self) -> None:
        dialog = HelpDialog(self.language, self.tabs.currentIndex(), self)
        dialog.exec()

    def selected_files_for_tool(self, tool: EmbeddedTool) -> list[Path]:
        if tool.window is None:
            return []
        paths = getattr(tool.window, "input_paths", [])
        result: list[Path] = []
        for path in paths:
            try:
                resolved = Path(path).expanduser().resolve()
            except TypeError:
                continue
            result.append(resolved)
        return result

    def restore_selected_files(self, window: QMainWindow, paths: list[Path]) -> None:
        if not paths or not hasattr(window, "set_files"):
            return
        try:
            window.set_files(paths)  # type: ignore[attr-defined]
        except Exception:
            pass

    def restore_selected_files_to_tab(self, index: int, paths: list[Path]) -> None:
        if not paths or index < 0 or index >= len(self.embedded_tools):
            return
        tool = self.embedded_tools[index]
        if tool.window is None or self.tool_is_busy(tool.window):
            return
        self.restore_selected_files(tool.window, paths)

    @staticmethod
    def tool_is_busy(window: QMainWindow) -> bool:
        if bool(getattr(window, "is_running", False)):
            return True
        for attr_name in ("worker_thread", "preview_thread", "ai_thread"):
            thread = getattr(window, attr_name, None)
            if (
                thread is not None
                and hasattr(thread, "isRunning")
                and thread.isRunning()
            ):
                return True
        return False

    def shutdown_embedded_tools(self) -> None:
        if self.shutting_down:
            return
        self.shutting_down = True
        if self.page_animation is not None:
            self.page_animation.stop()
            self.page_animation = None
        for tool in self.embedded_tools:
            if tool.window is not None:
                self.shutdown_tool_window(tool.window)

    @staticmethod
    def shutdown_tool_window(window: QMainWindow) -> None:
        timer = getattr(window, "preview_timer", None)
        if timer is not None and hasattr(timer, "stop"):
            timer.stop()
        if bool(getattr(window, "is_running", False)) and hasattr(window, "stop_job"):
            try:
                window.stop_job()  # type: ignore[attr-defined]
            except Exception:
                pass
        audit_worker = getattr(window, "audit_worker", None)
        if audit_worker is not None and hasattr(audit_worker, "cancel"):
            try:
                audit_worker.cancel()
            except Exception:
                pass
        ai_worker = getattr(window, "ai_worker", None)
        if ai_worker is not None and hasattr(ai_worker, "cancel"):
            window.ai_ignore_result = True  # type: ignore[attr-defined]
            ai_worker.cancel()
        for attr_name in (
            "preview_thread",
            "worker_thread",
            "audit_thread",
            "ai_thread",
        ):
            thread = getattr(window, attr_name, None)
            if (
                thread is None
                or not hasattr(thread, "isRunning")
                or not thread.isRunning()
            ):
                continue
            terminate_active_processes(grace_seconds=0.1)
            thread.quit()
            while not thread.wait(250):
                terminate_active_processes(grace_seconds=0.1)

    def apply_shell_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow#toolboxWindow {
                background: #07101a;
            }
            QWidget#shellCentral {
                background: #07101a;
            }
            QFrame#shellHeaderCard {
                background: #07101a;
                border: 0;
                border-radius: 0;
            }
            QLabel#shellEyebrow {
                color: #ff720d;
                font-size: 11px;
                font-weight: 700;
                letter-spacing: 0.12em;
            }
            QLabel#shellTitle {
                color: #f8fafc;
                font-size: 18px;
                font-weight: 600;
            }
            QLabel#shellSubtitle {
                color: #8fa4ba;
                font-size: 12px;
            }
            QFrame#toolSwitcher {
                background: #080f19;
                border: 1px solid #314256;
                border-radius: 23px;
            }
            QPushButton#toolSwitchButton {
                background: transparent;
                color: #94a3b8;
                border: 0;
                border-radius: 17px;
                min-height: 42px;
                padding: 7px 30px;
                font-size: 11px;
                font-weight: 700;
            }
            QPushButton#toolSwitchButton:hover {
                background: #172234;
                color: #dbe4f0;
            }
            QPushButton#toolSwitchButton[selected="true"] {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #fb923c, stop:0.55 #f97316, stop:1 #ea580c);
                color: #ffffff;
                border: 1px solid #fed7aa;
            }
            QPushButton#shellHelpButton, QPushButton#shellSettingsButton {
                color: #cbd5e1;
                background: transparent;
                border: 0;
                border-radius: 8px;
                padding: 0;
                min-width: 40px;
                max-width: 40px;
                min-height: 40px;
                max-height: 40px;
            }
            QPushButton#shellHelpButton:hover, QPushButton#shellSettingsButton:hover {
                background: #122235;
                color: #f8fafc;
            }
            QTabWidget#toolTabs {
                background: #07101a;
            }
            QTabWidget#toolTabs::pane {
                background: #07101a;
                border: 0;
                top: -1px;
            }
            QTabBar#toolTabBar {
                background: #07101a;
            }
            QTabBar#toolTabBar::tab {
                width: 0;
                height: 0;
                margin: 0;
                padding: 0;
                border: 0;
            }
            QLabel#toolLoadError {
                color: #f8fafc;
                background: #121a24;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 16px;
            }
            """
        )

    @staticmethod
    def create_watermark_window() -> QMainWindow:
        from pptx_output_watermark.gui import MainWindow as WatermarkWindow

        return WatermarkWindow()

    @staticmethod
    def create_video_window() -> QMainWindow:
        from pptx_video_compactor_gui import MainWindow as VideoWindow

        return VideoWindow()

    @staticmethod
    def create_manager_window() -> QMainWindow:
        from pptx_tools.video_manager_gui import MainWindow as ManagerWindow

        return ManagerWindow()

    @staticmethod
    def create_image_manager_window() -> QMainWindow:
        from pptx_tools.image_manager_gui import MainWindow as ImageManagerWindow

        return ImageManagerWindow()

    def error_widget(self, title: str, exc: Exception) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        label = QLabel(self.text["load_failed"].format(title=title, error=exc))
        label.setObjectName("toolLoadError")
        label.setWordWrap(True)
        layout.addWidget(label)
        return widget

    def closeEvent(self, event) -> None:  # noqa: N802
        self.shutdown_embedded_tools()
        for tool in self.embedded_tools:
            if tool.window is not None:
                tool.window.close()
        super().closeEvent(event)


def main() -> int:
    configure_app_logging()
    app = QApplication([])
    configure_ui_font(app)
    language = detect_language()
    app.setApplicationName(APP_NAME)
    app.setApplicationDisplayName(STRINGS[language]["app_name"])
    icon_path = resource_path("assets", "app_icon.png")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
