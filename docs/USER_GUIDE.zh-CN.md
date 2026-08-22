# Doc Media Toolkit 完整用户指南

[项目首页](../README.md) | [English README](../README.md)

`Doc Media Toolkit`（中文界面名：`文档媒体工具箱`）是一个统一的文档媒体处理与资产管理应用，把水印导出、动态压缩、PPTX 视频资产管理和文档图片资产管理整合到同一个代码仓库、桌面应用和统一 CLI 里；开发安装后的 Python 命令仍保留为 `pptx-tools`：

- `文档及媒体水印导出`：批量处理 `.docx` / `.pdf` / `.pptx`，也支持独立图片和视频。DOCX/PDF 导出为 PDF；PPTX 可导出为 PDF/PPTX；独立图片按原格式导出，独立视频按原容器或 MP4 导出。支持可编辑/图片化模式、水印文字/图片、字体补齐、图片清晰度、视频回贴。
- `文档及媒体动态压缩`：核心场景是按质量底线优化 PPTX 容量；独立图片/视频和 DOCX/PDF/XLSX 文档是附带输入，可拖入或添加到同一队列并共用设置，不单设专业压缩接口。视频按显示面积、复用次数、时长和码率分配预算；图片结合显示面积与本地内容分类处理。DOCX/PDF/XLSX 仅压缩内嵌图片，必须显式填写目标容量（MB）且图片预设不能为"不压缩"；PDF 需要 Poppler，DOCX/XLSX 渲染校验需要 LibreOffice/Pages，缺失时拒绝发布未验证输出。
- `PPTX 视频资产库`：从 PPTX 归档唯一高清源并登记形状关联，按 SHA-256 精确去重，并对不同编码版做保守内容匹配；压缩版 PPTX 保留可播放的低体积 MP4，需要时从视频库高清替换。
- `文档图片资产库`：归档独立图片以及 PPTX、DOCX、数字版 PDF 实际引用的内嵌图片，按 SHA-256 精确去重；相似图、命名和分类只给出审核建议，当前不做图片回填。

> 项目状态：正式版本。PPTX 压缩、水印、视频资产库及主要文档兼容能力已经过确认；图片资产管理是附加能力，尚未完成与核心功能同等级的深度实测。支持双语的工作区首次启动跟随系统界面语言，可用 `PPTX_TOOLS_LANG=zh` 或 `PPTX_TOOLS_LANG=en` 覆盖；视频库和图片库的业务界面目前仍以中文为主，请勿宣传为“完整双语”。

## 名称与定位

- 对外产品名：`Doc Media Toolkit` / `文档媒体工具箱`。
- 公开仓库名使用 `doc-media-toolkit`；Python 包和 CLI 继续使用 `pptx-tools`，避免破坏现有脚本。
- 项目聚焦“文档中的媒体处理与资产管理”，不是通用视频或图片编辑器。

## 当前整合方式

这是独立项目，不反向依赖原来的项目目录。核心代码已经复制到本仓库内：

- `src/pptx_output_watermark/`：文档及媒体水印导出核心与原 GUI。
- `src/pptx_video_compactor.py`：文档及媒体动态压缩核心。
- `src/pptx_video_compactor_gui.py`：文档及媒体动态压缩 GUI。
- `src/pptx_tools/gui.py`：统一页签 GUI 壳。
- `src/pptx_tools/cli.py`：统一 CLI 子命令入口。
- `src/pptx_tools/video_manager.py` / `image_manager.py`：视频与图片资产库领域模型。
- `src/pptx_tools/media_manager_ui.py` / `ui_theme.py`：两个资产库共用的 worker、样式和 UI 基础设施；资产库 GUI 之间不得互相导入。

模块边界、数据不变量、依赖方向和扩展规则见 [`ARCHITECTURE.md`](ARCHITECTURE.md)。共用依赖、资源和打包逻辑统一如下：

- 共用一个 Python 环境和 PySide6。
- 共用 `assets/`，其中包含图标和水印字体。
- 共用 `config/`，其中包含媒体压缩的视频预设；图片预设在媒体压缩核心中按平台通用规则处理。
- 顶栏 `设置` 可配置 OpenAI 兼容接口。Base URL 只接受无内嵌凭据、查询或片段的 HTTP(S) 地址，响应体有大小上限；API Key 仅保存在当前进程。未启用视觉能力时 AI 只读取代码候选与规格，启用后才发送压缩预览。AI 仅建议命名、分类、疑似合并组和主资源，不自动归并或删除。
- 视频库 `更多操作` 可导入/导出便携哈希目录。导入时只有与本机单一视频族存在哈希锚点的记录才会合并；无锚点和冲突记录会跳过，媒体文件不会随目录传输。
- 共用 `ffmpeg` / `ffprobe` 打包逻辑。
- 标准轻量包不内置 Microsoft Office / WPS / LibreOffice / Keynote / Pages；
  构建时可显式生成包含完整 LibreOffice 的离线版。

## 开发环境

```bash
uv sync --locked --all-extras
```

这会创建 `.venv` 并按 `uv.lock` 安装开发、测试和打包依赖。视频相关功能还需要
`ffmpeg` / `ffprobe`；源码运行时会从显式配置或 PATH 中解析它们。
独立安装包固定使用 Python 3.12；脚本会拒绝误用系统 Python 3.14，
避免把未经验证的运行时打入发布包。

源码运行 GUI：

```bash
.venv/bin/pptx-tools-gui
```

源码运行 CLI：

```bash
.venv/bin/pptx-tools --help
.venv/bin/pptx-tools watermark --help
.venv/bin/pptx-tools compact --help
.venv/bin/pptx-tools videos --help
.venv/bin/pptx-tools images --help
```

运行测试：

```bash
QT_QPA_PLATFORM=offscreen PYTHONPATH=src:tests .venv/bin/python scripts/run_tests_isolated.py
```

## CLI

统一 CLI 使用子命令：

```bash
pptx-tools watermark [pptx-output-watermark 参数...]
pptx-tools compact [媒体压缩参数...]
```

示例：

```bash
pptx-tools watermark input.docx --output-format pdf --output-mode editable --watermark-text "企业专属，注意保密"
pptx-tools compact input.pptx --profile high --image-profile high
pptx-tools compact input.docx --target-size-mb 8 --image-profile high
pptx-tools compact input.pdf --target-size-mb 12 --image-profile balanced
pptx-tools compact input.xlsx --target-size-mb 6 --image-profile high
DataProject="$HOME/Documents/PPTX Videos"
pptx-tools videos create "$DataProject"
pptx-tools videos add "$DataProject" input.pptx
pptx-tools videos upgrade "$DataProject" input_compacted.pptx
pptx-tools videos import-video "$DataProject" candidate.mp4 --family-id FAMILY_ID --source-quality original
pptx-tools videos set-source "$DataProject" VARIANT_ID
pptx-tools videos doctor "$DataProject"
# 深度读取并核对每个媒体实体的 SHA-256：
pptx-tools videos doctor "$DataProject" --verify-hashes
# 仅把 WMV/AVI 等不兼容媒体迁移为 MP4，不改已有 MP4：
pptx-tools videos upgrade "$DataProject" input.pptx --incompatible-only
ImageProject="$HOME/Documents/Document Images"
pptx-tools images doctor "$ImageProject"
pptx-tools images doctor "$ImageProject" --verify-hashes
```

## PPTX 视频资产库

需要让本机外部 AI Agent 协助命名、归类、导入、回填或清理时，先让其阅读
[`VIDEO_LIBRARY_AGENT.md`](VIDEO_LIBRARY_AGENT.md)。该规范要求
`video-project.json` 只读，并通过现有 CLI、GUI 或 `VideoProject` 公共方法修改。
项目级完成标准、数据不变量和发布门禁见
[`QUALITY_GATES.md`](QUALITY_GATES.md)。
统一主窗口、弹窗、字号、状态颜色和交互层级的实现规范见
[`UI_DESIGN.md`](UI_DESIGN.md)。

视频资产库工作区只保存视频实体；PPTX 不会被复制进库，但会记录文件哈希、路径别名和幻灯片/形状关联：

1. 新建或打开一个视频库文件夹；程序不预设固定的视频库目录，可按客户或主题建立多个独立库并随时切换。当前库栏显示库名和完整路径，最近打开的库会自动恢复。
2. 从一个或多个 PPTX 点击 `归档 PPTX 高清视频`。该操作把视频实体入库，并在同一工作流中登记 PPTX 哈希、路径别名及幻灯片/形状关联；任一步失败会回滚本轮新增实体。可填写 `客户/年份/项目` 形式的相对分类目录。默认只把超过 1080p 的源按原比例转为高质量 1080p；也可选择 `保留原片`。字节完全相同的视频按 SHA-256 去重；不同编码或分辨率只有在时长、宽高比、五帧画面和音频频谱均唯一匹配时才归为同一视频族。PPTX 和文件名不作为身份。
3. 点击 `导入外部视频并匹配` 或把 MP4/MOV/MKV 等文件拖到视频列表可批量导入。完全相同的文件直接复用并追加来源路径；内容指纹唯一命中时加入已有视频族；没有可靠唯一结果时显示待匹配视频与候选高清源的 10%/50%/90% 三帧对比、分辨率、时长、画面和音频差异，由你决定归入现有视频族、新建视频族或跳过。人工确认前不会修改已有视频族或 PPTX 关联。更高分辨率匹配版只保存为候选，确认质量后再选择该版本并点击 `设为高清源`。大库可按名称、路径、分辨率或哈希片段即时筛选，也可点击统计数字直接筛出无 PPTX 关联、多版本或文件异常的视频族；“待核对”是这三类的去重并集，无关联和多版本通常不是错误，文件异常才表示库内文件不可用。单击视频族或版本会在列表右侧覆盖显示一张代表封面、规格、引用数、路径、哈希和状态，关闭后恢复完整列表；三帧对比只保留在需要判断同源关系的核对弹窗中。筛选 `待核对` 后选中一个视频族，点击 `核实版本` 可只检查该族及其重复候选；`归并视频` 用于人工确认两个视频族确属同一内容。点击任一表头按真实数值排序，再次点击切换升降序；筛选无结果时界面会明确提示如何恢复全部列表。悬停版本的文件位置可查看库内路径和已记录的外部来源。
4. 压缩页可选择 `视频不入库 / 1080p 高清入库 / 原片入库`：入库时先保存高清源并登记 PPTX，压缩后只登记低码率 MP4 的 SHA-256 别名，不在库中重复保存低清文件。PPTX 改名或移动后会按整包哈希复用原记录并保存路径别名。同一内容后来出现更高分辨率候选时会保留版本但不自动替换高清源。
5. 需要保留原文件名时，可勾选 `覆盖原 PPTX`。该选项仅允许 PPTX、已开启视频入库、只压缩视频、关闭自动评估的安全组合；程序先完成视频入库，再以临时文件原子替换源 PPTX，最后刷新整包哈希和幻灯片/形状锚点关联。选择 `原片入库` 才能保留入库当时的视频原始字节。
6. 点击 `高清回填 PPTX（另存）`。程序先生成整份 PPTX 的回填清单，逐项列出精确哈希匹配、唯一内容匹配、已是高清源和待人工确认的视频；选中项目后才按需生成当前视频与目标源的 10%/50%/90% 三帧封面对比，也可播放完整视频。每项都可选择 `保持当前`、`仅本次回填` 或 `确认同一视频并记住`，并可搜索全部视频族纠正自动建议。只有最后一种会在输出 ZIP 校验成功后登记当前媒体哈希；一次性替换不会污染视频族身份。取消清单不会修改视频库或 PPTX，确认后始终另存，不覆盖输入文件。
7. MP4/M4V 容器、H.264/AAC 且不超过 1080p 的高清源直接按原字节回填，并在 PPTX 包内使用 `.mp4` 媒体部件。其他格式或更大画面仅在回填时临时生成 PowerPoint 兼容 MP4；不放大小视频，并保持宽高比和时长。`保留原片` 保证库内文件字节不变，但不保证不兼容源在 PPTX 内仍能原字节播放；默认 1080p 母版是高质量有损转换，不能恢复已丢失的 2K/4K 细节。

   回填确认清单中可选择回填质量档位，默认 `最佳` 与原有行为一致。各档位均为上限语义：源已满足档位规格即按原字节嵌入，否则按档位转码；码率探测缺失时按超限转码（安全方向）。

   | 档位 | 分辨率上限 | CRF | 码率上限 | 音频 | 默认输出后缀 |
   | ---- | ---------- | --- | -------- | ---- | ------------ |
   | 最佳 | 1920×1080 | 18 | 不限 | AAC 256k | `_high_quality` |
   | 高质量 | 1920×1080 | 20 | 12 Mbps | AAC 256k | `_hq1080p` |
   | 均衡 | 1280×720（竖版 720×1280） | 23 | 5 Mbps | AAC 128k | `_balanced720p` |

8. 点击 `库体检` 可只读检查媒体实体、哈希归属、PPTX 来源与形状关联、历史输出及待清理索引；需要时执行完整哈希复核。CLI 可用 `videos doctor` 生成同一 JSON 结果。失效的历史输出记录只是诊断信息，可单独清理，不会删除视频或 PPTX 关联。

视频库中的 `video-project.json` 是权威索引，记录视频族、高清源、实体 SHA-256、压缩别名、内容指纹和 PPTX 形状锚点；不保存 PPTX 或低清视频的额外实体副本。每个库独立包含 `media/`、`_cleanup/`、`reports/` 和清单备份，写入时拒绝旧窗口覆盖更新后的库。

视频文件名不参与身份判断。在视频库中选择视频族执行 `重命名` 会同步修改显示名和高清源文件名；选择视频族或具体版本均可移动文件。`添加版本` 同样要求内容指纹与所选视频族一致，避免手工误归类。若在 Finder 中改名或移动到其他下级目录，可用 `查找丢失视频` 按 SHA-256 递归重新关联。把完整视频库复制到另一台机器后，文件系统可能改变时间戳；这类文件显示为“待校验”而不是“已修改”，核验哈希一致后会刷新时间戳状态，不会因路径、文件名或时间戳不同直接改写视频身份。应用偏好由 Qt `QSettings` 保存在当前系统用户的应用设置目录，最近的视频库、分类目录、外部视频目录和列表排序会自动记住；“操作记录”只显示本次运行，四个工作区的长期滚动日志统一写入系统应用数据目录的 `Doc Media Toolkit/logs/`，可从界面的 `日志目录` 打开。

归并前会显示将迁移的视频版本、已知哈希别名、PPTX 数量和媒体引用次数。确认后，源视频族的全部 PPTX 关联及压缩版哈希都会迁移到目标族，并由清单一致性校验兜底；不会把引用悬空。候选名称中的历史 `.wmv`/`.avi` 扩展名只作为旧显示名处理，匹配窗口会隐藏该扩展名；实际文件格式以版本路径和编码信息为准。音频一致性使用有无音轨、频谱指纹及必要时的解码相关性判断；显示 `不同/未知` 时不能把它当作同一视频的自动证据，必须结合三帧和完整播放人工核对。

列表可按名称、路径、分辨率或哈希筛选；选中后可在右侧详情抽屉内使用 QtMultimedia 播放，双击视频族时播放当前高清源，不跳出应用窗口。

压缩页开启视频入库后会直接显示当前目标视频库，点击目标可切换；压缩前的视频源和压缩后媒体哈希都写入该库。关闭入库时不会建立可恢复关联。高清优化始终要求指定新文件或输出目录，不覆盖输入 PPTX。

高清优化不删除并重建视频形状，而是保留原海报、形状 ID、尺寸、裁剪、关系和播放时间线，只替换对应媒体包内容；WMV/AVI 迁移时媒体部件扩展名及关系目标会同步改为 `.mp4`。默认输出使用新文件名，不覆盖输入 PPTX。视频库采用“每个视频族一个当前最佳高清源”的模型，不保存每个 PPTX 的历史快照：旧 PPTX 可以恢复为当前最佳且可播放的高清版本，但如果同族后来升级了更高清源，恢复结果会使用新高清源，而不是强求回到该 PPTX 最初嵌入视频的逐字节版本。

### 整理视频库

`整理视频库` 扫描库内两类重复：同一视频族内的冗余版本（原片 vs 各压缩版），以及同一内容产生的跨族重复（完全相同的按 SHA-256，不同压缩版按时长/画面/音轨指纹聚类）。每组并排展示分辨率、码率、编码、大小、音轨、与权威源的 SSIM 和匹配置信度，并给出推荐保留项（保留权威源，或"体积更小但质量接近"——该推荐要求时长与音轨一致、SSIM 达到阈值且分辨率满足需求，再比较体积，不只看码率）。最终保留哪个由你勾选确认。

列表中的 `隔离异常` 只对非高清源、非当前版本且没有 PPTX 引用的不可读文件启用。操作仅把文件移入 `_cleanup/`，可从 `待清理` 恢复；不会直接删除，也不会改变健康版本和既有 PPTX 关联。

质量差距很小时可选择生成统一的 1080p 高质量版本：始终从已登记的权威源重新编码，不会对已有压缩版反复转码。整理对话框中的处理方式是三选一；“人工确认：连锁定版本也移入待清理”是独立附加选项，只对族内生效，并会在切换到其他处理方式时自动解除。清理动作先把移动意图写入 `index.json`，再把落选文件移到库内 `_cleanup/`；异常中断后，重开视频库会核验 SHA-256 并恢复为可还原状态。新索引使用库内相对路径，库整体移动后仍有效；索引损坏、内容变化或路径越界时拒绝还原/清空。确认 PPTX 哈希别名与视频族迁移完成后，才可在 `待清理` 对话框中永久清空。不同音轨、被裁剪、来源不明或置信度不足时整组默认跳过，绝不自动归并；族内如确实要处理锁定版本，可勾选人工强制隔离、保留已核实版本并完成二次确认，跨族归并仍禁止强制；跨族归并后的其余版本会在下一次扫描中继续做族内整理，避免同一轮重复处理。

图片库采用相同的意图日志与待清理策略：移除、重复图片合并和未引用文件清理只会先登记、再移入库内 `_cleanup/`，中断后可自动识别并从 `更多操作 → 查看待清理文件` 还原；只有再次确认“永久清空”才会删除文件。原始 PPTX、DOCX、PDF 和独立图片始终不受影响。

## 打包

macOS GUI，本机可运行包：

```bash
python scripts/build_standalone.py \
  --gui \
  --clean \
  --target-platform macos \
  --bundle-ffmpeg \
  --require-ffmpeg-bundle \
  --dmg \
  --name "Doc Media Toolkit"
```

如果需要指定 DMG 输出位置，可再加：

```bash
--dmg-output "/absolute/path/Doc Media Toolkit-macOS-arm64.dmg"
```

私下或内部环境需要无需安装办公软件的完整离线版时，使用 onedir 构建并显式加入
完整 LibreOffice 运行时：

```bash
python scripts/build_standalone.py \
  --gui \
  --target-platform macos \
  --bundle-ffmpeg \
  --require-ffmpeg-bundle \
  --bundle-libreoffice \
  --require-libreoffice-bundle \
  --dmg \
  --name "Doc Media Toolkit"
```

可用 `--libreoffice-root` 或 `PPTX_TOOLS_LIBREOFFICE_ROOT` 指定安装根目录。
构建脚本只接受同时包含转换程序和许可证的完整运行时。LibreOffice 版本不支持
`--onefile`，避免每次启动解压数百 MB；轻量版行为保持不变。该离线版公开前还要
单独核验 LibreOffice MPL 及其组件许可证要求的匹配源码获取方式，标准正式包不内置。

Windows GUI one-file：

```powershell
python scripts\build_standalone.py `
  --windows-onefile `
  --bundle-ffmpeg `
  --require-ffmpeg-bundle `
  --name "Doc Media Toolkit"
```

默认会使用 `assets/app_icon.ico` 作为 Windows 可执行文件图标；未显式传 `--name` 时，产物名默认也是 `Doc Media Toolkit`。

CLI one-file：

```bash
python scripts/build_standalone.py \
  --cli \
  --onefile \
  --clean \
  --bundle-ffmpeg \
  --require-ffmpeg-bundle \
  --name "Doc Media Toolkit"
```

说明：

- `--windows-onefile` 只适用于 Windows GUI。
- Windows CLI 仍使用上面的 `--cli --onefile` 命令，但必须在 Windows 主机上执行，不能跨平台编译。
- macOS `--dmg` 只适用于 GUI onedir 构建；生成的 DMG 会包含 `Doc Media Toolkit.app` 和 `Applications` 快捷方式，不会打入本机其他应用。

正式公开包必须先用固定源码生成 FFmpeg 8.1.2 运行时，再通过环境变量交给打包脚本：

```bash
scripts/build_ffmpeg_runtime.sh release-assets/ffmpeg-runtime
export PPTX_TOOLS_FFMPEG="$PWD/release-assets/ffmpeg-runtime/bin/ffmpeg"
export PPTX_TOOLS_FFPROBE="$PWD/release-assets/ffmpeg-runtime/bin/ffprobe"
export PPTX_TOOLS_FFMPEG_LICENSE_DIR="$PWD/release-assets/ffmpeg-runtime/licenses"
```

Windows 在 MSYS2 MINGW64 中运行同一脚本，生成 `.exe`。脚本以 SHA-256 固定
FFmpeg 8.1.2 与 x264，保留 libx264 以及 macOS VideoToolbox / Windows Media
Foundation，并生成同一 Release 必须携带的对应源码包。Homebrew/Gyan 预编译
FFmpeg 只能用于本地测试，不能进入正式公开包。

构建脚本默认拒绝打包单个超过 `260MB` 的 `ffmpeg/ffprobe` 二进制，避免误打进异常巨大的 full static build。私下测试的 Windows one-file 总体积超过 `200MB` 可以接受，但正式公开包必须使用 onedir portable ZIP；如确实需要取消单二进制上限，可加：

```bash
--max-bundled-binary-mb 0
```

构建脚本会把项目 MIT、第三方索引、Python/Qt/资源许可及各 Python 分发包的许可元数据放入安装包 `licenses/`；任一必需依赖缺少可用许可文本时构建失败。`--bundle-ffmpeg` 还会打入对应分发版的 `LICENSE/COPYING/NOTICE`；许可目录无法自动定位时，可用 `PPTX_TOOLS_FFMPEG_LICENSE_DIR` 显式指定。FFmpeg 是否属于 GPL 取决于真实构建参数，正式二进制发布还必须按 [`LICENSING.md`](LICENSING.md) 完成对应源码和产物级依赖审计。

统一项目优先使用 `PPTX_TOOLS_*` 环境变量；旧项目变量仍兼容：

```text
PPTX_TOOLS_FFMPEG / PPTX_TOOLS_FFPROBE
PPTX_TOOLS_SOFFICE
PPTX_TOOLS_LIBREOFFICE_ROOT
PPTX_TOOLS_WPS / PPTX_TOOLS_WPP
```

## 运行时依赖

为什么默认安装这些 Python 依赖、哪些只是 Windows/构建/传递依赖，见 [`DEPENDENCIES.md`](DEPENDENCIES.md)。默认安装面向完整桌面产品，不为缩短列表而让已声明功能在运行时临时缺包。

- 水印导出：`pypdfium2` 已随 Python 依赖打包，不再依赖 `pdftoppm`。
- PDF 导出：
  - `PPTX -> PDF`：Windows 优先使用 PowerPoint/WPS COM，失败后使用 LibreOffice；macOS 优先使用 LibreOffice，缺失或短时不可用时回退 Keynote。
  - `DOCX -> PDF`：Windows 优先使用 Word/WPS COM，失败后使用 LibreOffice；macOS 优先使用 LibreOffice，缺失或短时不可用时回退 Pages。
  - `PDF` 输入：不走外部 PDF 导出引擎，直接做可编辑 PDF 加水印或图片化 PDF 重建。
- Windows WPS：若 WPS COM 隐藏打开失败，会再尝试可见窗口打开；如 WPS 未注册 COM，可先手动打开一次 WPS/修复安装，或设置 `PPTX_TOOLS_WPP` 指向 `wpp.exe`。
- 媒体压缩中的视频处理：需要 `ffmpeg` / `ffprobe`，打包时建议内置；只压图片时不需要。
- 媒体压缩中的画质评估同样依赖 `ffmpeg`，因为通过 `ffmpeg + SSIM` 对处理后的媒体逐项对比；当前不再按“体积变化很小”直接免检。
- 增量提档依赖同一套 `ffmpeg` 解析链路：压缩后的自动评估、手动评估、提档后的再次评估都应使用同一份内置或外部 `ffmpeg`。
- 对 ffprobe 明确标记为零时长/单帧的异常音频轨道，压缩会保留原视频字节和关系，不静音、不删除，也不把它伪装成已编码结果；报告会标记 `unusable_audio_stream_preserved`，并按原媒体计入容量。
- 对未主动降帧的 VFR 视频，编码会保留源时间戳，避免按名义帧率重采样而静默丢帧；只有档位明确要求降帧时才使用目标帧率并执行对应帧数校验。
- 媒体压缩 GUI 分别提供视频和图片 SSIM 基础阈值并跨会话保存：视频预设默认高 `0.95`、均衡 `0.93`、低体积 `0.90`，图片默认 `0.99`。用户手动设置的视频阈值跨档位和会话保留；智能压缩会结合显示面积、复用次数和内容类型计算实际阈值。
- 图片预设新增 `PNG 无损`：只对 PNG 做无损重压，JPEG、GIF、SVG、EMF、WMF、WDP 等保持原字节；重压结果不更小时也自动保留原图。适合先安全清理历史 PPTX 中未优化的超大 PNG。
- GUI 默认使用 `PNG 无损`。外部视频自动匹配出现歧义时不会擅自归族；可在视频库列表选中正确视频族后使用 `添加版本`，验证通过后再按需 `设为高清源`。
- 四个 GUI 工作区会在 `QMainWindow` 构造前自行初始化统一 UI 字体；离屏截图、自动化脚本或直接实例化窗口时，不应再出现 `Sans Serif` alias 告警。
- 字体：水印绘制使用内置 Noto Sans SC；源 PPTX 字体补齐只有在手动开启时才会替换缺失字体，macOS 默认用 PingFang SC，Windows 默认用 Microsoft YaHei，Linux 默认用 Noto Sans CJK SC。

## 文档及媒体水印导出链路

按输入类型和目标类型，当前实际支持如下：

| 输入 | 可编辑 PDF | 图片化 PDF | 可编辑 PPTX | 图片化 PPTX |
| --- | --- | --- | --- | --- |
| `PPTX` | 支持 | 支持 | 支持 | 支持 |
| `DOCX` | 支持 | 支持 | 不支持，固定导出为 PDF | 不支持，固定导出为 PDF |
| `PDF` | 支持 | 支持 | 不支持，固定导出为 PDF | 不支持，固定导出为 PDF |

补充规则：

- 混合队列里只要存在 `PPTX`，顶部全局目标格式仍可选 `PPTX`；但 `DOCX/PDF` 条目会固定按 `PDF` 导出。
- 只有当整个队列都没有 `PPTX` 时，顶部 `格式` 才会锁定为 `PDF`。
- `视频回贴` 仅对 `PPTX + 图片化 PPTX` 生效；对 `DOCX/PDF` 始终无效。
- GUI 控件按“最大文案可见、不互相遮挡”为优先级处理；文档及媒体水印导出顶部格式、形式、质量和视频保留控件在常用窗口宽度下保持一行展示。

## 发布原则

- 该项目是聚合项目，不替代两个原项目的历史发布。
- 当前以本地验证为准；CI 和 release workflow 都只允许手动触发，普通提交和标签不会自动消耗 GitHub Actions 时长。
- 公开文档和 workflow 不写本机绝对路径。
- 开源仓库可以按 MIT 发布；DMG/EXE 不是“仅 MIT”产物，必须另外通过 Qt、FFmpeg、PDFium、Python 及平台原生库的二进制许可门禁。
- macOS 包声明最低系统版本 13.0；默认使用 ad-hoc 签名。配置 `PPTX_TOOLS_CODESIGN_IDENTITY` 后可做 Developer ID + Hardened Runtime 签名，另配置已创建的 `PPTX_TOOLS_NOTARY_PROFILE` 可在 DMG 生成后自动公证并 stapler。
- Windows 正式公开包使用 onedir portable ZIP；one-file 仅限私下测试。
- 手动触发的 Release workflow 从固定且校验哈希的 FFmpeg/x264 源码构建内置运行时；构建后检查内置二进制、启动 GUI、验证 DMG/签名结构，并生成对应源码包、SBOM、审计报告及 `SHA256SUMS-*.txt`。

### 本地发布前审计与压缩基准

除远程 workflow 外，提供不依赖 GitHub Actions 的本地入口，在每台构建主机上手动运行：

- `python scripts/release_audit.py --check`：检查 Git 提交溯源/干净工作树、`uv lock --check` / locked sync、可选 `pip-audit`（外部工具；从 `uv.lock` 导出当前平台全部 extra 的带哈希依赖后扫描，缺失时降级）、可选 CycloneDX SBOM、`dist/` 产物版本/哈希。详见 `docs/RELEASE.md`。
- `python scripts/run_compression_benchmark.py --manifest ...`：以 manifest 驱动复现智能目标容量压缩，输出目标容量误差、实际容量、质量/结构结果、纠偏轮数、耗时与 CPU/GPU/回退信息。语料不进 Git，契约见 `docs/COMPRESSION_BENCHMARK.md`。

跨平台产物必须在对应平台构建；macOS 默认 ad-hoc 签名，Developer ID 公证需单独配置；Windows 签名需在 Windows 主机完成，本地审计只记录哈希。

## 参与贡献与安全

- 开始贡献前请阅读 [`CONTRIBUTING.md`](../CONTRIBUTING.md) 和 [`CODE_OF_CONDUCT.md`](../CODE_OF_CONDUCT.md)。
- 安全问题不要提交公开 Issue，请按 [`SECURITY.md`](../SECURITY.md) 私下报告。
- 直接与传递依赖的许可索引见 [`THIRD_PARTY_NOTICES.md`](../THIRD_PARTY_NOTICES.md)，源码与二进制分发边界见 [`LICENSING.md`](LICENSING.md)。
- 开源前的仓库所有者操作与验证边界见 [`OPEN_SOURCE_CHECKLIST.md`](OPEN_SOURCE_CHECKLIST.md)。

本项目源码采用 [MIT License](../LICENSE)。第三方组件仍适用各自许可证。

## 文档及媒体动态压缩评估与清理

- 目标容量使用十进制 MB（`1 MB = 1,000,000 bytes`），对批量中的每个文件分别生效。成品超目标时最多做两轮容量纠偏；低于目标 95% 时最多做一轮质量回补，不填充无意义字节；编码计划不变时跳过重复重试，进度也不会回退。
- 视频阈值按预设默认为高 `0.95`、均衡 `0.93`、低体积 `0.90`，图片为 `0.99`，再按显示面积和复用次数调整。目标安全模式下 SSIM 未达用户阈值会保留当前压缩候选并在报告标记；强制绝对红线或结构/解码审计失败仍恢复原件，因此安全版可能大于目标。
- 安全版未达目标后才显示 `尝试强制版`。点击后还需二次确认，强制版另存并保留安全版；视频绝对红线 `0.90`，照片 `0.96`，截图/文字/线稿/Logo `0.98`，仍不达目标就停止。
- 目标容量默认使用 CPU `libx264` 两遍编码；`目标容量用 GPU` 默认关闭。Windows 可探测 NVENC、QSV、AMF/MF，单素材失败自动回退 CPU。普通预设的自动硬件/仅 CPU/优先 GPU 策略与两类基础阈值均作为共用设置持久化。
- 附带图片首版队列接受 JPEG、PNG、WebP、TIFF、BMP、GIF。不能完整保留动画、多页、透明度或元数据的资源会跳过有损处理并保留原件；不会自动转换格式。
- `画质评估` 支持单文件和批量。若在列表中选中了已完成条目，则只评估选中项；否则评估全部已完成项。
- 勾选 `自动评估优化` 后，普通压缩完成会自动评估；PPTX 中低于阈值的素材会自动提档并复评，独立图片和视频只自动评估。
- `提档优化` 同样支持单文件和批量。它只会对低于阈值的素材重新抽取源媒体并重写上一版压缩 PPTX，不会再生成第三个输出文件。
- 提档策略固定为：`低/中 -> 高`；高保真仍低于阈值时保留高保真压缩结果并停止，不自动恢复整份原始文档。手动设置的视频阈值跨档位和会话持久化。
- `report.json` 只作为当前会话中的评估/提档缓存；移除文件、关闭窗口或会话结束时会自动清理。
- `pptx_audit_*`、`pptx_compact_*`、`pptx_incremental_*` 临时目录会在成功、手动停止或窗口关闭时清理；启动时也会顺带清理 24 小时前的陈旧目录。
- `pptx_output_watermark_*` 临时目录同样会在关闭窗口或导出完成后清理；若异常退出，下一次启动会清理 24 小时前且不再被活跃进程占用的陈旧目录。
- macOS 上 `LibreOffice` 仍是主链路；`PPTX -> PDF` 会回退 `Keynote`，`DOCX -> PDF` 会回退 `Pages`。若系统未授予 Automation 权限，界面会提示前往系统设置，而不是误报“未安装”。
