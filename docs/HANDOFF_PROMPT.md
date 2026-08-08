# 新对话交接词：Doc Media Toolkit

请接手本机仓库：`<repository-path>`。

## 你的第一步

不要凭历史印象修改。先执行并阅读：

```bash
cd <repository-path>
sed -n '1,260p' README.md
sed -n '1,320p' docs/HANDOFF.md
sed -n '1,320p' docs/ARCHITECTURE.md
sed -n '1,260p' docs/RELEASE.md
sed -n '1,260p' docs/UI_DESIGN.md
sed -n '1,240p' docs/QUALITY_GATES.md
sed -n '1,260p' docs/VIDEO_LIBRARY_AGENT.md
git status --short
git log -5 --oneline
```

当前交接基线必须从 `git log`、工作树、最新测试输出和实际产物重新读取，本文不固化易过期的提交号、测试数量、哈希或包大小。CI 与 release workflow 都只允许手动触发。没有 Developer ID 与 notarization 证据时，不得声称已公证。

## 产品目标

项目是一个 PySide6 桌面工具，包含四个页签：

1. `水印导出`：DOCX / PDF / PPTX / 图片 / 视频预览、水印和导出。
2. `动态压缩`：PPTX、图片、视频压缩，视频/图片质量阈值、压缩后自动评估与提档、可选视频入库和 PPTX 回填。
3. `视频库`：视频族、版本、PPTX 形状关联、外部视频匹配、高清源管理、回填、去重整理和安全隔离。
4. `图片库`：独立图片及 PPTX / DOCX / PDF 内嵌图片归档、精确去重、相似候选、元数据整理和安全隔离；当前不做图片回填。

最终目标不是“看起来能运行”，而是：

- 材料不丢；
- 改名、移动、压缩、重复导入后仍能安全匹配；
- 不确定时让用户人工确认；
- 清理不直接删除，能恢复；
- 回填默认另存，不破坏原 PPTX；
- 水印页、压缩页、视频库页、图片库页的字体、颜色、按钮层级和弹窗风格统一；
- 主列表和预览占主要空间，低频操作进入“更多操作”或抽屉，而不是把所有功能隐藏。

## 关键代码入口

- 统一壳：`src/pptx_tools/gui.py`
- 视频库模型：`src/pptx_tools/video_manager.py`
- 视频库健康检查：`src/pptx_tools/video_library_health.py`
- 视频库界面：`src/pptx_tools/video_manager_gui.py`
- 图片库模型/界面：`src/pptx_tools/image_manager.py`、`src/pptx_tools/image_manager_gui.py`
- 资产库共享 UI：`src/pptx_tools/media_manager_ui.py`、`src/pptx_tools/ui_theme.py`
- 水印核心/界面：`src/pptx_output_watermark/`
- 压缩核心：`src/pptx_video_compactor.py`
- 压缩界面：`src/pptx_video_compactor_gui.py`
- 统一 CLI：`src/pptx_tools/cli.py`
- 本地打包：`scripts/build_standalone.py`
- 回归测试：`tests/`

## 数据模型和不可破坏的规则

视频库的权威清单是：

- `$LIB/video-project.json`：视频族、版本、高清源、已知哈希、内容指纹、PPTX 路径别名和形状锚点。
- `$LIB/video-project.json.bak`：最近有效备份。
- `$LIB/media/`：实际视频实体。
- `$LIB/_cleanup/`：隔离而未永久删除的文件。
- `$LIB/reports/`：健康检查和批处理报告。

身份规则：

- `family.id` 是视频族身份。
- `variant.id` 是具体版本身份。
- `source_variant_id` 是当前高清回填源。
- SHA-256 是精确身份；改名和移动本身不改变身份。
- SHA-256 变化时，只能用严格内容指纹辅助匹配：时长、宽高比、五帧画面、亮度、音频证据，且必须唯一。
- 文件名、目录名、时长、分辨率或肉眼相似，单独都不能认定同一视频。
- 多候选、不同音轨、裁剪、时长变化、不可读文件必须进入人工核对，不得自动归族。
- 新导入版本默认是候选，未经用户确认不得自动成为高清源。

关联规则：

- PPTX 不作为视频库实体保存，但 `decks[].assets[]` 保存 PPTX 哈希、路径别名、媒体 part、slide path 和 shape ID 锚点。
- PPTX 改名或移动不应丢关联；视频改名或移动也不应丢关联。
- 压缩后的新视频 SHA-256 会登记为同族已知别名，但仍以视频族当前高清源作为恢复目标。
- 一个视频族可被多个 PPTX 引用；同一实体只保留一份，引用计数必须保留。
- 归并视频族、删除/隔离版本、切换高清源时，必须迁移或校验全部 PPTX 引用；发现无法迁移必须阻止操作并提示。
- 回填 PPTX 默认另存。覆盖原文件只有在用户明确确认、临时 ZIP 成功、输出可读、媒体关系和锚点校验全部通过后才允许。
- WMV/AVI 等旧容器迁移到 MP4 时，必须同步关系和 Content Type；不得只改扩展名。
- 不得直接编辑 `video-project.json`，不得用 Finder、`mv`、`rm` 或批量重命名器修改库内视频。
- 整理视频库只能先移动到 `_cleanup/`，存在引用或索引异常时禁止永久清空。
- 视频与图片隔离均先写 `moving` 意图，再移动文件，再提交 `quarantined`；重启只在 SHA-256 一致时恢复。原路径和隔离路径必须分别位于库内媒体目录与 `_cleanup/`。

图片库的权威清单是 `$IMAGE_LIB/image-project.json`，实体位于 `images/`。文件 SHA-256 用于精确去重，解码像素哈希和 dHash 只用于生成相似候选；未经确认不得合并。PPTX/DOCX 只提取实际关系引用，PDF 只提取数字内嵌图像，当前不做图片回填。

## 当前已实现的主要能力

### 视频库

- PPTX 视频提取、高清源入库、PPTX 形状关联。
- 外部视频导入、精确哈希复用、严格内容指纹候选匹配。
- 视频族版本管理、手工核实、设置高清源、重命名、移动。
- PPTX 改名/移动、压缩 PPTX 再导入、外部重编码视频的关联维护。
- 视频封面/第一帧预览、播放、匹配候选和人工确认。
- 库体检、关联记录导出、丢失文件检查。
- 整理视频库：按 SSIM、时长、音轨、分辨率、码率、编码和体积给出保留建议；按组安全隔离；支持还原。处理方式为互斥单选，时长/音轨/内容一致性锁定项不会自动清理；族内只有在保留已核实版本、勾选独立的“人工确认：连锁定版本也移入待清理”复选项并二次确认后才能人工强制隔离，切换到其他处理方式会自动解除该复选项，跨族不允许强制归并。
- 归并、隔离异常和清理均有引用检查，不得静默破坏关联。

### 当前视频库界面基线

- 主列表是主要工作区，行显示视频族/版本、分辨率、时长、大小、哈希、关联/状态、文件位置。
- 工具栏保留常用操作：`核实版本`、`添加版本`、`整理视频库`；低频操作放入 `更多操作`。
- `播放`不常驻主工具栏，双击条目或详情抽屉内播放。
- 单击列表条目打开覆盖式右侧详情抽屉，不压缩列表列宽；抽屉包含封面/播放器、高清源标签、分辨率/时长/大小卡片、关联、位置、哈希、状态和底部操作。
- 搜索、统计筛选保持一行；统计数字可点击筛选；表头可排序。
- `关联/状态`必须尽量单行显示，不能因为按钮或列宽设计随意换行。
- 主导航使用“橙色文字 + 居中短下划线”，不使用胶囊式激活页签。

### 当前图片库界面基线

- 图片列表和预览是主工作区；支持精确重复、相似、过小、无来源等状态筛选。
- 名称、分类、标签和说明可人工编辑；AI 只给建议，Base URL 只接受安全 HTTP(S) 地址，API Key 不持久化。
- 移除、合并、孤儿清理全部进入可恢复隔离区，原文档和原始导入文件不被修改。

### 水印/压缩界面基线

- 统一深海军蓝底色、橙色主操作、蓝灰控件和一致字体层级。
- 水印页：左侧文件列表，中间效果预览，右侧导出设置；PPT 横向预览优先显示当前页和下一页，底部缩略图横向滚动；纵向 A4/PDF 显示一页并保留缩略图。
- 中间预览必须按实际页面比例适配，不能裁切关键内容；横向页不应被强行当成纵向页。
- 日志通过底部状态/抽屉按需展开；收起时不能留下大块空白。
- 高级设置应紧跟模板说明，不应单独占用无意义的一整行。
- 右侧按钮按“主要动作 / 常用选项 / 高级设置”分组并对齐，文字不能被遮挡。
- 动态压缩页以文件/结果列表为核心，保留视频质量、图片质量、视频库、入库分类、覆盖原 PPTX、自动评估等现有逻辑。
- 智能目标容量压缩的 PPTX 与附带图片/视频核心已在正式主线：共用队列和设置，视频阈值按预设为高 `0.95`/均衡 `0.93`/低体积 `0.90`、图片 `0.99`，最多两轮容量纠偏和一轮质量回补；目标安全模式 SSIM 未达用户阈值保留压缩候选并报告，强制绝对红线及解码/结构失败恢复原素材；安全版超目标后才允许二次确认生成独立强制版。明确零时长/单帧音轨的媒体必须保留原字节并计入不可压缩预算，报告标记 `unusable_audio_stream_preserved`；未主动降帧的 VFR 视频必须保留源时间戳。DOCX/PDF/XLSX 后端已合并 `main` 并随包分发，正式 GUI 压缩页与 CLI 已接入（`compact_input_path` 按后缀路由），队列含这三类文件时要求显式目标容量且图片预设不能为"不压缩"；PDF 非 ZIP 结构，压缩后自动画质评估跳过并记录日志。若另行构建格式实验应用，应用名、bundle ID、QSettings、日志、临时目录、dist 路径和 `_experimental` 输出必须隔离。

## 视觉规范

唯一规范：`docs/UI_DESIGN.md`。

核心值：

- window `#0b1017`
- surface `#0f1720`
- surface-raised `#111827`
- control `#18212d`
- border `#273244`
- border-strong `#334155`
- text-primary `#f8fafc`
- text-secondary `#cbd5e1`
- text-muted `#94a3b8`
- accent `#f97316`
- selection `#12385f`
- success `#22c55e`
- warning `#f59e0b`
- danger `#ef4444`

字号：

- 统一壳标题：18 px
- 页面/弹窗标题：16 px
- 分区标题：13 px
- 主要控件和正文：12 px
- 辅助信息、表头和状态：11 px

控件高度通常 30 px，紧凑工具栏 28 px，主要确认按钮 40 px；普通表格行 46-48 px。不得在页面内另建私有字号体系，也不要为了塞更多按钮继续缩小字号。

## 测试和打包

先运行：

```bash
cd <repository-path>
QT_QPA_PLATFORM=offscreen .venv/bin/python scripts/run_tests_isolated.py
.venv/bin/ruff check src tests scripts
.venv/bin/ruff format --check src tests scripts
.venv/bin/python -m compileall -q src scripts tests
.venv/bin/pip check
git diff --check
```

macOS 本地 DMG：

```bash
.venv/bin/python scripts/build_standalone.py \
  --gui --clean --target-platform macos \
  --bundle-ffmpeg --require-ffmpeg-bundle \
  --dmg \
  --dmg-output "<output-directory>/Doc Media Toolkit-macOS-arm64.dmg" \
  --name "Doc Media Toolkit"
```

打包后至少检查：

```bash
shasum -a 256 "<output-directory>/Doc Media Toolkit-macOS-arm64.dmg"
hdiutil verify "<output-directory>/Doc Media Toolkit-macOS-arm64.dmg"
codesign --verify --deep --strict --verbose=2 "/path/to/Doc Media Toolkit.app"
```

不要未经用户要求推送或触发 GitHub 构建；如果用户明确要求推送，再推送到 GitHub。

## 接下来工作的标准顺序

1. 读取本文和现有核心文档，尤其是 `ARCHITECTURE.md`、`QUALITY_GATES.md` 和 `UI_DESIGN.md`。
2. 核对 `git status`、最新提交、当前测试和实际 DMG。
3. 如果是 UI 工作，先对照 `docs/UI_DESIGN.md` 和当前运行界面，先判断现有代码是否已有对应组件，禁止重新造一套样式；公开截图只能使用匿名合成数据。
4. 如果是数据/资产库工作，先只读 `videos doctor` / `images doctor` 和报告，确认缺失、修改、异常、未关联和待清理数量，再决定写操作。
5. 修改前说明范围；修改后测试同类路径，不只测试用户指出的一个按钮或一个文件。
6. 涉及覆盖、归并、隔离清空或删除时，先生成报告和备份，明确显示影响的 PPTX、视频族、版本和引用数量。
7. 最终报告必须包含：改动文件、测试命令/结果、未验证边界、提交号、是否推送、DMG 路径和 SHA-256。

## 给新对话的直接指令

请把上面的仓库、文档、数据模型和验收规则作为事实基线。先核实再改，不要凭印象，不要重写已经存在的匹配/关联逻辑，不要把文件名当身份，不要直接删除源文件，不要默认覆盖 PPTX，不要触发 GitHub 构建。任何不确定匹配都要保留原文件并进入人工确认。UI 调整优先恢复已确认的深色海军蓝 + 橙色下划线风格，保持主列表/预览优先、文字可读、按钮不遮挡，并通过离屏截图和测试验证。
