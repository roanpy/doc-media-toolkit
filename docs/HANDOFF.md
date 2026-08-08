# Doc Media Toolkit Handoff

## 目标

把 `文档及媒体水印导出`、`文档及媒体动态压缩`、`PPTX 视频资产库` 和
`文档图片资产库` 整合为独立项目 `Doc Media Toolkit`，通过页签切换功能，并共用
依赖、资源、日志、测试和打包逻辑。

## 当前状态

已完成：

- 公开项目骨架：`doc-media-toolkit`（Python 包和 CLI 保留 `pptx-tools` 兼容名）。
- 已复制水印导出核心代码：`src/pptx_output_watermark/`。
- 已复制并扩展媒体压缩核心代码：`src/pptx_video_compactor.py`。
- 已复制并扩展媒体压缩 GUI：`src/pptx_video_compactor_gui.py`。
- 已复制水印字体、图标、媒体压缩视频预设配置。
- 新增统一 GUI：`src/pptx_tools/gui.py`。
- 新增统一 CLI：`src/pptx_tools/cli.py`。
- 新增统一打包脚本：`scripts/build_standalone.py`。
- 新增 PPTX 视频资产库核心与 GUI：`src/pptx_tools/video_manager.py`、`src/pptx_tools/video_manager_gui.py`。
- 新增文档图片资产库核心与 GUI：`src/pptx_tools/image_manager.py`、`src/pptx_tools/image_manager_gui.py`。
- 资产库共用 worker、样式和字体基础设施位于 `src/pptx_tools/media_manager_ui.py` 与 `ui_theme.py`；视频/图片 GUI 不互相导入。
- 修正媒体压缩模块复制到 `src/` 后的 `config/` 资源定位。
- 架构、依赖方向、数据不变量和扩展规则统一记录于 [`ARCHITECTURE.md`](ARCHITECTURE.md)。

## 结构

```text
doc-media-toolkit/
  assets/
    app_icon.*
    fonts/NotoSansSC[wght].ttf
  config/
    default.json
    balanced.json
    high.json
    aggressive.json
  src/
    pptx_tools/
      gui.py
      cli.py
      ui_theme.py
      media_manager_ui.py
      video_manager.py
      video_manager_gui.py
      image_manager.py
      image_manager_gui.py
    pptx_output_watermark/
      ...
    pptx_video_compactor.py
    pptx_video_compactor_gui.py
  scripts/
    build_standalone.py
    check_public_safety.py
    prepare_icons.py
    pyinstaller_runtime_hook.py
```

## 统一 GUI

入口：

```bash
pptx-tools-gui
```

实现方式：

- `src/pptx_tools/gui.py` 创建主窗口 `Doc Media Toolkit`（中文界面名 `文档媒体工具箱`）。
- 通过 `QTabWidget` 加四个页签：
  - `文档及媒体水印导出`
  - `文档及媒体动态压缩`
  - `PPTX 视频资产库`
  - `文档图片资产库`
- 每个页签实例化原项目的 `MainWindow`，取出 central widget 嵌入页签。
- 原窗口对象保存在 `embedded_tools`，避免被 GC，同时保留原信号和 worker 状态。

当前界面原则：

- 外层导航和品牌统一，内部四个工具保留各自任务结构，避免为了“统一”重写稳定核心。
- 水印页以文档预览为主，动态压缩页以文件/结果队列为主，视频资产库以视频族/版本列表为主，图片资产库以图片列表/预览为主；高级设置和日志默认收起。
- 水印页在同一滚动预览区纵向展示当前页和下一页，单页时不出现伪造的第二页背景；白/灰/黑预览底色统一作用于所有可见页面。视频详情的播放使用右侧面板内嵌 QtMultimedia，关闭、切换条目或退出时停止。
- 视频资产库把 PPTX 入库/回填放在列表上方，版本与维护操作分组；单击条目打开覆盖式详情卡，不压缩列表列宽。
- 视频匹配、高清回填、整理、待清理、库体检、帮助和通用确认弹窗使用同一字号、颜色、边距与安全动作层级。
- 在最小支持窗口内保持核心操作可见；超长路径使用省略与 tooltip，不挤压主按钮。
- 当前界面实现以 [`UI_DESIGN.md`](UI_DESIGN.md) 为唯一视觉与交互规范；
  设计图中的示例名称、数量和路径不覆盖现有业务规则。
- 字号、控件高度、圆角和间距只以 `docs/UI_DESIGN.md` 的视觉令牌为准；帮助页和业务弹窗不得维护独立字号体系。

## 统一 CLI

入口：

```bash
pptx-tools --help
pptx-tools watermark --help
pptx-tools compact --help
pptx-tools videos --help
pptx-tools images --help
```

实现方式：

- `watermark` 转发到 `pptx_output_watermark.cli:main`。
- `compact` 转发到 `pptx_video_compactor:main`，支持视频与图片媒体压缩。
- `videos` 转发到 `pptx_tools.video_manager:main`；`add` 原子化归档高清源并登记 PPTX/形状关联，`upgrade` 从库中高清优化 PPTX，`list` 只显示视频族，`doctor` 执行只读健康检查并可输出 JSON 报告。
- `images` 转发到 `pptx_tools.image_manager:main`；当前 CLI 提供图片库只读 `doctor` 和完整哈希复核，建库、导入、整理与恢复通过 GUI 或公共模型方法执行。
- 旧版 `detach/restore/activate/compress` 命令暂时保留兼容已有项目，但不出现在新 GUI 主流程。

## PPTX 视频资产库

- 视频库只保存视频实体，不复制 PPTX 或低码率视频副本；`decks` 保存 PPTX 哈希、主路径/别名、输出记录和幻灯片/形状锚点，供改名、覆盖压缩和回填校验。
- 入库先按原始媒体 SHA-256 精确去重；不同编码或分辨率只有在时长、宽高比、五帧 dHash、亮度和音频频谱均唯一匹配时才归为同一视频族，并把新 SHA-256 登记为别名而不复制低清实体。可指定 `media/` 下的安全相对分类目录。默认只将超过横向/纵向 1080p 包络的源按比例转为 CRF 18 MP4，也可选择保留原片；原始哈希仍登记为别名。
- PPTX 中同族的更高分辨率版本，或同分辨率但码率高 5% 的潜在高质量版本，会保留为候选；只登记普通压缩版哈希，不保存低清副本。候选必须人工确认后才能设为高清源。
- `import_external_video` 支持独立视频：精确哈希复用，唯一内容指纹命中时新增同族候选版本，无可靠唯一命中时可返回排序候选且不写文件。GUI 支持按钮或拖入视频列表批量导入；歧义项用封面、分辨率、时长、画面和音频差异人工确认，可归入现有族、新建族或跳过。只有显式人工确认才允许把非严格匹配版登记为同族已知哈希。匹配候选不再仅凭更高像素自动更新 `source_variant_id`，需在 GUI 核实后点击 `设为高清源`。`QSettings` 记住最近目录。
- Agent/CLI 使用 `videos import-video [--family-id ...]` 走同一哈希/指纹校验，再用 `videos set-source VARIANT_ID` 显式选定高清源；不提供跳过身份校验的公开入口。GUI 视频族行显示版本数和 PPTX 引用数，非当前高清源统一标记为候选。
- `import_variant` 默认验证所选文件与视频族的内容指纹；只有旧版内部兼容场景才可显式使用 `verify_identity=False`，未验证版本不会登记为已知哈希或自动升级高清源。
- 所有用户可见的 PPTX 入库路径统一调用 `archive_and_register_pptx`：先归档高清源，再登记 PPTX 哈希、路径别名和形状锚点；登记失败会回滚本轮新增清单和媒体文件。低层 `archive_pptx_videos` 只供需要“仅处理媒体实体”的内部流程使用。压缩页随后调用 `register_compressed_pptx_hashes` 把输出 MP4 哈希登记为别名；低清文件不复制进视频库。非覆盖输出记入 `optimized_outputs`，且 `register_optimized_output` 同时识别主路径和 `source_aliases`；已登记的压缩输出再次送入工具时，`add_deck` 会按输出 SHA-256 沿用原 PPTX 记录，不创建重复记录。内容指纹唯一匹配且后来输入的有效分辨率更高时，只新增候选版本，不自动更新 `source_variant_id`。`upgrade_pptx_from_library` 成功完成内容指纹匹配和输出校验后，会把输入压缩哈希固化为同族别名。
- `save()` 在原子替换前会用完整 schema 和跨实体引用规则验证下一版清单；当前清单不可读时拒绝直接覆盖，只允许 `open()` 从最近有效备份恢复。`videos doctor`/GUI“库体检”检查哈希唯一归属、实体文件、PPTX 来源、输出历史、未登记媒体和待清理索引；完整模式逐文件重算 SHA-256，并将仅时间戳变化与真实内容变化分开报告。精确哈希若意外归属多个视频族，匹配会失败关闭并要求先人工归并，不会任意选择第一个族。
- `覆盖原 PPTX` 仅允许 PPTX + 视频入库 + 视频压缩 + 图片不压缩 + 自动评估关闭。worker 在确认至少归档一个视频后才把 `output` 指向输入路径；核心通过临时 ZIP 原子替换，成功后用 `adopt_upgraded_deck_source` 刷新整包哈希、媒体 part 和幻灯片/形状锚点。覆盖结果不会与自身做 SSIM 审计。
- `upgrade_pptx_from_library` 优先接受实体哈希或已知压缩别名。未登记的外部重编码视频只有在时长（容差 150-250ms）、宽高比、五帧 dHash、亮度和音频频谱同时满足严格阈值且唯一命中时才自动建议回填。GUI 用 `review_pptx_matches(include_resolved=True)` 提取全部媒体的临时审阅副本，不再只显示未匹配项；PPTX 级确认窗口展示精确匹配、内容匹配、已是高清和未匹配状态，并按需生成 10%/50%/90% 三帧封面对比。每个媒体 part 可保持当前、仅本次覆盖，或确认同一视频并记住哈希。`family_overrides` 可纠正已有自动匹配，`keep_current_media` 明确禁止替换；只有 `remember_manual_matches` 中的人工项才在输出 ZIP 与媒体哈希校验成功后登记别名。`incompatible_only=True` 只迁移 WMV/AVI 等格式，已有 MP4 保持逐字节不变。兼容的 MP4/M4V、H.264/AAC 高清源均直接复用原字节，避免仅因 `.m4v` 扩展名重复转码；同一 PPTX 内多个形状命中同一高清源时共用一个媒体 part，避免重复嵌入造成超大包和 PowerPoint 修复提示。PPTX 包内仍使用 `.mp4` 媒体 part；WMV/AVI 会同步关系与 Content Type，因此幻灯片 XML、形状、海报、尺寸和播放时间线保持不变。
- 视频 SSIM 比较先用 `setpts=PTS-STARTPTS` 归零候选与原片时间轴，避免 AVI/WMV 等旧容器起始时间不同造成错帧评分。视频编码完成后同步校验分辨率、时长、音轨和帧数；`aggressive`/`balanced` 可按档位设计降低帧率，但拒绝目标帧率之外的意外少帧或截短。未主动降帧的 VFR 视频使用 `-fps_mode passthrough` 保留源时间戳；FFprobe 明确报告零时长且最多一帧的异常音频轨道不会进入编码，原媒体按不可压缩字节保留并报告 `unusable_audio_stream_preserved`。
- `scripts/batch_compact_library_pptx.py` 用于已登记视频库的离线批量交付：同族从权威高清源生成候选且不切换当前高清源；必须提供 SSIM `0.90-1.00` 的 `quality-fallback-selection.json`，默认交付使用 `0.95`，轻量可恢复版可显式使用 `0.90`。按 `aggressive → balanced → high → 保留原媒体` 单调提档，未评分、兼容性失败或所有档位不达标的视频族一律禁止替换。达标版本还需比 PPTX 当前媒体至少节省 15%（不兼容容器迁移除外），整份 PPTX 需节省至少 5% 或 10 MiB。脚本只在指定输出目录生成画质达标视频、保持相对目录结构的 PPTX 和可续跑 JSON 报告，并逐份校验 ZIP、非媒体成员、媒体哈希及幻灯片/形状锚点；续跑时先筛选源变化或输出缺失的 PPTX，只准备这些文件涉及的视频族，并可从 `_cleanup/` 精确恢复报告指定的已审计版本，避免整理视频库后无效重编码全库；它不覆盖源 PPTX。
- 已登记 PPTX 原地迁移后调用 `adopt_upgraded_deck_source`：按 `(slide_path, shape_id)` 锚点重新绑定媒体 part，变化后的媒体哈希必须属于原视频族；同时刷新 PPTX 哈希、结构哈希、`original_variant_id` 和主路径/别名。可恢复压缩版使用 `prefer_source_variant=True`，把当前低码率哈希登记为同族别名，但继续以权威高清源作为恢复目标，之后清理低码率实体文件不会破坏回填。一个视频族被多个 PPTX 引用时只保留一份 MP4 版本，各 PPTX 引用计数不变。高质量 WMV 转码使用 `-bf 0`，避免稀疏 VFR 时间戳在 MP4 中因 B 帧重排产生短于最后一帧的容器时长。
- 原始 MP4/H.264/AAC 且不超过横向 1920x1080 或纵向 1080x1920 时按原字节回填；其他容器/编码或超大画面临时转 H.264 Main CRF 18/AAC 256 kbps。缩放上限为 1080p，不放大低分辨率源，保持宽高比并取偶数尺寸。原片模式保证库内源字节不变，但不兼容源回填 PPTX 时仍会转码；默认 1080p 模式不能恢复降采样前的 2K/4K 细节。
- 回填质量档位：`BACKFILL_QUALITY_TIERS`（`video_manager.py:333`）定义 `best`（默认，≤1920×1080、CRF 18、不限码率、AAC 256k、输出后缀 `high_quality`）、`high`（≤1920×1080、CRF 20、≤12 Mbps、AAC 256k、后缀 `hq1080p`）和 `balanced`（横版 ≤1280×720 / 竖版 ≤720×1280、CRF 23、≤5 Mbps、AAC 128k、后缀 `balanced720p`）。档位是上限语义：源已满足档位规格（MP4/M4V 容器、H.264、AAC 或无音轨、偶数尺寸且在包络内、码率不超上限）即原字节嵌入，否则按档位转码；码率探测缺失按超限转码（安全方向）。执行路径 `_delivery_master`（`video_manager.py:2012`）与 GUI 预览 `plan_backfill_action`（`video_manager.py:397`）共享同一 `_backfill_compatibility` 判定（`video_manager.py:374`），避免决策漂移。`upgrade_pptx_from_library(..., quality_tier=...)`（`video_manager.py:2051`）按档位生成带后缀的默认输出名；转码母版只把内容哈希登记为视频族别名，不新增版本。GUI 确认对话框 `PptxUpgradeReviewDialog`（`video_manager_gui.py:366`）内置档位选择并由 `QSettings` 记住，输出位置选择移到确认之后以携带档位后缀。转码采用 CRF + maxrate/bufsize 封顶 VBR，并依赖 FFmpeg ≥ 5.1 的 `-fps_mode passthrough` 保留 VFR 帧；明确零时长且至多一帧的异常音轨不进入编码。
- 视频族记录 `source_variant_id` 和 `known_hashes`。手工设定高清源、PPTX 改名和移动均不会影响匹配；视频族归并会合并所有哈希别名。该模型不维护逐 PPTX 历史，因此高清优化使用视频族当前最佳源，而非恢复某份 PPTX 当年的逐字节媒体快照。
- GUI 重命名视频族时同步重命名权威高清源文件，选择视频族移动时默认移动其高清源。文件被 Finder 手工改名或移到下级目录后仍可按 SHA-256 递归重新关联。`video-project.json.bak`、写锁和 revision 冲突检测仍保留。
- `QSettings` 保存最近打开的视频库、压缩页入库模式、母版质量、最近分类目录及列表排序方向；可按客户或主题新建多个独立目录，并通过“切换 / 打开视频库”切换。当前库栏显示库名和完整路径，悬停同时说明库内清单/媒体/隔离/报告、全局偏好和共享日志位置。压缩页显示当前入库目标并允许直接切换，重新激活页签时会同步视频库页的选择。视频库列表支持按名称、路径、分辨率和哈希片段即时筛选，显示视频族/版本、无关联、多版本和异常统计，点击统计数字可直接筛选；“待核对”是无关联、多版本和异常的去重并集，无关联和多版本通常不是错误，文件异常才表示库内文件不可用。选中视频族后可用“核实版本”只扫描该族及其跨族重复候选；人工“归并视频”会先显示将迁移的版本、哈希别名、PPTX 和媒体引用数量。`merge_families` 默认要求内容身份已由哈希/指纹证明，GUI 人工确认路径显式传入 `confirmed_same_content=True`；归并会迁移所有 deck asset 的 `family_id`/`original_variant_id`，最终仍由 `save()` 的全清单引用校验兜底。外部视频版本保存 `origin_paths`，同哈希文件从新名称或目录再次导入时追加来源但不复制文件，悬停文件位置可查看。任一表头均可排序并切换方向，零结果时显示恢复全部列表的指引；详情抽屉使用 QtMultimedia 在应用内播放。共享滚动日志写入系统应用数据目录，不在视频库内持久化详细操作历史。
- 整理视频库：`scan_cleanup_groups` 分两步——族内版本先核实派生关系或内容指纹，跨族重复再按当前文件 SHA-256 交集或内容指纹聚类；带音轨的跨族指纹匹配还必须通过解码音频相关性。`focus_family_id` 仍从全库寻找跨族候选，但只对涉及选中族的组做昂贵的 SSIM 评估，供“核实版本”快速使用。逐候选算 `_ssim_videos`（复用画质审计的 fps=1 scale2ref SSIM 滤镜）和 `_fingerprint_confidence`。推荐保留遵循"时长与音轨一致、SSIM 达阈值、分辨率充足，再比体积"的规则，不只看码率。跨族组会压制同轮重叠的族内组，归并后的剩余版本留到下次扫描。`apply_cleanup_plan` 执行前再次校验版本归属与内容身份，生成统一版、归并和移除完成后只保存一次 manifest；移除实体版本时仍把其 SHA-256 保留为视频族别名，确保压缩 PPTX 可继续高清回填。文件隔离采用 `moving → quarantined` 意图日志：先原子写索引、再移动文件、再提交状态；重启只在目标哈希吻合时完成恢复。新索引保存库内相对路径，旧绝对路径仍兼容；原路径必须位于 `media/`，隔离路径必须位于 `_cleanup/`。失败时只回滚本轮产生的条目，避免覆盖其他进程的新索引。`restore_cleanup_entry` 可还原；索引损坏、内容变化或路径越界时失败关闭；`cleanup_pending_issues` 校验无引用残留后才允许 `empty_cleanup`。不同音轨、被裁剪、来源不明或置信度不足时 `safe_to_apply=False`，GUI 整组默认跳过；处理方式使用互斥单选，族内如确实要隔离锁定版本，必须在“保留勾选版本，其他移入待清理”模式下显式勾选独立复选项“人工确认：连锁定版本也移入待清理”、保留已核实版本并二次确认，跨族归并拒绝强制。统一 1080p 由 `create_unified_version` 从登记的权威源重新编码生成。
- `quarantine_abnormal_variant` 是单个异常文件的快速安全入口：仅允许处理非高清源、非当前版本、零 PPTX 引用且不可读的版本，复用 `_cleanup/` 索引并支持恢复；不直接删除文件。

## 文档图片资产库

- `image-project.json` 是图片库权威清单，`images/` 保存实体，`_cleanup/` 保存可恢复隔离文件；原始 PPTX、DOCX、PDF 和独立图片不被修改。
- 导入先按文件 SHA-256 精确复用，再记录解码像素哈希与 dHash 供相似候选审阅；不同内容不能仅凭视觉相似自动合并。
- PPTX/DOCX 只提取关系文件实际引用的图片；PDF 只提取数字内嵌图像，不进行整页截图或 OCR。当前不提供图片回填。
- 移除、合并和孤儿整理使用与视频库相同的 `moving → quarantined` 崩溃恢复协议；原路径限制在 `images/`，隔离路径限制在 `_cleanup/`，哈希不匹配时失败关闭。
- OpenAI 兼容接口仅提供名称、分类、标签、摘要和疑似合并建议。Base URL 只允许 HTTP(S)，不得内嵌凭据、查询或片段；响应体受限，API Key 不持久化，所有合并/删除仍需人工确认。

## 打包

统一脚本：

```bash
python scripts/build_standalone.py --help
```

macOS GUI：

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

如需指定 DMG 输出路径：

```bash
python scripts/build_standalone.py \
  --gui \
  --target-platform macos \
  --bundle-ffmpeg \
  --require-ffmpeg-bundle \
  --dmg \
  --dmg-output "/absolute/path/Doc Media Toolkit-macOS-arm64.dmg" \
  --name "Doc Media Toolkit"
```

Windows GUI one-file：

```powershell
python scripts\build_standalone.py `
  --windows-onefile `
  --bundle-ffmpeg `
  --require-ffmpeg-bundle `
  --name "Doc Media Toolkit"
```

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

补充说明：

- `--windows-onefile` 只负责 Windows GUI。
- Windows CLI 仍使用 `--cli --onefile`，但必须在 Windows 主机上执行。
- `--dmg` 只适用于 macOS GUI onedir 构建；生成的 DMG 内只应包含 `Doc Media Toolkit.app` 和 `Applications` 快捷方式。

FFmpeg 环境变量：

```text
PPTX_TOOLS_FFMPEG
PPTX_TOOLS_FFPROBE
PPTX_TOOLS_SOFFICE
PPTX_TOOLS_WPS
PPTX_TOOLS_WPP
```

脚本也兼容旧变量：

```text
PPTX_OUTPUT_WATERMARK_FFMPEG
PPTX_OUTPUT_WATERMARK_FFPROBE
PPTX_VIDEO_COMPACTOR_FFMPEG
PPTX_VIDEO_COMPACTOR_FFPROBE
```

## 依赖策略

合并后共用：

- `PySide6`
- `ffmpeg` / `ffprobe`，仅媒体压缩中的视频处理需要；图片压缩只依赖 Pillow
- 媒体压缩的画质评估同样依赖 `ffmpeg`，因为通过 `ffmpeg + SSIM` 做逐素材对比；当前不再按体积变化直接免检
- 媒体压缩 GUI 将 SSIM 基础门槛拆成独立且持久化的“视频≥”和“图片≥”：视频预设默认高 `0.95`、均衡 `0.93`、低体积 `0.90`，图片 `0.99`；用户手动设置的视频阈值跨档位和会话保留。智能目标压缩再结合显示面积、复用次数和内容类型计算实际阈值；自动评估与提档仍按素材类型取对应门槛。
- 智能目标容量核心已进入正式主线：十进制 MB、每文件独立目标、联合视频/图片预算、双尺度质量审计、最多两轮纠偏、一轮质量回补和计划签名去重均已接入。目标安全模式 SSIM 未达用户阈值时保留当前压缩候选并写入 `below_threshold`；强制绝对红线以及解码、结构或元数据审计失败仍恢复原素材。安全版超目标后才显示需二次确认的强制版入口；强制版另存且仍受绝对红线约束。目标模式默认 CPU 两遍，目标 GPU 默认关闭；普通预设的自动/CPU/GPU 策略为共用持久设置。独立图片/视频沿用同一拖入队列，不单设接口。
- 实验构建使用 `--experimental`，应用名、bundle ID、QSettings、日志、临时目录、构建目录和默认输出均与正式版隔离。DOCX/DOCM、PDF（含扫描件）、XLSX/XLSM 图片压缩后端已进入正式 GUI/CLI；这些格式要求显式目标容量和启用图片预设，并保留格式专用结构/布局门禁。
- 图片 `PNG 无损` 档只重新封装 PNG，其他图片格式原字节复制；若新 PNG 不更小则回退原字节。质量审计遇到全透明或纯色 PNG 的 FFmpeg `SSIM=0/NaN` 时，仅在尺寸一致且解码 RGBA 像素完全相同时回退为 `1.0`。
- 压缩 GUI 默认图片档为 `PNG 无损`。外部视频自动匹配歧义时保持未导入，用户可选中正确视频族后用 `添加版本` 完成人工归族，再显式 `设为高清源`；禁止仅凭文件名强制合并。
- 四个 GUI 工作区会在 `QMainWindow` 构造前先跑统一字体初始化；离屏截图、自动化脚本或直接实例化窗口时，不应再冒 `Sans Serif` alias 告警
- `pypdfium2`
- `python-pptx`
- `pypdf`
- `reportlab`
- `Pillow`
- `comtypes`，仅 Windows

不内置：

- Microsoft Office
- WPS Office
- LibreOffice
- Keynote，macOS 兜底导出 PDF 使用
- Pages，macOS DOCX 兜底导出 PDF 使用

这些仍作为运行时 PDF 导出引擎检测。

文档及媒体水印导出当前链路：

- `PPTX -> PDF`
  - Windows：PowerPoint COM -> WPS COM -> LibreOffice
  - macOS：LibreOffice -> Keynote
- `DOCX -> PDF`
  - Windows：Word COM -> WPS COM -> LibreOffice
  - macOS：LibreOffice -> Pages
- `PDF` 输入不做二次文档转 PDF，直接进入 PDF 水印或图片化流程。
- `DOCX/PDF` 不支持导出为 `PPTX`；在混合队列中会固定生效为 `PDF`。

媒体压缩评估与提档：

- `画质评估` 支持单文件和批量，对已处理媒体逐项跑 SSIM。
- `自动评估优化` 会在本轮压缩后自动评估；PPTX 低分素材会自动提档并复评，独立图片和视频只自动评估。
- `WMV` 或旧编码视频只有在当前打包或系统 `ffmpeg` 能解码时才能跑 SSIM；无法解码时应显示明确的 FFmpeg 解码/评分失败原因，而不是泛化为“未找到评分”。
- `提档优化` 支持单文件和批量，只会回源抽取低分素材并覆盖上一版压缩 PPTX，不再生成第三个输出文件。
- `提档优化` 的策略固定为：`低/中 -> 高`；高保真仍低于阈值时保留压缩结果并停止自动提档，提示降低阈值或替换源素材。只有解码、结构或元数据审计失败才恢复原素材。
- `report.json` 是当前会话的评估/提档缓存，移除文件、关闭窗口或会话结束时自动清理。
- `pptx_audit_*`、`pptx_compact_*`、`pptx_incremental_*` 临时目录会在成功、停止、关闭窗口时清理，启动时也会清掉 24 小时前的陈旧目录。
- `pptx_output_watermark_*` 临时目录也纳入同一套清理：正常完成即时清理；异常退出残留会在下次启动时按 24 小时阈值和进程存活状态回收。
- macOS 文档导出链路保持 `LibreOffice -> Keynote/Pages`。依赖检测现在会区分“未安装”和“缺少 Automation 权限”，并支持从 GUI 直接打开系统设置。
- 最近一轮 GUI 收尾以“控件不互相遮挡、最大文字可展示”为准：文档及媒体水印导出顶部 `格式/形式/质量/保留视频` 控件保持一行；预览区底部和右侧日志区对齐；媒体压缩的 `质量阈值` 数字框加宽并补了上下箭头绘制。评估行不再用固定窗口断点，而是按可用宽度和控件 `sizeHint` 决定同行或分行；宽窗口评估/操作控件整组右对齐并使用 16 px 间距，窄窗口仍把操作按钮移到标题行；统一壳嵌入页保留 6 px 顶部安全内边距以露出首个面板上边框；880–1440 px × 560–700 px 的四工作区几何矩阵由回归测试覆盖。

## 下一步建议

1. 在新对话先运行基础校验：

```bash
python -m py_compile pptx_tools_cli.py pptx_tools_gui.py scripts/build_standalone.py
PYTHONPATH=src python -m pptx_tools.cli --help
PYTHONPATH=src python -m pptx_tools.cli watermark --help
PYTHONPATH=src python -m pptx_tools.cli compact --help
```

2. 本地启动 GUI 验证四个页签能打开：

```bash
PYTHONPATH=src python pptx_tools_gui.py
```

页签切换会保留四个工具各自的设置和日志；退出时通过协作取消终止外部进程并等待 worker、preview、audit 线程结束，不再使用 `QThread.terminate()`。

3. 分别用小 PPTX 跑：

- 水印导出 PDF editable。
- 水印导出 PPTX image。
- 媒体压缩 high/high。

4. 再做 macOS 本地打包验证。

5. Release workflow 已包含三平台测试、Windows FFmpeg 固定版本与哈希校验、GUI 冒烟、DMG/签名结构检查及产物 SHA-256 清单；后续发布前关注实际 Actions 运行结果。

## 注意事项

- 旧两个项目仍保留，`doc-media-toolkit` 不依赖它们。
- 后续同步旧项目修复时，需要手动复制或 cherry-pick 到 `doc-media-toolkit`。
- 不要在公开文档、workflow 或脚本中写本机绝对路径。
- Windows 上 FFmpeg 用 essentials，避免 full static 导致 one-file 暴涨。
- macOS 当前只做本地可运行包；公证分发需要 Developer ID 和 notarization secrets。
