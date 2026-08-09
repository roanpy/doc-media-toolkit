# Doc Media Toolkit 质量目标与验收门禁

本文定义 `0.2.x` 的最终交付目标。它不是功能愿望清单；进入提交和打包前，
每一项必须有代码、测试、只读报告或人工界面验证作为证据。

## 1. 产品目标

- 一个桌面应用完成水印导出、动态压缩、PPTX 视频资产管理和文档图片管理。
- 默认路径安全：不覆盖输入、不直接删除库内媒体、不自动确认歧义内容。
- AI 不是数据一致性的前提。没有 AI 时，程序仍能完成哈希去重、保守内容匹配、
  人工核对、关联迁移、隔离恢复和高清回填。
- AI 或外部 Agent 只辅助语义命名、分类和复杂候选判断，所有写入仍走公开 API、
  CLI 或 GUI。

## 2. 数据不变量

1. 每个视频族至少有一个版本，`source_variant_id` 与 `active_variant_id`
   必须指向本族版本。
2. 每个 PPTX 媒体引用必须同时指向有效视频族和该族内有效版本。
3. 精确媒体哈希只能唯一归属一个视频族；冲突时失败关闭，不任意选择。
4. 清单写入采用锁、revision 检查、临时文件、完整下一版校验和原子替换。
5. PPTX 入库必须同时完成媒体归档和 PPTX/形状登记；任一步失败回滚本轮新增项。
6. 归并或移除版本时，所有 PPTX 引用必须迁移到目标族的有效保留版本。
7. 待清理文件先进入 `_cleanup/`；存在任何清单引用或索引异常时禁止永久清空。
8. 高清回填默认另存，并校验 ZIP、媒体哈希、关系、幻灯片/形状锚点和输出路径。
9. 图片按文件 SHA-256 和解码像素哈希去重；不同内容的相似图只能人工确认合并。
10. 图片移除、合并和孤儿清理先进入 `_cleanup/`，来源记录必须迁移到保留图片。
11. 视频和图片隔离均采用 `moving → quarantined` 意图日志；中断恢复必须核验
    SHA-256，原路径只允许位于受管媒体目录，隔离路径只允许位于 `_cleanup/`。
12. 新建隔离索引只保存项目相对路径；读取仍兼容旧绝对路径，但不得允许库外路径。
13. AI Base URL 只允许无内嵌凭据、查询或片段的 HTTP(S) 地址，响应体必须有上限；
    API Key 不得持久化或出现在错误消息、日志和 URL 中。

## 3. 匹配与人工边界

- 自动精确匹配：SHA-256 唯一命中。
- 自动内容匹配：时长、宽高比、五帧画面、亮度和音频证据同时满足严格阈值，
  且只能有一个候选。
- 人工核对：多候选、不同音轨、裁剪/时长变化、低置信度或不可读媒体。
- 人工确认提供封面对比、完整播放、分辨率、时长、画面差异和音频结论。
- 仅本次替换不得污染身份；只有“确认同一视频并记住”才能登记新哈希别名。

## 4. 交互验收

- 四个页签的产品抬头、标题、说明和帮助文案一致。
- 统一壳嵌入页保留 6 px 顶部安全内边距，首个面板上边框在 macOS/Windows
  缩放和默认尺寸下均可见。
- 在最小支持窗口内，主操作不截断；列表/预览是主要空间，日志可收起。
- 统一壳的四个工作区必须在 880/1000/1100/1180/1280/1440 px 宽度与
  560/620/700 px 高度的组合下通过离屏几何检查；可见交互控件不得越出当前页边界。
- 压缩页评估控件按实际可用宽度与 `sizeHint` 合并或分行；同行控件垂直中心偏差
  不得超过 1 px，隐藏的强制输出按钮不得占宽。
- 所有可见交互控件的实际高度不得小于当前样式的 `sizeHint`；压缩结果表头与评估
  行至少保留 8 px 间距。该检查必须覆盖统一壳嵌入后以及 880–1440 × 560–700
  的离屏尺寸矩阵。
- 文件入口支持按钮、多选和拖放；拖放区域有明确文案与可访问描述。
- 视频资产库状态数字可点击筛选；表头可按真实值排序；零结果有恢复指引。
- 高风险操作说明影响范围、默认取消并提供恢复路径。
- 长路径使用省略显示和完整 tooltip；按钮状态、加载进度、失败原因和输出位置可见。
- 中英文 README、`HANDOFF.md`、`UI_DESIGN.md`、`RELEASE.md`、
  `ARCHITECTURE.md`、开源政策与第三方许可说明和实际入口、页签、依赖方向、
  数据协议及语言覆盖一致。

## 5. 自动化门禁

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python scripts/run_tests_isolated.py
ruff check src/ tests/ scripts/
ruff format --check src/ tests/ scripts/
.venv/bin/python -m compileall -q src scripts tests
.venv/bin/python -m pip check
.venv/bin/python scripts/check_public_safety.py
git diff --check
```

视频资产库还必须通过：

```bash
pptx-tools videos doctor VIDEO_LIBRARY
pptx-tools videos doctor VIDEO_LIBRARY --verify-hashes
pptx-tools images doctor IMAGE_LIBRARY
pptx-tools images doctor IMAGE_LIBRARY --verify-hashes
```

快速检查验证结构、路径和引用；完整检查读取每个媒体实体并重算 SHA-256。
历史输出已被用户删除可作为信息项存在，不得误报为媒体或关联损坏。

### 5.1 本地发布前审计与压缩基准

发布前在目标平台运行依赖/产物审计入口：

```bash
python scripts/release_audit.py --check
```

该入口覆盖 Git 分支/提交溯源与干净工作树检查、`uv lock --check` /
`uv sync --locked --dry-run`、可选 `pip-audit`
（外部工具，缺失时降级为提示，不加入运行时依赖）、可选的 uv CycloneDX SBOM 与
`dist/` 产物版本/哈希。详细字段与边界见 `docs/RELEASE.md`。
公开二进制必须额外使用 `--public-binary --with-sbom ... --evidence ...`；缺少
pip-audit、SBOM、平台签名、公证（macOS）、恶意软件扫描、原生库清单或 FFmpeg
对应源码中的任一项都会失败关闭。普通候选构建只生成证据，不代表允许发布。

压缩基准入口用于对脱敏样本复现智能目标容量结果：

```bash
python scripts/run_compression_benchmark.py --manifest /path/to/manifest.json
python scripts/run_compression_benchmark.py --self-check
```

基准语料（样本文件与 manifest）**不进 Git**，仅以绝对路径引用；契约见
`docs/COMPRESSION_BENCHMARK.md`。

## 6. 发布门禁

- 构建前工作树变更范围清楚，测试与静态检查全部通过。
- macOS `.app` 的 `CFBundleIdentifier` 为反向域名，短版本与构建版本来自
  `pptx_tools.__version__`。
- 包内 `ffmpeg`/`ffprobe` 可执行，`codesign --verify --deep --strict` 通过。
- 包内 `licenses/` 包含项目、Python、Qt、字体/图标、全部运行时 Python 直接/传递
  依赖和所选 FFmpeg/LibreOffice 的匹配许可；缺一项不得发布。
- 记录 FFmpeg `-version/-buildconf`；包含 `--enable-gpl` 的二进制必须完成对应源码
  交付。当前公开门禁直接拒绝 Windows one-file，正式包使用可检查、可替换的 onedir。
- 正式候选只能使用 `scripts/build_ffmpeg_runtime.sh` 输出的 FFmpeg 8.1.2/x264/zlib
  固定源码构建，并把该平台对应源码包与安装包放在同一 Release；Homebrew/Gyan
  二进制不得进入正式公开包。
- 延迟加载的 DOCX/PDF/XLSX 后端与 `pikepdf` 必须能从 PyInstaller 归档中定位；
  产物完整性以模块、运行时、签名和启动检查为准，不以 DMG 大小猜测。
- DMG 只包含 `.app` 与 `Applications` 快捷方式，`hdiutil verify` 通过。
- 打包应用离屏启动后保持存活，并能正常退出，不产生新的崩溃报告。
- 本地 DMG 输出到用户指定的本机交付目录。普通提交不触发 CI，且 CI 与
  release workflow 都只允许手动触发，普通 push 和 tag 均不启动远程构建。
- 发布前运行 `scripts/release_audit.py --check`：工作树干净、锁文件一致、可选漏洞扫描、
  可选 SBOM 与产物哈希齐备；`pip-audit` 为外部工具，CycloneDX 由 uv 导出。
- 跨平台产物必须在对应平台构建；PyInstaller 不支持交叉编译，macOS 本地验证
  不能当作 Windows 可运行证据。
- macOS 候选产物默认 ad-hoc 签名；Developer ID 公证需单独配置 keychain profile。
  Windows 产物签名需在 Windows 主机用证书完成。ad-hoc、未公证或未签名产物只能
  留在 Draft/候选区，不能通过公开二进制门禁。
- Release workflow 只有 `contents: read`，只上传候选 workflow artifact，不接收 tag
  发布参数，也不包含 `gh release`。公开发布必须在独立人工复核中完成。
- 压缩基准语料不进 Git；基准结果目录默认 `benchmark-results/`，已加入
  `.gitignore`。

## 7. 完成定义

只有在代码、文档、全量测试、真实库只读体检、实际 GUI 截图检查、Git 提交推送、
本地 DMG 和产物校验都完成后，本轮目标才算完成。任何未验证边界必须在交付说明中
明确列出，不能用“应该可以”代替证据。
