# Doc Media Toolkit Architecture

本文是当前运行架构的权威说明。功能行为以代码和数据校验为准，视觉规范见
[`UI_DESIGN.md`](UI_DESIGN.md)，完成门禁见 [`QUALITY_GATES.md`](QUALITY_GATES.md)。

## 1. 系统边界

Doc Media Toolkit 是本地优先的 PySide6 桌面应用，同时提供统一 CLI。四个工作区
共享进程、资源、日志、设置和打包运行时，但各自保留独立业务模型：

1. 水印导出：DOCX / PDF / PPTX / 图片 / 视频预览、水印和导出。
2. 动态压缩：PPTX 内嵌媒体及独立图片/视频压缩、SSIM 审计和提档。
3. PPTX 视频资产库：视频身份、版本、高清源、PPTX 形状关联和回填。
4. 文档图片资产库：图片归档、来源、精确去重、相似审阅和元数据整理。

应用默认不依赖云服务。OpenAI 兼容接口是可选的整理建议能力，不参与数据一致性
判断，也不能自动归并或删除资产。

对外产品名固定为 `Doc Media Toolkit` / `文档媒体工具箱`，公开仓库和 Python 分发名为
`doc-media-toolkit`；导入包保留 `pptx_tools`，CLI 保留 `pptx-tools` 作为兼容标识。主壳、水印、压缩和帮助中心通过系统界面语言
或 `PPTX_TOOLS_LANG` 选择中英文（后者也接受 `en`）；视频库和图片库业务界面当前仍以中文为主，
因此发布说明不得宣称完整双语。

## 2. 入口与依赖方向

```text
pptx_tools_gui.py ──> pptx_tools.gui (统一壳)
                         ├─> pptx_output_watermark.gui ──> 水印核心
                         ├─> pptx_video_compactor_gui ───> 压缩核心
                         ├─> pptx_tools.video_manager_gui ─> video_manager
                         └─> pptx_tools.image_manager_gui ─> image_manager

pptx_tools_cli.py ──> pptx_tools.cli
                         ├─> pptx_output_watermark.cli
                         ├─> pptx_video_compactor.main
                         ├─> pptx_tools.video_manager.main
                         └─> pptx_tools.image_manager.main
```

依赖规则：

- `pptx_tools.gui` 负责导航、帮助、设置和生命周期，不承载领域写入逻辑。
- GUI 调用各自模型的公开方法；模型不得导入 GUI。
- 视频 GUI 与图片 GUI 不得互相导入。两者共用的 worker、样式位于
  `media_manager_ui.py`，字体、通用 QSS 和帮助能力位于 `ui_theme.py`。
- `video_manager.py` 可复用压缩核心与 FFmpeg 运行时来做指纹、转码和 PPTX 媒体
  重写，但资产身份和关联校验仍由视频模型负责。
- 跨工作区共享只发生在稳定能力上；不要为单一调用预建接口、工厂或插件层。

## 3. 运行时组件

| 层 | 主要模块 | 职责 |
| --- | --- | --- |
| 壳与设置 | `pptx_tools.gui` | 四页签、帮助、AI 设置、激活同步、退出协调 |
| 共享 UI | `ui_theme`, `media_manager_ui`, `language` | 字体、QSS、控件帮助、系统语言选择、资产库后台 worker |
| 水印 | `pptx_output_watermark/` | 文档转换、预览、PDF/PPTX/媒体水印 |
| 压缩 | `pptx_video_compactor*`, `pptx_quality_audit` | 媒体计划、流式提取、编码、SSIM、PPTX 重写 |
| 视频资产 | `video_manager`, `video_library_health` | 视频族/版本、PPTX 锚点、匹配、回填、体检 |
| 图片资产 | `image_manager` | 文档图片提取、哈希/像素身份、来源、体检 |
| 运行保障 | `project_lock`, `process_utils`, `app_logging` | 跨进程写锁、子进程收口、滚动日志 |
| 发布 | `scripts/build_standalone.py` | PyInstaller、运行时资源、签名修复、DMG |
| 验证辅助 | `scripts/release_audit.py`, `scripts/run_compression_benchmark.py` | 本地依赖/产物审计、manifest 驱动压缩基准 |

压缩不是图片/视频专业编辑器，而是文档容量优化的二级能力。PPTX、独立图片和独立
视频共用同一拖入队列、预算规则、质量门禁和少量通用设置；不建立独立媒体压缩接口。
DOCX/DOCM、PDF、XLSX/XLSM 的图片压缩后端已经通过格式专用结构与视觉门禁进入
正式 GUI/CLI；统一入口按后缀延迟加载后端，要求显式目标容量和启用图片预设。
格式安全合同与后续扩展边界记录在 `SMART_TARGET_COMPRESSION.md`。

压缩页的响应式布局沿用原生 Qt layout。可合并的评估控件按当前内容区可用宽度与
控件 `sizeHint` 动态判定，不使用固定窗口断点；空间不足时只把操作按钮移到标题行，
不引入第二套布局框架或改变压缩业务状态。统一壳复制子窗口样式后会保留 6 px
顶部安全内边距，让嵌入页首个面板的上边框不被标签页边缘遮住；该处理只影响视觉尺寸。

## 4. 数据所有权与不变量

### 4.1 视频库

```text
VIDEO_LIBRARY/
  video-project.json       # 权威清单
  video-project.json.bak   # 最近有效备份
  media/                   # 受管视频实体
  _cleanup/index.json      # 可恢复隔离索引
  reports/                 # 显式生成的报告
```

- `family.id` 表示内容身份，`variant.id` 表示具体文件版本。
- `source_variant_id` 和 `active_variant_id` 必须指向本族有效版本。
- SHA-256 是精确身份；内容指纹只在时长、宽高比、五帧、亮度和音频证据严格且
  唯一时辅助归族。
- `decks[].assets[]` 保存 PPTX 媒体 part 与 `(slide_path, shape_id)` 锚点；归并、
  移除或切换源时不得留下悬空引用。
- PPTX 与普通压缩媒体不是库内实体。库保存权威视频实体、路径/哈希别名和关联。

### 4.2 图片库

```text
IMAGE_LIBRARY/
  image-project.json       # 权威清单
  image-project.json.bak   # 最近有效备份
  images/                  # 受管图片实体
  _cleanup/index.json      # 可恢复隔离索引
```

- 文件 SHA-256 唯一标识精确重复；解码像素 SHA-256 和 dHash 只产生复用或相似
  候选，不把不同内容自动合并。
- 来源可来自独立文件、PPTX/DOCX 关系引用或 PDF 内嵌图像；原始文件不被修改。
- 当前没有图片回填能力，任何文档都不会因图片入库或整理被重写。

## 5. 写入、并发与恢复协议

### 5.1 清单提交

视频和图片清单写入都遵循：

1. 获取项目写锁；
2. 比较内存 revision 与磁盘 revision，拒绝旧窗口覆盖；
3. 构造并完整校验下一版清单；
4. 写临时文件并准备最近有效备份；
5. 使用 `os.replace` 原子替换；
6. 失败时保留最后有效清单和内存状态。

损坏的当前清单不得直接覆盖；只能由 `open()` 从有效备份恢复后再提交。

### 5.2 媒体隔离

隔离不是删除。图片和视频使用独立 cleanup 锁以及同一事务顺序：

1. 将包含 token、原路径、隔离路径、SHA-256 和快照的 `moving` 意图原子写入索引；
2. 把文件移动到 `_cleanup/`；
3. 将该项提交为 `quarantined`；
4. 之后才允许清单引用迁移/保存完成整个业务操作。

异常重启时，仅当隔离文件存在、原文件缺失且 SHA-256 符合预期，才把 `moving`
恢复为 `quarantined`；移动尚未发生则移除意图；其他组合保持为未完成问题并禁止
永久清空。新索引只保存项目相对路径，读取兼容旧绝对路径。视频原路径必须位于
`media/`，图片原路径必须位于 `images/`，隔离路径必须位于各自 `_cleanup/`。

还原先核验哈希和当前清单状态。若崩溃发生在“文件已隔离、清单尚未移除引用”之间，
还原会把文件放回原路径并清除索引，而不会重复添加清单记录。永久清空仅在索引、
引用、路径和文件状态全部安全时启用。

## 6. 文件与媒体处理

- PPTX/DOCX 使用 ZIP/XML 关系定位实际媒体，不根据文件名猜测引用。
- 大型 PPTX 媒体从 ZIP 成员流式复制到临时文件，避免按单个大视频/图片大小增加
  等量峰值内存。
- PPTX 输出写入临时 ZIP，校验结构、媒体关系、内容类型和必要锚点后原子替换。
- WMV/AVI 等不兼容媒体回填时生成 H.264/AAC MP4，并同步 relationship 与
  Content Type；不能只改扩展名。
- 高清回填按质量档位的上限语义交付：源已满足档位规格即原字节嵌入，否则按
  档位转码；执行路径与 GUI 预览共享同一 `_backfill_compatibility` 判定，
  避免决策漂移。
- 水印、压缩、审计和视频库共享 FFmpeg 定位与子进程登记。停止或退出时按进程组
  终止并等待，不使用 `QThread.terminate()`。

### 6.1 智能目标容量闭环

```text
关系图/显示面积/内容分类
          ↓
联合媒体预算 → 编码 → 双尺度质量审计 → 容器重打包 → 读取真实容量
                    ↑                         │
                    └── 最多两轮纠偏 / 一轮质量回补 ──┘
```

- 十进制目标容量对每个输入文件独立生效；规划和审计仍以质量底线为约束。目标容量
  模式在 SSIM 低于用户阈值时保留已编码候选并把素材标为 `below_threshold`，让用户
  看到真实的容量/质量取舍；强制模式触及绝对红线仍恢复原素材。普通安全模式和格式/
  结构审计失败也恢复原素材，不用体积变化代替质量证据。
- 视频基础阈值按预设默认为高 `0.95`、均衡 `0.93`、低体积 `0.90`，图片基础阈值为
  `0.99`；用户手动改动的视频阈值跨档位和会话保留。面积、复用次数和内容类型决定
  实际阈值。强制模式必须由安全版失败后的二次确认触发，并仍受视频 `0.90`、照片
  `0.96`、文字/线稿类 `0.98` 绝对红线约束。
- 目标模式默认 CPU x264 两遍；可选 GPU 路径逐素材探测并回退。Windows 优先级为
  NVENC、QSV、AMF、Media Foundation。音频只允许降低码率，不允许删除音轨。
- 目标容量纠偏最多两轮，并以媒体编码计划签名防止无效重复编码；GUI 进度按文件和
  全局值单调递增，不因重试回退。JSON 与 Markdown 报告记录真实目标差值、每轮容量、
  素材实际阈值、`below_threshold` 或恢复原因。
- FFprobe 明确证明为零时长且不超过一帧的音频流视为不可安全重编码：资产计划将其
  标为原媒体保留、恢复原始媒体路径并把原字节计入预算；不会因 FFmpeg 丢失该流而
  静音或删除音轨，报告使用 `unusable_audio_stream_preserved` 说明原因。
- 未主动降帧的 VFR 视频使用 FFmpeg `-fps_mode passthrough` 保留源时间戳；只有目标
  档位确实降帧时才插入 `fps=` 滤镜，并按目标帧率校验输出帧数，避免名义帧率导致
  静默丢帧。
- 视频库整理只自动隔离通过时长、音轨和内容一致性校验的版本；若整组只有被锁定的
  候选，界面明确说明不会执行操作。处理方式在界面中保持互斥；族内可由用户在“保留勾选版本，
  其他移入待清理”模式下单独勾选“人工确认：连锁定版本也移入待清理”，但必须保留已核实版本、
  经过二次确认且仍只进入可恢复隔离区，切换到其他处理方式会自动解除该复选项，跨族归并不支持强制。

## 7. AI 与其他信任边界

- Base URL 必须是有效 HTTP(S) 地址，不得包含用户名、密码、查询或片段；本地
  OpenAI 兼容服务可使用 HTTP。
- API Key 只保存在当前进程，不进入 QSettings、URL、日志或异常文本。
- HTTP 响应体限制为 5 MiB，错误体只读取有限前缀；重试等待可响应取消。
- 发送内容取决于视觉开关：关闭时只发送代码候选与规格，开启时才发送压缩预览。
- AI 返回值经过类型、长度、条目 ID 和置信度范围校验；结果只是建议，人工应用后
  仍通过模型公开写入路径和清单校验。
- 所有清单路径、ZIP 目标、分类路径、输出路径和 cleanup 索引都在各自边界再次
  校验，不能信任 GUI 或索引文件已经正确。

## 8. 长任务与生命周期

- GUI 的阻塞工作运行在 `QThread` worker 中；worker 只通过信号回传消息、结果和
  失败，不直接操作界面控件。
- 取消通过线程安全事件传入处理函数，并传播到 FFmpeg/外部进程链路。
- 页签首次创建后保持存活；切换页签不销毁窗口状态。重新激活资产库时只在没有
  正在运行的任务时 reload。
- 统一壳嵌入子工作区后会再次按 Qt 当前样式的 `sizeHint` 校正交互控件最小高度；
  这只约束视觉布局，不改变压缩、导出或资产关联协议。
- 退出时统一请求取消、终止已登记子进程、等待 worker/preview/audit 线程，然后
  释放 QtMultimedia 播放器。

## 9. 配置、日志与外部运行时

- `QSettings` 保存非秘密偏好和最近库路径；临时目录路径不会被持久化为最近库。
- 轮转日志位于平台应用数据目录的 `Doc Media Toolkit/logs/`，不写入资产库。
- FFmpeg/ffprobe 用于视频编码、探测、指纹和质量审计；发布包必须内置匹配许可，
  并按真实构建参数履行 LGPL/GPL 及对应源码义务。
- 正式候选由 `scripts/build_ffmpeg_runtime.sh` 从 SHA-256 固定的 FFmpeg 8.1.2、
  x264 与 zlib 源码构建；zlib.net 不可用时仅切换到同一上游版本的 GitHub 源，仍必须通过同一 SHA-256。
  只保留项目所需的 libx264、VideoToolbox/Media Foundation+D3D11VA
  和内置编解码/滤镜，并为每个平台生成对应源码、配置、工具链与哈希证据。
- 打包脚本把项目、Python、Qt、资源与运行时 Python 直接/传递依赖的许可材料统一
  放入 `licenses/`；缺少必需许可文本时失败关闭。具体分发边界见 `LICENSING.md`。
- Office/WPS/LibreOffice/Keynote/Pages 是文档导出的可选外部引擎；标准轻量包不
  内置它们。完整 LibreOffice 仅由显式 onedir 构建启用。
- Windows Office/WPS COM 转换在打开文档前强制禁用自动化宏；宏启用格式在引擎
  无法确认该设置时拒绝打开。LibreOffice 始终通过参数数组启动，不经过命令壳层。
- OOXML 关系目标先统一为 ZIP 的 POSIX 分隔符再解析；PDF 渲染和水印在分配位图或
  平铺内容前执行页面尺寸与像素预算检查。
- 发布构建固定使用 Python 3.12；开发/测试支持范围由 `pyproject.toml` 和 CI 矩阵
  共同定义。

## 10. 验证与发布

本地完整门禁：Ruff lint/format、compileall、pip check、公共安全扫描、隔离全量
测试、真实库只读快检/完整哈希复核、四工作区截图检查，以及实际 `.app`/DMG 的
内置工具、签名结构、离屏启动与 `hdiutil verify`。

`ci.yml` 与 `release.yml` 都只允许手动触发，避免普通 push 或 tag 消耗远程构建。
两者都从 `uv.lock` 导出带哈希的精确依赖，并在测试/打包前执行全仓 Ruff 格式检查。平台产物必须在对应
平台构建，不能把 macOS 本地验证当作 Windows 可运行证据。

`scripts/release_audit.py` 是不依赖远程 CI 的本地发布前审计入口，覆盖 Git 分支/提交
溯源与干净工作树、`uv lock --check`、可选 `pip-audit`（外部工具，从 `uv.lock`
导出当前目标平台全部 extra 的带哈希 requirements 后禁用再次解析）、可选 uv
CycloneDX SBOM 与 `dist/` 产物版本/哈希。`scripts/generate_native_inventory.py` 对
目标平台解包后的 onedir/`.app` 记录相对原生文件路径、SHA-256、架构和依赖工具输出；
`scripts/scan_release_artifact.py` 调用 ClamAV 或 Windows Defender 生成绑定产物哈希的
恶意软件报告；SBOM 和原生清单的证据描述同样必须绑定最终安装包 SHA-256，无扫描器时
失败关闭。三者都只生成证据，不发布、不修改核心算法。
`scripts/run_compression_benchmark.py` 以 manifest 驱动复现智能目标
容量压缩结果；两个入口都不修改核心算法。基准语料（样本与 manifest）不进 Git，
仅以绝对路径引用，契约见 `COMPRESSION_BENCHMARK.md`。跨平台基准必须在对应平台
运行，因为 GPU 探测、FFmpeg 路径和编码器可用性都是平台相关的。

macOS 产物默认 ad-hoc 签名；Developer ID 公证需单独配置 keychain profile，本地
审计不执行公证。Windows 产物签名需在 Windows 主机用证书完成，本地审计只记录哈希，
不签名。

## 11. 扩展规则

新增功能时按以下顺序放置：

1. 先确定数据所有者和不可破坏的不变量；
2. 将业务规则放入对应核心/模型，GUI 只做输入、展示和显式确认；
3. 优先复用现有核心、标准库或已安装依赖；只有两个以上稳定调用方需要时才抽共享层；
4. 新的资产库 GUI 共性进入 `media_manager_ui.py`/`ui_theme.py`，不得从另一个业务
   GUI 间接导入；
5. 新的清单实体必须加入 schema、下一版校验、备份恢复、doctor 和迁移测试；
6. 新的删除/覆盖流程必须先提供可恢复或原子提交路径；
7. 新的外部协议必须限制协议、大小、超时、秘密和取消；
8. 同步 README、HANDOFF、UI、QUALITY、RELEASE 与本文，再通过完整门禁。新增
   本地审计/基准入口时同样同步 `COMPRESSION_BENCHMARK.md`。

## 12. 受控技术债务

- 视频模型和三个成熟处理 GUI 仍是较大的工作流模块。它们的复杂度来自已有业务
  状态与跨平台回退；在没有稳定边界和等价回归证据前，不按行数拆分。
- 图片 CLI 目前以只读 doctor 为主；写操作已有 GUI 和模型 API。只有出现可重复的
  自动化需求时再扩充 CLI，避免形成第二套校验逻辑。
- 图片库尚无专用批准主界面截图；当前以共享视觉令牌、本文件、`UI_DESIGN.md` 和
  实际 Qt 截图验收为准。
