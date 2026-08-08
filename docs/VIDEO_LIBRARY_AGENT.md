# 视频库外部 AI Agent 操作规范

本文可直接交给具备本机文件和终端权限的 AI Agent。目标是让 Agent
协助核查、命名、归类、导入和回填，同时不破坏视频实体、PPTX 关联或清理恢复能力。

## 当前环境

```bash
REPO="${REPO:?Set REPO to the local pptx-tools checkout}"
LIB="${PPTX_VIDEO_LIBRARY:?Set PPTX_VIDEO_LIBRARY to the target video library}"
PY="$REPO/.venv/bin/python"
export PYTHONPATH="$REPO/src"
```

在启动 Agent 前由用户或受信任的本机环境设置这两个变量。Windows 或其他机器
应使用对应的本机路径，不得把任何用户目录硬编码进代码或提交。

## 如何交给外部 Agent

- 本机 Agent：把本文末尾提示词直接发给它；它通过上述仓库和视频库路径工作。
- 代码 Agent：可读取本地仓库，或从
  `https://github.com/roanpy/doc-media-toolkit` 获取代码，再以本机仓库为执行目录。
- 纯在线聊天 AI：没有本机文件权限，最多分析脱敏后的报告，不能安全执行入库、
  重命名、归并、回填或验证。
- 不要向不可信服务上传原始视频、PPTX、`video-project.json` 或应用日志；
  这些内容可能包含客户、项目和绝对路径信息。

## 权威信息

- `$LIB/video-project.json`：唯一权威索引，记录视频族、版本、高清源、已知哈希、
  内容指纹、PPTX 哈希/路径别名及幻灯片/形状锚点。
- `$LIB/video-project.json.bak`：最近一次有效清单备份。
- `$LIB/media/`：视频实体。文件名和目录不是身份标识。
- `$LIB/_cleanup/`：隔离区。`index.json` 使用可恢复的移动意图；未完成引用、
  哈希、路径和状态检查前不得清空。
- `$LIB/reports/`：Agent 会话报告和批处理报告。
- macOS 日志：
  `~/Library/Application Support/Doc Media Toolkit/logs/app.log`
- Windows 日志：
  `%LOCALAPPDATA%\Doc Media Toolkit\logs\app.log`

日志会轮转，不能代替本次 Agent 的交付报告。`video-project.json` 只记录当前状态，
不保存完整操作历史。GUI 操作写入应用日志；直接调用模块 CLI 时，Agent 还必须
保存 stdout/stderr。

## 代码入口

- 视频库模型：`src/pptx_tools/video_manager.py`
- 视频库健康检查：`src/pptx_tools/video_library_health.py`
- 视频库界面：`src/pptx_tools/video_manager_gui.py`
- PPTX 媒体压缩：`src/pptx_video_compactor.py`
- 压缩界面及压缩后关联：`src/pptx_video_compactor_gui.py`
- 真实媒体端到端测试：`tests/test_video_manager_e2e.py`
- 视频库单元测试：`tests/test_video_manager.py`
- 已验证回归测试：`tests/test_verified_regressions.py`

修改业务数据时优先使用 CLI 或 `VideoProject` 的公开方法。不要复制内部实现另写脚本。

## 不可违反的规则

1. 不得直接编辑 `video-project.json` 或 `video-project.json.bak`。
2. 不得用 Finder、`mv`、`rm` 或批量重命名器改动库内视频。
3. 不得调用 `import_variant(..., verify_identity=False)`。
4. 不得仅凭文件名、时长、分辨率或肉眼相似认定同一视频。
5. 多个指纹候选、不同音轨、裁剪或时长变化必须保持未归并。
6. 新版本只作为候选入库；未经用户确认不得设为高清源。
7. 不得直接删除重复视频。只能经 GUI“整理视频库”移入 `_cleanup/`。
8. 不得清空 `_cleanup/`，除非 `cleanup_pending_issues()` 返回空且用户明确确认。
9. 高清回填必须输出到新路径。未经用户明确确认不得覆盖源 PPTX。
10. 任何写操作前记录清单 revision，写后重新打开视频库并验证。
11. 不得手工改写 `_cleanup/index.json`。新条目使用库内相对路径；原路径必须在
    `media/`，隔离路径必须在 `_cleanup/`，重启恢复必须通过 SHA-256 核验。

## 身份与关联模型

- `family.id`：视频族稳定身份。
- `variant.id`：具体视频版本稳定身份。
- `source_variant_id`：当前高清回填使用的权威版本。
- `variant.sha256`：实体文件精确身份。
- `known_hashes`：同族已知实体和压缩版本哈希别名。
- `content_fingerprint`：时长、宽高比、五帧画面和音频频谱的保守匹配依据。
- `decks[].assets[]`：PPTX 媒体与视频族的关联，包含媒体 part、
  幻灯片路径和 shape ID 锚点。
- `source_aliases`：PPTX 改名或移动后的已知路径。

因此，PPTX 或视频改名不会天然丢失身份。外部重编码造成 SHA-256 改变时，
只有内容指纹唯一匹配才会自动归入已有视频族；否则必须人工确认。

## 使用场景与人工边界

| 场景 | 程序自动完成 | 需要人工确认 |
|---|---|---|
| 别人发来的新 PPTX | 按精确哈希或唯一内容指纹复用视频族；否则新建视频族并记录形状锚点 | 更高质量候选是否设为高清源 |
| 工具生成的压缩版再次压缩 | 识别已登记输出，沿用同一 PPTX 记录；登记每轮新视频哈希 | 无安全歧义时不需要 |
| 工具外压缩或改名的 PPTX | 文件改名不影响；未知压缩哈希可按唯一内容指纹匹配并固化别名 | 多候选、不同音轨、裁剪或时长变化 |
| 手工导入外部视频 | 返回 existing / matched / created / ambiguous；安全匹配后纳入版本管理 | matched 候选是否升级为高清源，ambiguous 选择哪个族 |
| 视频文件改名或移动 | 通过 GUI 操作同步清单；库外移动后可按 SHA-256 重连 | 文件内容同时改变时按新版本导入 |
| 高清回填 PPTX | 按视频族替换媒体 part，保留幻灯片、形状和播放动画关系，默认另存 | 覆盖原文件及最终 PowerPoint 播放验收 |
| 重复清理 | 结合哈希、内容指纹、SSIM、音轨、时长和分辨率给出建议并先隔离 | 归并、切换高清源和永久清空 |
| 语义命名与项目分类 | 保存人工或 Agent 通过公共方法提交的名称和目录 | 仍需人或可访问素材的 AI 理解业务语义 |

程序能够独立维护视频实体、版本、PPTX 锚点、压缩哈希别名和回填关系；
AI 不是关联正确性的前提。AI 主要用于业务语义命名、分类和复杂歧义辅助判断，
不得绕过程序的身份校验或用户确认。

## 标准工作流

### 1. 只读核查

先执行：

```bash
cd "$REPO"
git status --short
"$PY" -m pptx_tools.video_manager list "$LIB" > /tmp/video-library-list.json
```

再用只读代码统计状态：

```bash
"$PY" - <<'PY'
import os
from pathlib import Path
from pptx_tools.video_manager import VideoProject

root = Path(os.environ["PPTX_VIDEO_LIBRARY"])
project = VideoProject.open(root)
statuses = {"available": 0, "missing": 0, "modified": 0}
for family in project.families():
    for variant in family["variants"]:
        statuses[project.status(variant)] += 1
print({
    "revision": project.data["revision"],
    "families": len(project.families()),
    "variants": sum(len(item["variants"]) for item in project.families()),
    "decks": len(project.decks()),
    "statuses": statuses,
    "pending_cleanup": len(project.pending_cleanup()),
})
PY
```

任何 `missing`、`modified`、清单恢复警告或 ZIP/PPTX 错误都必须先报告，不能继续清理。

也可直接使用与 GUI 相同的只读检查：

```bash
"$PY" -m pptx_tools.video_manager doctor "$LIB" \
  --report "$LIB/reports/agent-library-health.json"
```

需要证明实体内容未变时增加 `--verify-hashes`。该模式会读取全部视频，耗时更长；
它能区分时间戳变化与真实 SHA-256 不一致。不得把
`--prune-stale-outputs` 与其他写操作混在同一步执行。

### 2. 从 PPTX 归档高清视频并建立关联

```bash
"$PY" -m pptx_tools.video_manager add \
  "$LIB" \
  "/absolute/path/example.pptx" \
  --source-quality 1080p \
  --category "客户/项目"
```

需要逐字节保留 2K/4K 或旧格式原片时才使用：

```bash
--source-quality original
```

归档不会修改输入 PPTX。完全相同视频只保存一份，各 PPTX 仍分别保留关联。
`videos add` 必须走 `archive_and_register_pptx`，会在同一用户工作流中同时归档
媒体并登记 PPTX/形状关联。Agent 不得用低层 `archive_pptx_videos` 代替该命令。

### 3. 导入外部视频

让程序先自动匹配：

```bash
"$PY" -m pptx_tools.video_manager import-video \
  "$LIB" \
  "/absolute/path/video.mp4" \
  --source-quality original \
  --category "客户/项目"
```

结果处理：

- `existing`：实体或哈希已存在，不重复保存。
- `matched`：唯一内容指纹命中，新增为已有族候选版本。
- `created`：没有安全候选，创建新视频族。
- `ambiguous`：多个候选，不写入；列出候选并让用户选择。
- `failed`：不修改库，报告错误。

用户确认目标族后才允许：

```bash
"$PY" -m pptx_tools.video_manager import-video \
  "$LIB" \
  "/absolute/path/video.mp4" \
  --family-id "用户确认的-family-id" \
  --source-quality original
```

即使指定 `family-id`，程序仍会校验内容身份，不安全时会拒绝。

### 4. 设为高清源

先向用户展示候选版本的分辨率、时长、码率、大小、音轨、路径和预览结论。
用户确认后执行：

```bash
"$PY" -m pptx_tools.video_manager set-source "$LIB" "variant-id"
```

不要仅因分辨率更高就自动执行。

### 5. PPTX 高清回填

桌面程序优先使用 `高清回填 PPTX（另存）`。它会先显示所有内嵌视频，而不是
直接执行：

- `精确匹配`：当前媒体 SHA-256 已属于某个视频族。
- `内容匹配`：时长、宽高比、五帧画面和音频唯一匹配。
- `待确认`：只给出近似候选，默认保持当前，不自动建立关联。
- `保持当前`：该媒体不变，不新增哈希关系。
- `仅本次回填`：使用人工选择的视频族，但不声称二者是同一视频。
- `确认同一视频并记住`：输出校验成功后才把当前媒体哈希加入所选视频族。

选中一项后可查看当前视频和目标源的三帧封面对比、播放完整视频，并搜索全部
视频族。取消整份清单不会修改 PPTX 或视频库。

```bash
"$PY" -m pptx_tools.video_manager upgrade \
  "$LIB" \
  "/absolute/path/input.pptx" \
  --output "/absolute/path/output_high_quality.pptx"
```

只迁移 WMV/AVI 等不兼容媒体：

```bash
--incompatible-only
```

CLI 默认采用安全自动结果，不提供交互式清单。需要 Agent 半自动处理时，必须先
调用公共只读接口：

```python
items = project.review_pptx_matches(
    input_pptx,
    review_directory,
    include_resolved=True,
)
```

让用户核对后，再把决策分别传给
`upgrade_pptx_from_library(family_overrides=..., remember_manual_matches=...,
keep_current_media=...)`。不确定是否同一内容时只能放入 `family_overrides`，
不得同时放入 `remember_manual_matches`；用户选择保持的媒体必须放入
`keep_current_media`。审阅目录是临时副本，任务结束后可清理。

回填后必须检查：

- 输出可作为 ZIP 打开。
- 幻灯片和 shape ID 锚点未变化。
- 未匹配和歧义视频保持原样。
- 视频分辨率、时长、音轨和帧数有效。
- PowerPoint 能打开且不提示修复。

### 6. 文件被手工改名或移动

不要手工修 JSON。执行：

```bash
"$PY" -m pptx_tools.video_manager relink \
  "$LIB" \
  "$LIB/media"
```

程序按 SHA-256 重新关联。内容已经被修改而哈希变化时，不得强制重连；
应按“导入外部视频”流程处理。

### 7. 命名、移动、归并和清理

优先使用应用 GUI：

- `重命名`：同步更新清单和高清源文件。
- `移动文件`：通过 `VideoProject.move_variant` 更新路径。
- `归并视频`：保留双方已知哈希和 PPTX 引用。
- `整理视频库`：计算 SSIM、时长、音轨和分辨率后给出候选。
- `待清理`：先隔离、可还原，最后才清空。

AI 可以分析抽帧、字幕、音频和来源 PPTX 来建议名称、分类及保留版本，
但最终归并、高清源切换和永久清空必须由用户确认。

## Agent 会话报告

每次操作在 `$LIB/reports/` 新建：

```text
agent-session-YYYYMMDD-HHMMSS.md
```

同时保存命令原始输出：

```bash
SESSION="$LIB/reports/agent-session-$(date +%Y%m%d-%H%M%S)"
mkdir -p "$LIB/reports"
"$PY" -m pptx_tools.video_manager list "$LIB" \
  2>&1 | tee "$SESSION.log"
```

后续写操作继续用 `2>&1 | tee -a "$SESSION.log"` 追加，不得只保留聊天摘要。

报告至少包含：

```markdown
# 视频库 Agent 会话

- 开始时间：
- 仓库提交：
- 视频库 revision（前）：
- 视频库 revision（后）：
- 操作范围：

## 只读核查

- 视频族：
- 视频版本：
- PPTX 记录：
- missing / modified：
- 待清理：

## 建议

| 对象 ID | 当前名称 | 建议名称/分类 | 证据 | 置信度 | 是否需用户确认 |
|---|---|---|---|---|---|

## 已执行

| 命令/公共方法 | 对象 ID | 结果 | 输出路径 |
|---|---|---|---|

## 未执行

- 歧义候选：
- 不同音轨/裁剪：
- 需要用户确认：

## 验证

- 清单可重新打开：
- missing / modified：
- PPTX ZIP：
- PowerPoint 人工打开：
- 测试结果：
```

会话报告可能包含私有文件名和路径，只保存在视频库本地，不提交到公开 GitHub。

## 代码修改要求

如果任务是修改程序而不是整理业务数据：

1. 先读 `README.md`、`docs/HANDOFF.md`、`docs/ARCHITECTURE.md`、
   `docs/RELEASE.md` 和本文。
2. 检查 `git status`，不得覆盖用户未提交修改。
3. 复用 `VideoProject` 现有方法，不创建第二套匹配或清单写入逻辑。
4. 不降低哈希、内容指纹、音频和歧义保护。
5. 至少运行：

```bash
QT_QPA_PLATFORM=offscreen "$PY" scripts/run_tests_isolated.py
"$PY" -m ruff check src tests scripts
"$PY" -m ruff format --check src tests scripts
git diff --check
```

6. GitHub Actions 只允许手动触发；普通提交不启动构建。
7. 报告修改内容、测试结果、提交号和产物路径。

## 可直接复制给外部 Agent 的提示词

```text
你在本机协助维护 Doc Media Toolkit 视频库。

仓库：
<REPO：本机 pptx-tools 仓库绝对路径>

视频库：
<PPTX_VIDEO_LIBRARY：本机目标视频库绝对路径>

开始前必须阅读：
- README.md
- docs/HANDOFF.md
- docs/RELEASE.md
- docs/VIDEO_LIBRARY_AGENT.md

严格遵守 docs/VIDEO_LIBRARY_AGENT.md：
- video-project.json 只读，不得直接编辑；
- 不得手工移动、重命名或删除库内视频；
- 所有修改必须调用现有 CLI、GUI 或 VideoProject 公共方法；
- 先做只读核查和建议清单；
- 歧义、归并、设为高清源、永久清理和覆盖 PPTX 必须先由我确认；
- 文件名、时长或分辨率不能单独证明同源；
- 每次操作在视频库 reports/ 写 agent-session 报告；
- 修改代码必须跑全套测试，不要未经明确要求触发 GitHub 构建。

先只读核查当前 revision、视频族/版本/PPTX 数量、missing、modified、
待清理项和目标文件，再给出分步执行计划。未经确认不要执行高影响操作。
```
