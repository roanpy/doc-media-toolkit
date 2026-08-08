# 本地压缩基准（Manifest 驱动）

状态：入口已实现，复用现有智能目标容量核心；不引入新运行时依赖。
适用入口：`scripts/run_compression_benchmark.py`

本文记录基准入口的契约、manifest 格式、输出字段和安全边界。压缩算法本身的
规格见 [`SMART_TARGET_COMPRESSION.md`](SMART_TARGET_COMPRESSION.md)，架构定位见
[`ARCHITECTURE.md`](ARCHITECTURE.md)，发布门禁见 [`QUALITY_GATES.md`](QUALITY_GATES.md)。

## 1. 目标与边界

提供一个本地、可复现的入口，对脱敏样本运行现有压缩核心，并聚合输出机器可读 JSON
与人类可读 Markdown。基准**只调用** `pptx_video_compactor.compact_input_path`，
不修改压缩算法、不引入第二套编码路径、不为了虚构样本改变核心行为。

- 样本文件由用户在仓库外维护（例如 `~/benchmarks/pptx-samples/`），仅以绝对路径
  出现在 manifest 中；**样本与 manifest 都不进入 Git**（见第 5 节）。
- 基准在用户指定的输出目录生成压缩产物和报告，默认 `./benchmark-results/`，该目录
  已加入 `.gitignore`。
- 不接入 GitHub Actions；基准是本机或构建主机上的手动入口。

## 2. Manifest 格式

manifest 是一个 JSON 对象，至少包含 `samples` 数组。每个样本是一个对象：

```json
{
  "notes": "脱敏样本，2026-08-08，macOS arm64",
  "samples": [
    {
      "path": "/absolute/path/to/benchmarks/samples/deck-a.pptx",
      "label": "deck-a",
      "target_size_mb": 8.0,
      "profile": "high",
      "image_profile": "high",
      "encoder": "auto",
      "quality_mode": "safe",
      "video_ssim_threshold": 0.95,
      "image_ssim_threshold": 0.99,
      "preset": "medium"
    }
  ]
}
```

字段说明：

| 字段 | 必填 | 说明 |
| --- | --- | --- |
| `path` | 是 | 样本绝对路径。基准会 `expanduser().resolve()` 并校验存在。 |
| `target_size_mb` | 否 | 十进制目标容量（1 MB = 1,000,000 bytes）。省略时走预设压缩而非目标模式。 |
| `label` | 否 | 报告中显示的名称；默认用文件名。 |
| `profile` | 否 | 视频预设，默认 `high`。与 CLI 一致。 |
| `image_profile` | 否 | 图片预设，默认 `high`。 |
| `encoder` | 否 | `auto`/`cpu`/`gpu`，默认 `auto`。 |
| `quality_mode` | 否 | `safe`/`forced`，默认 `safe`。 |
| `video_ssim_threshold` | 否 | 视频基础 SSIM 阈值，默认 `0.95`。 |
| `image_ssim_threshold` | 否 | 图片基础 SSIM 阈值，默认 `0.99`。 |
| `preset` | 否 | ffmpeg x264 预设，默认 `medium`。 |
| `reserve_mb` | 否 | 显式预留（十进制 MB）；默认动态。 |
| `meta_*` | 否 | 仅保留简单值作为 `manifest_extra` 溯源；其他未知字段忽略，避免把敏感值写入报告。 |

manifest 中不得包含密码、Token 或真实客户信息。基准只记录样本路径的 SHA-256，
不复制样本内容。

## 3. 运行

```bash
# 正式基准
python scripts/run_compression_benchmark.py --manifest /path/to/manifest.json

# 指定输出目录
python scripts/run_compression_benchmark.py --manifest ... --output-dir ./bench-out

# 无样本自检（验证导入链并生成空骨架）
python scripts/run_compression_benchmark.py --self-check
```

基准对每个样本：

1. 构造与 `pptx-tools compact` CLI 相同的 `argparse.Namespace`；
2. 调用 `compact_input_path`，计时包围整个调用；
3. 读取核心返回的 `report_path` 指向的 `.report.json`，提取目标容量、纠偏轮数、
   质量/结构结果、GPU/回退信号；
4. 把每个样本的结果聚合进 `benchmark.json` 与 `benchmark.md`。

## 4. 输出字段

`benchmark.json` 顶层：

- `schema` / `version`：报告模式标识。
- `generated_at` / `host`：生成时间与主机信息（平台、Python、架构）。
- `samples[]`：逐样本结果。
- `aggregate`：跨样本汇总。

每个 `samples[]` 条目至少包含：

| 字段 | 含义 |
| --- | --- |
| `label` | 样本标签或文件名。 |
| `input_path` / `input_sha256` | 输入文件名与其 SHA-256；不写入绝对路径。 |
| `output_path` / `output_sha256` | 输出文件名与其 SHA-256（失败时为 `null`）。 |
| `target_size_mb` / `target_bytes` | 目标容量。 |
| `actual_bytes` | 实测输出容量。 |
| `delta_bytes` | 目标容量误差（实际 − 目标）。 |
| `target_ratio` | 实际/目标比率。 |
| `target_status` | `met` / `quality_limited` / `source_already_meets` / `not_requested` / `not_measured`。 |
| `correction_rounds` | 容量纠偏轮数（来自 `target_capacity_attempts` 中 `kind=correction`）。 |
| `quality_giveback_rounds` | 质量回补轮数。 |
| `elapsed_sec` | 单样本耗时。 |
| `encoder_requested` | 请求的编码器模式。 |
| `assets` | 素材级汇总：数量、状态分布、`gpu_used`、`restored_original_assets`、`below_threshold_assets`。 |
| `skipped` / `skip_reason` | 是否跳过及原因。 |
| `report_path` | 核心生成的单样本报告路径。 |
| `error` | 错误信息（正常为 `null`）。 |

`aggregate` 包含样本数、完成/跳过/错误数、是否使用 GPU、纠偏轮数合计、恢复原件与
`below_threshold` 素材数、目标比率均值/最小/最大、容量差值均值、耗时合计/均值。

`benchmark.md` 提供汇总段落与逐样本表格，便于人工快速浏览。

## 5. 安全边界与不进 Git

- 样本文件和 manifest 必须放在仓库外。`.gitignore` 已忽略 `benchmark-results/`
  目录与常见媒体扩展名（`*.pptx`/`*.pdf`/`*.mp4` 等）。
- 基准不修改源样本，输出写入 `--output-dir` 指定目录。
- 基准不访问网络、不调用 AI 分类、不修改视频/图片资产库。
- 跨平台基准必须在对应平台运行：压缩核心的 GPU 探测、FFmpeg 路径和编码器可用性
  都是平台相关的，不能把一台机器的基准结果当作另一平台的证据。

## 6. 与核心的关系

基准是核心的**只读消费者**。它复用 `compact_input_path` 的返回值和核心已经写入的
`.report.json`，不维护第二套报告逻辑。如果核心报告字段发生变化，基准的字段提取
需要同步更新；这种耦合是刻意的，避免基准与核心行为漂移。
