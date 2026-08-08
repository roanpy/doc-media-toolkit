# PPTX 高清回填质量档位（Backfill Quality Tiers）

日期：2026-08-05
状态：已批准（用户确认「上限语义 / 只记别名 / 对话框内选项 / 含均衡档」）

## 背景与目标

当前 `upgrade_pptx_from_library` 回填时，`_delivery_master` 对「兼容」的族高清源
（mp4/m4v + h264 + aac/无音轨 + ≤1080p 包络 + 偶数尺寸）直接**原样嵌入**。当源是
高码率 1080p（如 50 Mbps）时，回填后的 PPTX 体积巨大，用户不得不再跑一次压缩。

目标：在回填流程中加入**质量档位**选择，使回填输出可直接使用，无需二次压缩。

- 默认档 = 「最佳」，行为与现状**完全一致**（零回归）。
- 其余档位按档位规格（分辨率上限 + CRF + 码率上限 + 音频码率）输出。
- 删除/覆盖红线不变：回填永远另存新文件，不覆盖输入。

## 档位定义（上限语义）

| 档位 | 分辨率上限（横/竖） | CRF | 视频码率上限 | 音频码率 | 输出名后缀 |
|---|---|---|---|---|---|
| 最佳（默认） | 1920×1080 / 1080×1920 | 18 | 不设限 | 256k | `_high_quality` |
| 高质量 | 1920×1080 / 1080×1920 | 20 | 12 Mbps | 256k | `_hq1080p` |
| 均衡 | 1280×720 / 720×1280 | 23 | 5 Mbps | 128k | `_balanced720p` |

**上限语义**：族源同时满足以下全部条件 → 原样嵌入（零损失，现状逻辑的扩展）；
任一不满足 → 按档位规格转码。

1. 容器/编码兼容：`.mp4`/`.m4v` + h264 + aac/无音轨；
2. 分辨率 ≤ 档位上限（沿用 `_video_envelope` 横竖包络）且宽高为偶数；
3. **码率 ≤ 档位码率上限**（新增维度；「最佳」档不检查）。

码率上限是必要维度：仅按分辨率判定会让 1080p/50Mbps 的源被原样嵌入，达不到
「免去二次压缩」的目标。码率取自 `probe_video` 已有的 `bitrate_kbps`
（format bit_rate，回退 video bit_rate）；探测不到码率时视为「超限」→ 转码
（安全方向，宁可多转一次不可漏压）。

## 别名学习（不变 + 扩展）

- 转码产生的档位母版**只记别名、不注册版本**（不入 `variants`）：其 sha256 追加到
  族 `known_hashes` 并随 `self.save()` 持久化（沿用现有 `learned_aliases` 机制）。
- 效果：用回填产物（如均衡档压缩版）再次回填时可通过哈希匹配回族，但族结构不被
  派生物污染；临时母版随 `work` 目录在 `finally` 中删除。
- 「canonical-content」语义不变：内嵌视频已与某档母版同哈希时走
  `already_high_quality` 跳过（现有逻辑，按 master_digest 判定，天然档位感知）。

## 核心改动（`src/pptx_tools/video_manager.py`）

### 1. 档位规格表

模块级常量（KISS，dict 而非新类）：

```python
BACKFILL_QUALITY_TIERS: dict[str, dict[str, Any]] = {
    "best":     {"label": "最佳",   "max_width": 1920, "max_height": 1080,
                 "crf": 18, "bitrate_kbps": 0, "audio_bitrate": "256k",
                 "suffix": "high_quality"},
    "high":     {"label": "高质量", "max_width": 1920, "max_height": 1080,
                 "crf": 20, "bitrate_kbps": 12000, "audio_bitrate": "256k",
                 "suffix": "hq1080p"},
    "balanced": {"label": "均衡",   "max_width": 1280, "max_height": 720,
                 "crf": 23, "bitrate_kbps": 5000, "audio_bitrate": "128k",
                 "suffix": "balanced720p"},
}
DEFAULT_BACKFILL_TIER = "best"
```

`bitrate_kbps = 0` 表示不设限。竖版上限由 `_video_envelope` 等价物按档位
max_width/max_height 推出（横 1280×720 ↔ 竖 720×1280）。

### 2. `_transcode_high_quality_mp4` 泛化

新增关键字参数：`crf=18`、`bitrate_kbps=0`、`audio_bitrate="256k"`、
`max_width=1920`、`max_height=1080`（默认值 = 现状）。码率上限 >0 时在命令中追加
`-maxrate {kbps}k -bufsize {2×kbps}k`。6 处归档/归一化调用方（1634/1690/1707/
1901→改造/2943/2955/4632）不传新参，行为不变。

同时移植 `pptx_video_compactor.py` 已有的两项加固（本地小助手，不跨模块依赖）：

- **VFR 帧保护**：回填转码不降低帧率，无条件加 `-fps_mode passthrough`
  （对应 `append_frame_rate_mode` 在非降帧分支的行为）。要求 ffmpeg ≥ 5.1。
- **空音轨保护**：探测源音轨，若 ffprobe 明确报告 duration ≤0 且 frames ≤1
  （`audio_stream_is_usable` 语义）→ 命令中不 map 音频、不加 `-c:a/-b:a`；
  探测不确定时保持现状（map + aac）。

### 3. `_delivery_master(family, work, tier="best")`

- 按档位取上限：`_video_envelope` 泛化为可传 max_width/max_height（默认 1920/1080）。
- 达标判定新增：`tier.bitrate_kbps == 0 or metadata_bitrate <= tier.bitrate_kbps`，
  其中 `metadata_bitrate = variant.get("bitrate_kbps")`（缺失时 probe 补）。
- 转码调用传入档位全部参数。
- 返回值签名不变 `(Path, sha256)`。

### 4. `upgrade_pptx_from_library(..., quality_tier="best")`

- 新关键字参数，默认 `"best"`；非法值抛 `ValueError`。
- `masters` 缓存键由 `family["id"]` 改为 `(family["id"], tier)`（一次运行一个
  档位，键改动仅为防御性正确）。
- 默认输出名：`f"{stem}_{tier.suffix}.pptx"`（best → `_high_quality`，与现状同名）。
- 返回 dict 新增 `"quality_tier": tier`。

### 5. 执行计划预览（供 GUI 列显示）

新增 `plan_backfill_action(metadata, tier) -> str`（模块级纯函数）：复用共享判定
（只读，不落盘），返回 `"原样嵌入（已达标）"` / `f"转码至 ≤{max_height}p · CRF {crf}[ · ≤NMbps]"`
/（尺寸未知时）`"按档位规格回填"`。判定逻辑抽取为共享私有函数
`_backfill_compatibility(metadata, spec) -> bool`，`_delivery_master` 与
`plan_backfill_action` 共用，杜绝两处判定漂移。

## GUI 改动（`src/pptx_tools/video_manager_gui.py`）

`BackfillConfirmDialog`：

1. 顶部（summary 下方）加档位选择行：`QLabel("回填质量") + QComboBox`
   （最佳/高质量/均衡，data=tier key），默认「最佳」。
2. 选择持久化：`QSettings` 键 `backfill/quality_tier`，构造时恢复、切换时写入
   （沿用 CleanupDialog 的 QSettings 模式）；持久化值非法时回退 best。
3. 切换档位时重算「执行计划」列（调 `plan_backfill_action`，纯探测不转码）。
4. 对话框新增 `quality_tier()` 访问器；调用方（回填入口）把
   `dialog.quality_tier()` 传入 `upgrade_pptx_from_library`。
5. 标题/intro 文案补充当前档位说明。

样式：复用 `MEDIA_MANAGER_STYLESHEET` 现有控件样式，不新增 objectName。

## 测试

**单元**（`tests/test_video_manager.py`）：

- 档位表完整性：三档字段齐全、best = 现状常量（1920/1080/18/0/256k）。
- `_delivery_master`：
  - 默认档行为回归 = 现状（兼容源原样返回，不兼容转码）。
  - 均衡档：1080p 兼容源 → 转码；720p 低码率兼容源 → 原样。
  - 码率触发：1080p 低码率源在高质量档原样；1080p 20Mbps 源在高质量档转码
    （分辨率达标但码率超限）。
  - 码率缺失（probe 无 bit_rate）→ 非 best 档一律转码。
  - 竖版：720×1280 源在均衡档原样；1080×1920 源在均衡档转码。
- `upgrade_pptx_from_library`：quality_tier 默认值 = best；非法值 ValueError；
  输出名后缀随档位；返回 dict 含 quality_tier；别名学习仍发生（转码母版哈希入
  known_hashes）。
- `plan_backfill_action` 与 `_delivery_master` 判定一致性（同输入同结论）。
- `_transcode_high_quality_mp4`：默认参数命令与现状逐 token 一致；带
  bitrate_kbps 时命令含 `-maxrate/-bufsize`；空音轨源不含 `-map 0:a?`；
  命令含 `-fps_mode passthrough`。

**GUI**（`tests/test_verified_regressions.py` DesktopLifecycleTest 模式）：

- 对话框默认选中「最佳」；QSettings 持久化往返；非法持久化值回退 best。
- 切换档位后执行计划列文本更新。
- `quality_tier()` 访问器返回值正确。

**E2E**（真实 ffmpeg，沿用现有 E2E 模式）：

- 合成 1280×720 高码率短视频建族 → 均衡档回填 → 输出内嵌视频 ≤720p 且
  h264/aac 128k；默认档回填 → 与现状一致。

## 文档

README.md（功能段+档位表）、docs/HANDOFF.md、docs/ARCHITECTURE.md（回填小节）、
CHANGELOG.md 同步。

## 非目标（明确不做）

- 不扩展「内嵌已达标即跳过」的判定（保持 canonical-content 语义，由用户用
  keep_radio 显式保留）。
- 不降低帧率（无 fps 档位）。
- 图片回填不进本特性。
- 不加 CLI 参数（后续需要再说）。
- 不推送 GitHub、不触发 CI、不打包 DMG。

## 风险与对策

- **判定漂移**：达标判定抽取为单一共享函数，预览与实际执行共用。
- **ffmpeg 版本**：`-fps_mode` 需 ≥5.1；环境为 8.1.2，文档注明最低版本。
- **码率探测缺失**：视为超限走转码（安全方向），单测覆盖。
- **零回归**：所有新参数默认值 = 现状；best 档命令与现状逐 token 一致有单测。
