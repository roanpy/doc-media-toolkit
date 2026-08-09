# PPTX 高清回填质量档位 Implementation Plan（Archived）

> This is a historical implementation record, not a current release
> specification. Current validation and release gates live in
> [`docs/QUALITY_GATES.md`](../../QUALITY_GATES.md) and [`docs/RELEASE.md`](../../RELEASE.md).

**Goal:** 给 PPTX 高清回填加质量档位（最佳/高质量/均衡），默认最佳=现状，其余档按（分辨率上限, CRF, 码率上限, 音频码率）输出，免去回填后二次压缩。

**Architecture:** 上限语义——族源同时满足容器/编码兼容 ∧ 分辨率≤档上限 ∧ 码率≤档上限（best 不查码率）→ 原样嵌入；否则按档位转码。判定逻辑抽为单一共享函数 `_backfill_compatibility`，执行（`_delivery_master`）与预览（`plan_backfill_action`）共用。转码母版只记别名（known_hashes），不注册版本。

**Tech Stack:** Python 3 / unittest / PySide6 GUI / ffmpeg 8.1.2（`-fps_mode` 需 ≥5.1）

## Global Constraints

- 默认档 `best` 行为与现状**完全一致**（含不新增 ffprobe 调用）；所有新参数默认值=现状。
- 回填永远另存新文件，不覆盖输入 PPTX。
- 转码母版**只记别名**，不写入 `variants`。
- 不降低帧率；图片回填不在范围；不加 CLI 参数。
- 测试一律用 `.venv/bin/python`（禁止系统/Homebrew python）；ruff 用 `/opt/homebrew/bin/ruff`。
- 提交仅本地；**不推送 GitHub、不触发 CI、不打包 DMG**。
- 当时的全量验证使用 `.venv/bin/python scripts/run_tests_isolated.py`；当前测试数量和结果以该脚本的实际输出为准。

## 关键现有件（已核实）

- `src/pptx_tools/video_manager.py:333` `_video_envelope(width, height)` → `(1920,1080)` if `width>=height` else `(1080,1920)`；调用方 338（`_archived_dimensions`）、1683（导入归一化）、1887（`_delivery_master`）。
- `src/pptx_tools/video_manager.py:346` `_transcode_high_quality_mp4(source, target, width, height, *, family_id="", limit_1080p=True)`；命令：`ffmpeg -y -i src -map 0:v:0 -map 0:a? -vf scale=W:H:flags=lanczos -c:v libx264 -preset medium -crf 18 -profile:v main -bf 0 -pix_fmt yuv420p -c:a aac -b:a 256k -movflags +faststart [-metadata comment=...]`。
- `src/pptx_tools/video_manager.py:172` `probe_video(path)`：ffprobe `-show_streams -show_format`，返回 width/height/duration_sec/bitrate_kbps/video_codec/audio_codec/has_audio；`run_binary(...).stdout` 为 JSON。
- `src/pptx_tools/video_manager.py:1877` `_delivery_master(family, work)`：`metadata = variant if variant.get("width") else {**variant, **probe_video(source)}`；兼容判定（suffix/h264/aac-or-none/≤包络/偶数）→ 原样返回 `(source, variant["sha256"])`；否则转码到 `work / f"{family['id']}.mp4"`。
- `src/pptx_tools/video_manager.py:1910` `upgrade_pptx_from_library(...)`：默认输出 `{stem}_high_quality.pptx`（1928）；`masters` 缓存键 `family["id"]`（1991-1994）；别名学习 1996-1997；返回 dict（2077-2086）。
- `src/pptx_tools/video_manager_gui.py:363` `PptxUpgradeReviewDialog`：树列 `["PPTX 视频","当前规格","匹配结果","引用","执行计划"]`；`_add_item`（510）初始化 plans 与第 4 列；`_refresh_row`（634）只刷当前行；`decisions()`（719）。
- `src/pptx_tools/video_manager_gui.py:3375` `upgrade_pptx()`：先选输出（单文件 getSaveFileName 默认 `_high_quality`；多文件选目录拼 `_high_quality`）→ 后台 review → `_review_pptx_upgrade`（3437，逐 PPTX 弹对话框）→ 后台 upgrade（3491 传 `output_path`）。
- `family_choices` 构建（3449-3467）：来自 `source_variant(family)`，variant 自带 width/height/bitrate_kbps/video_codec/audio_codec（3764-3767 处证实）。
- MainWindow settings：`QSettings("Doc Media Toolkit", "Doc Media Toolkit")`（1456）。
- 测试模式：`tests/test_video_manager.py` 用 `make_video_pptx` / `no_probe` / patch `pptx_tools.video_manager.probe_video` 与 `._transcode_high_quality_mp4`（fake 签名 `(_source, target, _w, _h, **_kwargs)`）；E2E 在 `tests/test_video_manager_e2e.py`（`ffmpeg_available()` skip、`resolve_binary`）。
- 参考加固语义 `src/pptx_video_compactor.py:1356` `audio_stream_is_usable`：仅当 ffprobe 明确给出 duration≤0 且 frames≤1 才 False，未知一律 True。

---

### Task 1: 档位表 + 转码泛化（含 VFR/空音轨加固）

**Files:**
- Modify: `src/pptx_tools/video_manager.py`（333-395 区域）
- Test: `tests/test_video_manager.py`

**Interfaces:**
- Produces:
  - `BACKFILL_QUALITY_TIERS: dict[str, dict[str, Any]]`（键 best/high/balanced；字段 label/max_width/max_height/crf/bitrate_kbps/audio_bitrate/suffix）
  - `DEFAULT_BACKFILL_TIER = "best"`
  - `_tier_spec(tier: str) -> dict[str, Any]`（非法键抛 ValueError）
  - `_video_envelope(width, height, max_width=1920, max_height=1080)`
  - `_archived_dimensions(width, height, max_width=1920, max_height=1080)`
  - `_audio_stream_usable(path: Path) -> bool`
  - `_transcode_high_quality_mp4(source, target, width, height, *, family_id="", limit_1080p=True, max_width=1920, max_height=1080, crf=18, bitrate_kbps=0, audio_bitrate="256k")`

- [ ] **Step 1: 写失败测试**

在 `tests/test_video_manager.py` 新增测试类（放在文件内现有 delivery 测试附近，import 区补 `BACKFILL_QUALITY_TIERS, DEFAULT_BACKFILL_TIER`）：

```python
class BackfillTranscodeTierTest(unittest.TestCase):
    def test_tier_table_shape_and_best_matches_current(self) -> None:
        self.assertEqual(DEFAULT_BACKFILL_TIER, "best")
        self.assertEqual(set(BACKFILL_QUALITY_TIERS), {"best", "high", "balanced"})
        best = BACKFILL_QUALITY_TIERS["best"]
        self.assertEqual(
            (best["max_width"], best["max_height"], best["crf"]),
            (1920, 1080, 18),
        )
        self.assertEqual(best["bitrate_kbps"], 0)  # 0 = 不设限
        self.assertEqual(best["audio_bitrate"], "256k")
        self.assertEqual(best["suffix"], "high_quality")
        balanced = BACKFILL_QUALITY_TIERS["balanced"]
        self.assertEqual((balanced["max_width"], balanced["max_height"]), (1280, 720))
        self.assertEqual((balanced["crf"], balanced["bitrate_kbps"]), (23, 5000))
        self.assertEqual(balanced["audio_bitrate"], "128k")

    def _capture_command(self, audio_usable: bool = True) -> list[str]:
        from pptx_tools import video_manager as vm

        commands: list[list[str]] = []
        with (
            patch.object(vm, "_audio_stream_usable", return_value=audio_usable),
            patch.object(vm, "run_binary") as run,
        ):
            run.return_value = subprocess.CompletedProcess([], 0, "", "")
            vm._transcode_high_quality_mp4(
                Path("in.mp4"), Path("out.mp4"), 3840, 2160
            )
        return commands or run.call_args.args[0]

    def test_default_command_matches_current_plus_fps_passthrough(self) -> None:
        cmd = self._capture_command()
        self.assertEqual(
            cmd,
            [
                "ffmpeg", "-y", "-i", "in.mp4",
                "-map", "0:v:0", "-map", "0:a?",
                "-vf", "scale=1920:1080:flags=lanczos",
                "-fps_mode", "passthrough",
                "-c:v", "libx264", "-preset", "medium", "-crf", "18",
                "-profile:v", "main", "-bf", "0", "-pix_fmt", "yuv420p",
                "-c:a", "aac", "-b:a", "256k",
                "-movflags", "+faststart",
                "out.mp4",
            ],
        )

    def test_tier_command_applies_crf_maxrate_audio_and_scale(self) -> None:
        from pptx_tools import video_manager as vm

        with (
            patch.object(vm, "_audio_stream_usable", return_value=True),
            patch.object(vm, "run_binary") as run,
        ):
            run.return_value = subprocess.CompletedProcess([], 0, "", "")
            vm._transcode_high_quality_mp4(
                Path("in.mp4"), Path("out.mp4"), 1920, 1080,
                max_width=1280, max_height=720, crf=23,
                bitrate_kbps=5000, audio_bitrate="128k",
            )
        cmd = run.call_args.args[0]
        self.assertIn("scale=1280:720:flags=lanczos", cmd)
        self.assertEqual(cmd[cmd.index("-crf") + 1], "23")
        self.assertEqual(cmd[cmd.index("-maxrate") + 1], "5000k")
        self.assertEqual(cmd[cmd.index("-bufsize") + 1], "10000k")
        self.assertEqual(cmd[cmd.index("-b:a") + 1], "128k")
        # 默认档不出现码率上限
        self.assertNotIn("-maxrate", self._capture_command())

    def test_empty_audio_track_drops_audio_arguments(self) -> None:
        cmd = self._capture_command(audio_usable=False)
        self.assertNotIn("0:a?", cmd)
        self.assertNotIn("-c:a", cmd)
        self.assertNotIn("-b:a", cmd)

    def test_audio_stream_usable_only_false_when_proven_empty(self) -> None:
        from pptx_tools import video_manager as vm

        def fake_run(payload: dict):
            class Result:
                stdout = json.dumps(payload)
            return Result()

        cases = [
            ({"streams": []}, True),  # 无音轨：保持可选映射（现状等价）
            ({"streams": [{}]}, True),  # 字段未知
            ({"streams": [{"duration": "0.000000", "nb_frames": "1"}]}, False),
            ({"streams": [{"duration": "0.000000", "nb_frames": "120"}]}, True),
            ({"streams": [{"duration": "2.5", "nb_frames": "0"}]}, True),
            ({"streams": [{"duration": "N/A", "nb_frames": "0"}]}, True),
        ]
        for payload, expected in cases:
            with patch.object(vm, "run_binary", side_effect=lambda *a, _p=payload, **k: fake_run(_p)):
                self.assertIs(vm._audio_stream_usable(Path("x.mp4")), expected, payload)
        with patch.object(vm, "run_binary", side_effect=RuntimeError("boom")):
            self.assertIs(vm._audio_stream_usable(Path("x.mp4")), True)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager.BackfillTranscodeTierTest -v`
Expected: FAIL（`ImportError` / `AttributeError: BACKFILL_QUALITY_TIERS`）

- [ ] **Step 3: 实现**

`src/pptx_tools/video_manager.py`：

(a) `_video_envelope` / `_archived_dimensions` 泛化（默认值=现状，1683 处调用不变）：

```python
def _video_envelope(
    width: int, height: int, max_width: int = 1920, max_height: int = 1080
) -> tuple[int, int]:
    return (max_width, max_height) if width >= height else (max_height, max_width)


def _archived_dimensions(
    width: int, height: int, max_width: int = 1920, max_height: int = 1080
) -> tuple[int, int]:
    scale = min(1.0, max_width / width, max_height / height)
    return max(2, int(width * scale) // 2 * 2), max(2, int(height * scale) // 2 * 2)
```

（`_archived_dimensions` 函数体保持现有实现，仅加两个默认参数并替换字面量。）

(b) 档位表 + 解析（放在 `_video_envelope` 之前）：

```python
BACKFILL_QUALITY_TIERS: dict[str, dict[str, Any]] = {
    "best": {
        "label": "最佳",
        "max_width": 1920,
        "max_height": 1080,
        "crf": 18,
        "bitrate_kbps": 0,  # 0 = 不设限
        "audio_bitrate": "256k",
        "suffix": "high_quality",
    },
    "high": {
        "label": "高质量",
        "max_width": 1920,
        "max_height": 1080,
        "crf": 20,
        "bitrate_kbps": 12000,
        "audio_bitrate": "256k",
        "suffix": "hq1080p",
    },
    "balanced": {
        "label": "均衡",
        "max_width": 1280,
        "max_height": 720,
        "crf": 23,
        "bitrate_kbps": 5000,
        "audio_bitrate": "128k",
        "suffix": "balanced720p",
    },
}
DEFAULT_BACKFILL_TIER = "best"


def _tier_spec(tier: str) -> dict[str, Any]:
    try:
        return BACKFILL_QUALITY_TIERS[tier]
    except KeyError:
        raise ValueError(f"Unknown backfill quality tier: {tier}") from None
```

(c) 空音轨判定（语义同 compactor `audio_stream_is_usable`，本地小助手）：

```python
def _audio_stream_usable(path: Path) -> bool:
    """Return False only when ffprobe explicitly proves an empty audio track."""
    try:
        result = run_binary(
            [
                "ffprobe",
                "-v",
                "error",
                "-select_streams",
                "a:0",
                "-show_entries",
                "stream=duration,nb_frames",
                "-print_format",
                "json",
                str(path),
            ],
            capture=True,
        )
        payload = json.loads(result.stdout or "{}")
        stream = next(iter(payload.get("streams") or []), None)
    except Exception:
        return True
    if stream is None:
        return True  # 无音轨：-map 0:a? 本就不映射，保持现状
    duration_known = stream.get("duration") not in (None, "", "N/A")
    frames_known = stream.get("nb_frames") not in (None, "", "N/A")
    if not duration_known or not frames_known:
        return True
    try:
        duration = float(stream.get("duration") or 0.0)
        frames = int(stream.get("nb_frames") or 0)
    except (TypeError, ValueError):
        return True
    return duration > 0.0 or frames > 1
```

(d) `_transcode_high_quality_mp4` 泛化 + 加固：

```python
def _transcode_high_quality_mp4(
    source: Path,
    target: Path,
    width: int,
    height: int,
    *,
    family_id: str = "",
    limit_1080p: bool = True,
    max_width: int = 1920,
    max_height: int = 1080,
    crf: int = 18,
    bitrate_kbps: int = 0,
    audio_bitrate: str = "256k",
) -> None:
    output_width, output_height = (
        _archived_dimensions(width, height, max_width, max_height)
        if limit_1080p
        else (max(2, width // 2 * 2), max(2, height // 2 * 2))
    )
    has_audio = _audio_stream_usable(source)
    command = ["ffmpeg", "-y", "-i", str(source), "-map", "0:v:0"]
    if has_audio:
        command += ["-map", "0:a?"]
    command += [
        "-vf",
        f"scale={output_width}:{output_height}:flags=lanczos",
        # 回填不降帧率：保留 VFR 时间戳，避免 FFmpeg 静默丢帧（需 ffmpeg ≥ 5.1）。
        "-fps_mode",
        "passthrough",
        "-c:v",
        "libx264",
        "-preset",
        "medium",
        "-crf",
        str(crf),
        "-profile:v",
        "main",
        # Some WMV files carry sparse variable frame timestamps. B-frame
        # reordering can make MP4 declare a duration shorter than its final
        # presentation timestamp, which truncates playback in PowerPoint.
        "-bf",
        "0",
        "-pix_fmt",
        "yuv420p",
    ]
    if bitrate_kbps > 0:
        command += ["-maxrate", f"{bitrate_kbps}k", "-bufsize", f"{bitrate_kbps * 2}k"]
    if has_audio:
        command += ["-c:a", "aac", "-b:a", audio_bitrate]
    command += ["-movflags", "+faststart"]
    if family_id:
        command.extend(["-metadata", f"comment=pptx-tools-family:{family_id}"])
    run_binary([*command, str(target)], capture=True)
```

注意：现有实现 `-crf 18` 是字面量，改为 `str(crf)`；`command` 构造从单一列表字面量改为分段拼接，token 顺序须与测试断言一致（`-fps_mode` 在 `-vf` 之后、`-c:v` 之前；`-maxrate/-bufsize` 在 `-pix_fmt` 之后、`-c:a` 之前）。

- [ ] **Step 4: 跑测试确认通过 + 回归**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager.BackfillTranscodeTierTest tests.test_video_manager -v 2>&1 | tail -5`
Expected: PASS（含现有 279 中 video_manager 相关全部）

- [ ] **Step 5: 提交**

```bash
cd /path/to/pptx-tools
git add src/pptx_tools/video_manager.py tests/test_video_manager.py
git commit -m "feat: add backfill quality tier table and tier-aware transcode"
```

---

### Task 2: 共享达标判定 + `_delivery_master` 档位化 + `plan_backfill_action`

**Files:**
- Modify: `src/pptx_tools/video_manager.py`（1877-1908）
- Test: `tests/test_video_manager.py`

**Interfaces:**
- Consumes: Task 1 的 `BACKFILL_QUALITY_TIERS` / `_tier_spec` / `_video_envelope` / 泛化后 `_transcode_high_quality_mp4`
- Produces:
  - `_backfill_compatibility(metadata: dict[str, Any], spec: dict[str, Any]) -> bool`（metadata 键：suffix/video_codec/audio_codec/width/height/bitrate_kbps）
  - `plan_backfill_action(metadata: dict[str, Any], tier: str = DEFAULT_BACKFILL_TIER) -> str`
  - `VideoProject._delivery_master(self, family, work, tier: str = DEFAULT_BACKFILL_TIER)`

- [ ] **Step 1: 写失败测试**

```python
class BackfillCompatibilityTest(unittest.TestCase):
    def _meta(self, **overrides):
        metadata = {
            "suffix": ".mp4",
            "video_codec": "h264",
            "audio_codec": "aac",
            "width": 1920,
            "height": 1080,
            "bitrate_kbps": 8000,
        }
        metadata.update(overrides)
        return metadata

    def test_best_tier_matches_current_semantics(self) -> None:
        from pptx_tools.video_manager import _backfill_compatibility, _tier_spec

        spec = _tier_spec("best")
        self.assertTrue(_backfill_compatibility(self._meta(), spec))
        self.assertTrue(_backfill_compatibility(self._meta(bitrate_kbps=50000), spec))
        self.assertFalse(_backfill_compatibility(self._meta(suffix=".wmv"), spec))
        self.assertFalse(_backfill_compatibility(self._meta(video_codec="hevc"), spec))
        self.assertFalse(_backfill_compatibility(self._meta(audio_codec="ac3"), spec))
        self.assertFalse(_backfill_compatibility(self._meta(width=3840, height=2160), spec))
        self.assertFalse(_backfill_compatibility(self._meta(width=1919), spec))  # 奇数
        self.assertFalse(_backfill_compatibility(self._meta(width=0), spec))

    def test_capped_tiers_add_resolution_and_bitrate_ceilings(self) -> None:
        from pptx_tools.video_manager import _backfill_compatibility, _tier_spec

        balanced = _tier_spec("balanced")
        high = _tier_spec("high")
        self.assertFalse(_backfill_compatibility(self._meta(), balanced))  # 1080p 超 720p
        self.assertTrue(
            _backfill_compatibility(self._meta(width=1280, height=720, bitrate_kbps=4000), balanced)
        )
        self.assertFalse(
            _backfill_compatibility(self._meta(width=1280, height=720, bitrate_kbps=8000), balanced)
        )
        self.assertTrue(_backfill_compatibility(self._meta(bitrate_kbps=12000), high))
        self.assertFalse(_backfill_compatibility(self._meta(bitrate_kbps=20000), high))
        # 码率探测缺失：非 best 档视为超限
        self.assertFalse(
            _backfill_compatibility(self._meta(width=1280, height=720, bitrate_kbps=0), balanced)
        )
        # 竖版包络：均衡档 720×1280
        self.assertTrue(
            _backfill_compatibility(
                self._meta(width=720, height=1280, bitrate_kbps=4000), balanced
            )
        )
        self.assertFalse(
            _backfill_compatibility(
                self._meta(width=1080, height=1920, bitrate_kbps=4000), balanced
            )
        )

    def test_plan_backfill_action_texts(self) -> None:
        from pptx_tools.video_manager import plan_backfill_action

        self.assertEqual(
            plan_backfill_action(self._meta(width=1280, height=720, bitrate_kbps=4000), "balanced"),
            "原样嵌入（已达标）",
        )
        self.assertEqual(
            plan_backfill_action(self._meta(), "balanced"), "转码至 ≤720p · CRF 23 · ≤5Mbps"
        )
        self.assertEqual(plan_backfill_action(self._meta(), "best"), "原样嵌入（已达标）")
        self.assertEqual(plan_backfill_action(self._meta(width=0, height=0), "balanced"), "按档位规格回填")
        with self.assertRaises(ValueError):
            plan_backfill_action(self._meta(), "bogus")

    def test_plan_matches_delivery_decision(self) -> None:
        # 同一 metadata：预览分支与共享判定一致
        from pptx_tools.video_manager import _backfill_compatibility, _tier_spec, plan_backfill_action

        for meta, tier in [
            (self._meta(), "best"),
            (self._meta(), "balanced"),
            (self._meta(width=1280, height=720, bitrate_kbps=4000), "balanced"),
            (self._meta(bitrate_kbps=20000), "high"),
        ]:
            compatible = _backfill_compatibility(meta, _tier_spec(tier))
            text = plan_backfill_action(meta, tier)
            self.assertEqual(compatible, text == "原样嵌入（已达标）", (meta, tier))


class DeliveryMasterTierTest(unittest.TestCase):
    def _library_with_source(self, root: Path, metadata: dict) -> tuple:
        source = root / "source.pptx"
        make_video_pptx(source, b"tier-source", "Source")
        library = VideoProject.create(root / "library")
        with patch("pptx_tools.video_manager.probe_video", return_value=metadata):
            library.archive_pptx_videos(source)
        family = library.families()[0]
        return library, family

    def _base_metadata(self) -> dict:
        return {
            **no_probe(Path()),
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "has_audio": True,
            "bitrate_kbps": 8000,
        }

    def test_default_tier_embeds_compatible_source_as_is(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            library, family = self._library_with_source(root, self._base_metadata())
            with patch("pptx_tools.video_manager._transcode_high_quality_mp4") as transcode:
                delivery, _ = library._delivery_master(family, root)
            transcode.assert_not_called()
            self.assertEqual(delivery.suffix.lower(), ".mp4")

    def test_balanced_tier_transcodes_1080p_source_with_tier_params(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            library, family = self._library_with_source(root, self._base_metadata())

            def fake_transcode(_source, target, _w, _h, **kwargs):
                self.assertEqual(kwargs["crf"], 23)
                self.assertEqual(kwargs["bitrate_kbps"], 5000)
                self.assertEqual(kwargs["audio_bitrate"], "128k")
                self.assertEqual((kwargs["max_width"], kwargs["max_height"]), (1280, 720))
                target.write_bytes(b"balanced-delivery")

            with patch(
                "pptx_tools.video_manager._transcode_high_quality_mp4",
                side_effect=fake_transcode,
            ):
                delivery, digest = library._delivery_master(family, root, tier="balanced")
            self.assertEqual(delivery, root / f"{family['id']}.mp4")
            self.assertEqual(digest, hashlib.sha256(b"balanced-delivery").hexdigest())

    def test_balanced_tier_embeds_small_low_bitrate_source(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            metadata = self._base_metadata()
            metadata.update({"width": 1280, "height": 720, "bitrate_kbps": 4000})
            library, family = self._library_with_source(root, metadata)
            with patch("pptx_tools.video_manager._transcode_high_quality_mp4") as transcode:
                delivery, digest = library._delivery_master(family, root, tier="balanced")
            transcode.assert_not_called()
            self.assertEqual(digest, library.source_variant(family)["sha256"])

    def test_high_tier_transcodes_high_bitrate_1080p(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            metadata = self._base_metadata()
            metadata["bitrate_kbps"] = 20000
            library, family = self._library_with_source(root, metadata)

            def fake_transcode(_source, target, _w, _h, **kwargs):
                self.assertEqual(kwargs["bitrate_kbps"], 12000)
                target.write_bytes(b"hq-delivery")

            with patch(
                "pptx_tools.video_manager._transcode_high_quality_mp4",
                side_effect=fake_transcode,
            ):
                library._delivery_master(family, root, tier="high")

    def test_invalid_tier_raises(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            library, family = self._library_with_source(root, self._base_metadata())
            with self.assertRaises(ValueError):
                library._delivery_master(family, root, tier="bogus")
```

注意：`archive_pptx_videos` 会把 `probe_video` 返回值写入 variant（现有测试 718 同款手法），因此 `_delivery_master` 内 `variant.get("bitrate_kbps")` 有值，best 档不会额外 probe。

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager.BackfillCompatibilityTest tests.test_video_manager.DeliveryMasterTierTest -v`
Expected: FAIL（`_backfill_compatibility` 不存在 / `_delivery_master` 无 tier 参数）

- [ ] **Step 3: 实现**

`src/pptx_tools/video_manager.py`（紧跟 `_tier_spec` 之后）：

```python
def _backfill_compatibility(metadata: dict[str, Any], spec: dict[str, Any]) -> bool:
    """Return True when a library source can be embedded as-is for the tier."""
    width = int(metadata.get("width") or 0)
    height = int(metadata.get("height") or 0)
    if width <= 0 or height <= 0:
        return False
    max_width, max_height = _video_envelope(
        width, height, spec["max_width"], spec["max_height"]
    )
    bitrate_cap = int(spec.get("bitrate_kbps") or 0)
    bitrate = int(metadata.get("bitrate_kbps") or 0)
    return (
        str(metadata.get("suffix") or "").lower() in {".mp4", ".m4v"}
        and str(metadata.get("video_codec", "")).lower() == "h264"
        and str(metadata.get("audio_codec", "")).lower() in {"", "aac"}
        and width <= max_width
        and height <= max_height
        and width % 2 == 0
        and height % 2 == 0
        and (bitrate_cap <= 0 or 0 < bitrate <= bitrate_cap)
    )


def plan_backfill_action(
    metadata: dict[str, Any], tier: str = DEFAULT_BACKFILL_TIER
) -> str:
    """Preview the per-item action for a family source under a tier (read-only)."""
    spec = _tier_spec(tier)
    if int(metadata.get("width") or 0) <= 0 or int(metadata.get("height") or 0) <= 0:
        return "按档位规格回填"
    if _backfill_compatibility(metadata, spec):
        return "原样嵌入（已达标）"
    text = f"转码至 ≤{spec['max_height']}p · CRF {spec['crf']}"
    if int(spec.get("bitrate_kbps") or 0) > 0:
        text += f" · ≤{int(spec['bitrate_kbps']) // 1000}Mbps"
    return text
```

`_delivery_master` 改为（保持返回签名不变）：

```python
    def _delivery_master(
        self, family: dict[str, Any], work: Path, tier: str = DEFAULT_BACKFILL_TIER
    ) -> tuple[Path, str]:
        spec = _tier_spec(tier)
        variant = self.source_variant(family)
        source = self.require_variant_path(variant)
        bitrate_cap = int(spec.get("bitrate_kbps") or 0)
        metadata = (
            variant
            if variant.get("width") and (bitrate_cap <= 0 or variant.get("bitrate_kbps"))
            else {**variant, **probe_video(source)}
        )
        width = int(metadata.get("width") or 0)
        height = int(metadata.get("height") or 0)
        if width <= 0 or height <= 0:
            raise RuntimeError(f"Cannot read video dimensions: {source}")
        if _backfill_compatibility({**metadata, "suffix": source.suffix}, spec):
            return source, variant["sha256"]

        target = work / f"{family['id']}.mp4"
        _transcode_high_quality_mp4(
            source,
            target,
            width,
            height,
            family_id=family["id"],
            max_width=spec["max_width"],
            max_height=spec["max_height"],
            crf=spec["crf"],
            bitrate_kbps=int(spec.get("bitrate_kbps") or 0),
            audio_bitrate=spec["audio_bitrate"],
        )
        if not target.is_file() or target.stat().st_size == 0:
            raise RuntimeError(
                f"Failed to create PowerPoint-compatible video: {source}"
            )
        return target, sha256_file(target)
```

要点：best 档（cap=0）且 variant 有 width 时不新增 probe，与原行为逐字节一致；`limit_1080p` 不传（默认 True），best 档 max 默认 1920×1080。

- [ ] **Step 4: 跑测试确认通过 + 回归**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager -v 2>&1 | tail -3`
Expected: OK（含既有 `test_delivery_master_reuses_compatible_m4v_without_transcoding` 等）

- [ ] **Step 5: 提交**

```bash
git add src/pptx_tools/video_manager.py tests/test_video_manager.py
git commit -m "feat: tier-aware delivery master with shared compatibility check"
```

---

### Task 3: `upgrade_pptx_from_library` 接档位 + 输出命名

**Files:**
- Modify: `src/pptx_tools/video_manager.py`（1910-2092）
- Test: `tests/test_video_manager.py`

**Interfaces:**
- Consumes: Task 2 的 `_delivery_master(family, work, tier)` / `_tier_spec`
- Produces: `upgrade_pptx_from_library(..., quality_tier: str = DEFAULT_BACKFILL_TIER)`；返回 dict 新增 `"quality_tier": str`

- [ ] **Step 1: 写失败测试**

```python
class UpgradeQualityTierTest(unittest.TestCase):
    def _setup_library(self, root: Path) -> tuple:
        source = root / "source.pptx"
        compact = root / "compact.pptx"
        make_video_pptx(source, b"tier-source-bytes", "Source")
        make_video_pptx(compact, b"tier-compact-bytes", "Compact")
        library = VideoProject.create(root / "library")
        metadata = {
            **no_probe(Path()),
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "has_audio": True,
            "bitrate_kbps": 8000,
        }
        with patch("pptx_tools.video_manager.probe_video", return_value=metadata):
            library.archive_pptx_videos(source)
        family = library.families()[0]
        family["known_hashes"].append(sha256_file(compact.with_suffix(".mp4")))
        return library, compact, family

    def test_default_output_name_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            library, compact, _ = self._setup_library(Path(temp_dir))
            with patch(
                "pptx_tools.video_manager._transcode_high_quality_mp4",
                side_effect=lambda _s, target, _w, _h, **_k: target.write_bytes(b"x"),
            ):
                result = library.upgrade_pptx_from_library(compact)
            self.assertEqual(result["output_pptx"].name, "compact_high_quality.pptx")
            self.assertEqual(result["quality_tier"], "best")

    def test_tier_suffix_and_result_field(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            library, compact, _ = self._setup_library(Path(temp_dir))

            def fake_transcode(_s, target, _w, _h, **kwargs):
                self.assertEqual(kwargs["crf"], 23)
                target.write_bytes(b"balanced-master")

            with patch(
                "pptx_tools.video_manager._transcode_high_quality_mp4",
                side_effect=fake_transcode,
            ):
                result = library.upgrade_pptx_from_library(
                    compact, quality_tier="balanced"
                )
            self.assertEqual(result["output_pptx"].name, "compact_balanced720p.pptx")
            self.assertEqual(result["quality_tier"], "balanced")
            # 别名学习仍生效：转码母版哈希入 known_hashes
            digest = hashlib.sha256(b"balanced-master").hexdigest()
            self.assertIn(digest, library.families()[0]["known_hashes"])

    def test_invalid_tier_raises_before_writing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            library, compact, _ = self._setup_library(Path(temp_dir))
            with self.assertRaises(ValueError):
                library.upgrade_pptx_from_library(compact, quality_tier="bogus")
            self.assertFalse((Path(temp_dir) / "compact_bogus.pptx").exists())
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager.UpgradeQualityTierTest -v`
Expected: FAIL（`TypeError: unexpected keyword argument 'quality_tier'`）

- [ ] **Step 3: 实现**

`upgrade_pptx_from_library` 修改点：

(a) 签名加 `quality_tier: str = DEFAULT_BACKFILL_TIER`（放在 `keep_current_media` 之后）；
(b) 函数体开头（`input_pptx` resolve 之后）：

```python
        spec = _tier_spec(quality_tier)
```

(c) 默认输出名（1928 行处）：

```python
            else input_pptx.with_name(f"{input_pptx.stem}_{spec['suffix']}.pptx")
```

(d) masters 缓存键（1991-1994）改为族+档位二元组（杜绝未来同运行多档复用隐患）：

```python
        masters: dict[tuple[str, str], tuple[Path, str]] = {}
        ...
                    master_key = (family["id"], quality_tier)
                    master = masters.get(master_key)
                    if master is None:
                        master = self._delivery_master(family, work, quality_tier)
                        masters[master_key] = master
```

(e) 两个 return dict 都加 `"quality_tier": quality_tier`（2041 空结果分支与 2077 正常分支）。

- [ ] **Step 4: 跑测试确认通过 + 回归**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager -v 2>&1 | tail -3`
Expected: OK

- [ ] **Step 5: 提交**

```bash
git add src/pptx_tools/video_manager.py tests/test_video_manager.py
git commit -m "feat: quality_tier option for library backfill with tiered output names"
```

---

### Task 4: GUI 档位选择器 + 执行计划列 + 输出流程重排

**Files:**
- Modify: `src/pptx_tools/video_manager_gui.py`（363-507 对话框；3375-3507 流程）
- Test: `tests/test_verified_regressions.py`

**Interfaces:**
- Consumes: `plan_backfill_action(metadata, tier)`、`BACKFILL_QUALITY_TIERS`、`DEFAULT_BACKFILL_TIER`（Task 1/2）
- Produces:
  - `PptxUpgradeReviewDialog.quality_tier() -> str`
  - `family_choices` 项新增键：width/height/bitrate_kbps/video_codec/audio_codec/suffix
  - MainWindow 流程：输出位置选择移到对话框之后；`upgrade_pptx_from_library(..., quality_tier=...)`

- [ ] **Step 1: 写失败测试（GUI，offscreen）**

`tests/test_verified_regressions.py` 新增（QApplication 夹具沿用 `DesktopLifecycleTest` 的 `setUpClass` 模式；QSettings 用字典假实现隔离；缩略图加载打桩避免触碰真实文件）：

```python
class _FakeSettings:
    store: dict[str, object] = {}

    def __init__(self, *_args: object) -> None:
        pass

    def value(self, key: str, default: object = None) -> object:
        return self.store.get(key, default)

    def setValue(self, key: str, val: object) -> None:
        self.store[key] = val


class BackfillTierDialogTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        _FakeSettings.store.clear()
        self._settings_patch = patch(
            "pptx_tools.video_manager_gui.QSettings", _FakeSettings
        )
        self._thumb_patch = patch(
            "pptx_tools.video_manager_gui._set_video_thumbnail", lambda *a, **k: None
        )
        self._settings_patch.start()
        self._thumb_patch.start()

    def tearDown(self) -> None:
        self._thumb_patch.stop()
        self._settings_patch.stop()

    def _dialog(self):
        from pptx_tools.video_manager_gui import PptxUpgradeReviewDialog

        items = [{
            "media_path": "ppt/media/media1.mp4",
            "metadata": {"width": 320, "height": 240, "duration_sec": 2.0},
            "match_kind": "exact",
            "family_id": "fam1",
            "family_name": "族A",
            "target_source": "/tmp/a.mp4",
            "occurrences": [],
            "source": "/tmp/embedded.mp4",
            "sha256": "x",
        }]
        families = [{
            "id": "fam1", "name": "族A", "source_path": "/tmp/a.mp4",
            "source_sha256": "y", "resolution": "1920×1080",
            "width": 1920, "height": 1080, "bitrate_kbps": 8000,
            "video_codec": "h264", "audio_codec": "aac", "suffix": ".mp4",
        }]
        return PptxUpgradeReviewDialog(None, Path("/tmp/deck.pptx"), items, families)

    def test_default_tier_is_best_and_accessor(self) -> None:
        dialog = self._dialog()
        self.assertEqual(dialog.quality_tier(), "best")

    def test_settings_roundtrip_and_invalid_fallback(self) -> None:
        _FakeSettings.store["backfill/quality_tier"] = "balanced"
        self.assertEqual(self._dialog().quality_tier(), "balanced")
        _FakeSettings.store["backfill/quality_tier"] = "bogus"
        self.assertEqual(self._dialog().quality_tier(), "best")

    def test_tier_switch_persists_and_updates_plan_column(self) -> None:
        dialog = self._dialog()
        row = dialog.tree.topLevelItem(0)
        self.assertIn("原样嵌入", row.text(4))  # best：1080p h264 达标
        dialog.tier_combo.setCurrentIndex(dialog.tier_combo.findData("balanced"))
        self.assertIn("转码至 ≤720p", row.text(4))
        self.assertEqual(_FakeSettings.store["backfill/quality_tier"], "balanced")
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_verified_regressions.BackfillTierDialogTest -v`
Expected: FAIL（`AttributeError: tier_combo` / `quality_tier`）

- [ ] **Step 3: 实现**

(a) import 区补：`from pptx_tools.video_manager import ... BACKFILL_QUALITY_TIERS, DEFAULT_BACKFILL_TIER, plan_backfill_action`（并入现有 import 列表）。

(b) `PptxUpgradeReviewDialog.__init__`：summary 之后、splitter 之前插入档位行：

```python
        tier_row = QHBoxLayout()
        tier_row.addWidget(QLabel("回填质量"))
        self.tier_combo = QComboBox()
        for key, spec in BACKFILL_QUALITY_TIERS.items():
            self.tier_combo.addItem(
                f"{spec['label']}（≤{spec['max_height']}p"
                + (f" · ≤{spec['bitrate_kbps'] // 1000}Mbps" if spec["bitrate_kbps"] else "")
                + "）",
                key,
            )
        saved_tier = str(
            QSettings("Doc Media Toolkit", "Doc Media Toolkit")
            .value("backfill/quality_tier", DEFAULT_BACKFILL_TIER)
        )
        self.tier_combo.setCurrentIndex(
            max(0, self.tier_combo.findData(saved_tier))
        )
        self.tier_combo.currentIndexChanged.connect(self._tier_changed)
        tier_row.addWidget(self.tier_combo, 1)
        layout.addLayout(tier_row)
```

注意：`_add_item` 循环在 `__init__` 尾部、档位行建立之后执行（现有顺序：按钮盒之后 `for item in items: self._add_item(item)`），初始化时 `_add_item` 需要能读档位 → 档位行代码必须放在 `_add_item` 循环之前（满足：在 splitter 前）。

(c) `__init__` 末尾 `self.settings = QSettings("Doc Media Toolkit", "Doc Media Toolkit")`（实例属性，供切换时写回）；上面恢复也改用该属性——把它移到档位行之前。

(d) 新方法：

```python
    def quality_tier(self) -> str:
        tier = str(self.tier_combo.currentData() or DEFAULT_BACKFILL_TIER)
        return tier if tier in BACKFILL_QUALITY_TIERS else DEFAULT_BACKFILL_TIER

    def _tier_changed(self) -> None:
        self.settings.setValue("backfill/quality_tier", self.quality_tier())
        self._refresh_all_rows()

    def _refresh_all_rows(self) -> None:
        for index in range(self.tree.topLevelItemCount()):
            self._refresh_row(self.tree.topLevelItem(index))

    def _family_metadata(self, family: dict[str, Any]) -> dict[str, Any]:
        return {
            "suffix": family.get("suffix") or "",
            "video_codec": family.get("video_codec") or "",
            "audio_codec": family.get("audio_codec") or "",
            "width": int(family.get("width") or 0),
            "height": int(family.get("height") or 0),
            "bitrate_kbps": int(family.get("bitrate_kbps") or 0),
        }

    def _refresh_row(self, row: QTreeWidgetItem | None = None) -> None:
        row = row or self.tree.currentItem()
        if row is None:
            return
        item = row.data(0, Qt.ItemDataRole.UserRole)
        plan = self.plans[str(item["media_path"])]
        family = self._family_by_id(plan.get("family_id"))
        if plan["action"] == "keep":
            text = "保持当前"
        elif family is None:
            text = "请选择目标视频族"
        else:
            base = (
                f"关联并回填 · {family['name']}"
                if plan["action"] == "remember"
                else f"仅本次回填 · {family['name']}"
            )
            text = f"{base} · {plan_backfill_action(self._family_metadata(family), self.quality_tier())}"
        row.setText(4, text)
```

（原 `_refresh_row(self)` 签名改为可选参数；`_family_changed`/`_action_changed` 调用处不变。`_add_item` 中初始第 4 列文本改为调 `self._refresh_row(row)` 统一生成——`_add_item` 末尾 `self.tree.addTopLevelItem(row)` 后调用。）

(e) `_review_pptx_upgrade`：`family_choices.append({...})` 增加键：

```python
                        "width": int(variant.get("width") or 0),
                        "height": int(variant.get("height") or 0),
                        "bitrate_kbps": int(variant.get("bitrate_kbps") or 0),
                        "video_codec": str(variant.get("video_codec") or ""),
                        "audio_codec": str(variant.get("audio_codec") or ""),
                        "suffix": source_path.suffix.lower(),
```

(f) 输出流程重排（`upgrade_pptx` + `_review_pptx_upgrade`）：
- `upgrade_pptx` 中删除 3386-3405 的输出选择块，改为只记录 `single = len(paths) == 1`；review 完成后调 `_review_pptx_upgrade(paths, analyses, review_root)`（去掉 outputs 参数）。
- `_review_pptx_upgrade`：对话框循环中收集 `tiers[source] = dialog.quality_tier()`；全部接受后：

```python
        if single:
            source = paths[0]
            suffix = BACKFILL_QUALITY_TIERS[tiers[source]]["suffix"]
            default = source.with_name(f"{source.stem}_{suffix}.pptx")
            output, _ = QFileDialog.getSaveFileName(
                self, "保存高清优化 PPTX", str(default), "PowerPoint (*.pptx)"
            )
            if not output:
                shutil.rmtree(review_root, ignore_errors=True)
                return
            outputs = {source: Path(output)}
        else:
            directory = QFileDialog.getExistingDirectory(
                self, "选择高清优化输出目录", str(paths[0].parent)
            )
            if not directory:
                shutil.rmtree(review_root, ignore_errors=True)
                return
            outputs = {
                source: Path(directory)
                / f"{source.stem}_{BACKFILL_QUALITY_TIERS[tiers[source]]['suffix']}.pptx"
                for source in paths
            }
```

- operation 中 `upgrade_pptx_from_library(..., quality_tier=tiers[source], ...)`。
- 取消语义保持：任一对话框 reject → 清 review_root、记日志、return（现有逻辑保留，仅去掉对 outputs 的依赖）。

(g) `_upgrade_finished`：生成的结果若含非最佳档，日志追加一行档位说明：

```python
        labels = sorted(
            {
                spec["label"]
                for item in generated
                for spec in [
                    BACKFILL_QUALITY_TIERS.get(item.get("quality_tier") or "best")
                ]
                if spec and spec["label"] != "最佳"
            }
        )
        if labels:
            self.append_log(f"回填档位：{'、'.join(labels)}")
```

- [ ] **Step 4: 跑测试确认通过 + 回归**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_verified_regressions -v 2>&1 | tail -3`
Expected: OK

- [ ] **Step 5: 提交**

```bash
git add src/pptx_tools/video_manager_gui.py tests/test_verified_regressions.py
git commit -m "feat: quality tier selector in backfill confirm dialog"
```

---

### Task 5: E2E（真实 ffmpeg）

**Files:**
- Test: `tests/test_video_manager_e2e.py`

**Interfaces:**
- Consumes: 全部前序任务；现有 E2E 夹具（`ffmpeg_available`、`resolve_binary`、真实视频生成助手、库归档流程，参考 445 行 `test_library_archives_source_and_upgrades_real_compacted_pptx`）

- [ ] **Step 1: 写 E2E 测试**

`tests/test_video_manager_e2e.py` 的 E2E 类（`@unittest.skipUnless(ffmpeg_available(), ...)` 那个）新增，复用模块级助手 `make_real_video` / `make_video_pptx` / `transcode_video`：

```python
    def test_backfill_balanced_tier_transcodes_to_720p(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            video = root / "source.mp4"
            make_real_video(video, seconds=2.0, size="1920x1080")
            original_bytes = video.read_bytes()
            source = root / "source.pptx"
            make_video_pptx(source, original_bytes, "Tier E2E")

            library = VideoProject.create(root / "library")
            library.archive_pptx_videos(source)  # 真实 probe：1080p/h264/aac
            family = library.families()[0]

            low = root / "low.mp4"
            transcode_video(video, low, audio_frequency=None)  # 320×180 aac
            compact = root / "compact.pptx"
            make_video_pptx(compact, low.read_bytes(), "Compact")
            family["known_hashes"].append(sha256_file(low))

            scanned = scan_embedded_videos(compact)
            media_path = next(iter(scanned))

            result = library.upgrade_pptx_from_library(compact, quality_tier="balanced")
            self.assertEqual(result["matched"], 1)
            self.assertEqual(result["quality_tier"], "balanced")
            output = Path(result["output_pptx"])
            self.assertEqual(output.name, "compact_balanced720p.pptx")

            embedded = root / "embedded.mp4"
            with ZipFile(output) as archive:
                embedded.write_bytes(archive.read(media_path))
            probed = probe_video(embedded)
            self.assertLessEqual(probed["width"], 1280)
            self.assertLessEqual(probed["height"], 720)
            self.assertGreater(probed["width"], 0)
            self.assertEqual(probed["video_codec"], "h264")
            self.assertEqual(probed["audio_codec"], "aac")
            # 别名学习：均衡母版哈希入库，未注册新版本
            self.assertEqual(len(library.families()[0]["variants"]), 1)

            # 对照：默认最佳档对 1080p h264/aac 源原样嵌入（零损失）
            control = library.upgrade_pptx_from_library(compact)
            self.assertEqual(control["quality_tier"], "best")
            with ZipFile(control["output_pptx"]) as archive:
                self.assertEqual(archive.read(media_path), original_bytes)
```

（`sha256_file`、`probe_video` 已在该测试文件的 import 区或需补入；`ZipFile`、`scan_embedded_videos` 现有 import 已有。）

- [ ] **Step 2: 跑该测试确认通过**

Run: `cd /path/to/pptx-tools && .venv/bin/python -m unittest tests.test_video_manager_e2e -v 2>&1 | tail -5`
Expected: OK（含既有 E2E 全绿；ffmpeg 缺失时 skip）

- [ ] **Step 3: 提交**

```bash
git add tests/test_video_manager_e2e.py
git commit -m "test: end-to-end balanced tier backfill transcodes to 720p"
```

---

### Task 6: 文档同步 + 全量门槛

**Files:**
- Modify: `README.md`、`docs/HANDOFF.md`、`docs/ARCHITECTURE.md`、`CHANGELOG.md`
- 已有: `docs/BACKFILL_QUALITY_TIERS.md`（spec，随本任务提交）

- [ ] **Step 1: README**：高清回填小节加档位表（三档：上限/CRF/码率/音频/后缀）+「默认最佳=原行为；码率探测缺失按超限转码」说明。
- [ ] **Step 2: docs/HANDOFF.md**：新增本特性条目（范围、关键文件:行、档位表、ffmpeg ≥5.1 要求）。
- [ ] **Step 3: docs/ARCHITECTURE.md**：回填流程小节补「档位上限语义 + 共享 `_backfill_compatibility`」一句架构说明。
- [ ] **Step 4: CHANGELOG.md**：Unreleased/新版本条目：回填质量档位 + VFR/空音轨加固 + 对话框档位选择。
- [ ] **Step 5: 全量门槛**

```bash
cd /path/to/pptx-tools
.venv/bin/python scripts/run_tests_isolated.py            # 期望全绿（279+新增）
/opt/homebrew/bin/ruff check src tests                     # All checks passed
/opt/homebrew/bin/ruff format --check src tests            # no changes
.venv/bin/python -m compileall src tests -q                # 无输出
git diff --check                                           # 无输出
```

- [ ] **Step 6: 提交**

```bash
git add README.md docs/HANDOFF.md docs/ARCHITECTURE.md CHANGELOG.md docs/BACKFILL_QUALITY_TIERS.md
git commit -m "docs: backfill quality tiers"
```

---

## 风险与对策

- **零回归**：Task 1 命令断言锁定默认命令（含新增 `-fps_mode` 的明确基线）；best 档不新增 probe（`_delivery_master` cap=0 短路）。
- **判定漂移**：`_backfill_compatibility` 唯一来源，预览/执行共用，并有 `test_plan_matches_delivery_decision` 对拍。
- **GUI 流程重排**：输出选择移到确认对话框之后；取消路径都清理 review_root，测试覆盖对话框层。
- **E2E 时长**：1080p 2s 源 + preset medium，秒级；E2E 文件已有更大用例。
