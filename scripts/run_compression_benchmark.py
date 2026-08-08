#!/usr/bin/env python3
"""Local manifest-driven compression benchmark entry.

Runs the existing smart-target compression core (``pptx_video_compactor.compact_input_path``)
over a user-supplied, sanitized sample manifest and aggregates the results into a
machine-readable JSON report plus a human-readable Markdown summary.

Design contract (see ``docs/COMPRESSION_BENCHMARK.md``):

- Sample files are referenced by absolute paths in the manifest; they are NOT copied
  into the repository and must never be committed. The manifest path itself is also
  expected to live outside the working tree (e.g. ``~/benchmarks/...``).
- This script does not modify the compression core. It only constructs the same
  ``argparse.Namespace`` the CLI builds and reads back the per-file ``.report.json``
  the core already writes.
- No new runtime dependencies: stdlib + the project's own modules only.

The benchmark records, per sample and in aggregate:

- target capacity error (``target_bytes`` vs ``actual_bytes``), ratio and status;
- real output capacity (measured from the produced file);
- per-asset quality / structure results (``quality_status``, ``quality_reason``,
  ``applied_threshold``, restored-original flags);
- correction rounds, derived from ``presentation.target_capacity_attempts``;
- wall-clock duration per sample;
- CPU/GPU encoder usage and fallback signals, derived from per-asset ``status``
  (``encoded`` vs ``encoded_gpu``) and ``quality_reason``.

Usage::

    python scripts/run_compression_benchmark.py --manifest /path/to/manifest.json
    python scripts/run_compression_benchmark.py --manifest ... --output-dir ./bench-out
    python scripts/run_compression_benchmark.py --self-check
    python scripts/run_compression_benchmark.py --help
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import platform
import statistics
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

# The compression core lives as a top-level module under src/.
_SRC = Path(__file__).resolve().parents[1] / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

# Imported after sys.path mutation so the project module is resolvable.
from pptx_video_compactor import (  # noqa: E402
    DEFAULT_OVERSCAN,
    compact_input_path,
)


BENCHMARK_VERSION = 1
DEFAULT_OUTPUT_DIR = Path("benchmark-results")
VIDEO_PROFILES = {"none", "high", "balanced", "aggressive"}
IMAGE_PROFILES = {"none", "lossless", "high", "balanced", "aggressive"}
ENCODER_MODES = {"auto", "cpu", "gpu"}
QUALITY_MODES = {"safe", "forced"}


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _sha256(path: Path) -> str | None:
    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass
class SampleSpec:
    """One entry in the benchmark manifest."""

    path: str
    target_size_mb: float | None = None
    label: str | None = None
    profile: str = "high"
    image_profile: str = "high"
    encoder: str = "auto"
    quality_mode: str = "safe"
    video_ssim_threshold: float = 0.95
    image_ssim_threshold: float = 0.99
    preset: str = "medium"
    reserve_mb: float | None = None
    extra: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_manifest_entry(
        cls,
        entry: dict[str, Any],
        *,
        default_encoder: str = "auto",
        default_quality_mode: str = "safe",
    ) -> "SampleSpec":
        if not isinstance(entry, dict):
            raise ValueError("manifest sample entries must be JSON objects")
        path = entry.get("path")
        if not isinstance(path, str) or not path.strip():
            raise ValueError("each manifest sample needs a non-empty 'path' string")
        if not Path(path).expanduser().is_absolute():
            raise ValueError(f"sample path must be absolute: {path!r}")
        target_size_mb = entry.get("target_size_mb")
        reserve_mb = entry.get("reserve_mb")
        video_ssim_threshold = entry.get("video_ssim_threshold", 0.95)
        image_ssim_threshold = entry.get("image_ssim_threshold", 0.99)
        if target_size_mb is not None and (
            not isinstance(target_size_mb, (int, float))
            or isinstance(target_size_mb, bool)
            or not math.isfinite(float(target_size_mb))
            or float(target_size_mb) <= 0
        ):
            raise ValueError("target_size_mb must be a positive finite number")
        if reserve_mb is not None and (
            not isinstance(reserve_mb, (int, float))
            or isinstance(reserve_mb, bool)
            or not math.isfinite(float(reserve_mb))
            or float(reserve_mb) < 0
        ):
            raise ValueError("reserve_mb must be a non-negative finite number")
        for field_name, threshold in (
            ("video_ssim_threshold", video_ssim_threshold),
            ("image_ssim_threshold", image_ssim_threshold),
        ):
            if (
                not isinstance(threshold, (int, float))
                or isinstance(threshold, bool)
                or not math.isfinite(float(threshold))
                or not 0 < float(threshold) <= 1
            ):
                raise ValueError(f"{field_name} must be a number in (0, 1]")
        profile = entry.get("profile", "high")
        image_profile = entry.get("image_profile", "high")
        encoder = entry.get("encoder", default_encoder)
        quality_mode = entry.get("quality_mode", default_quality_mode)
        if profile not in VIDEO_PROFILES:
            raise ValueError(f"profile must be one of {sorted(VIDEO_PROFILES)}")
        if image_profile not in IMAGE_PROFILES:
            raise ValueError(f"image_profile must be one of {sorted(IMAGE_PROFILES)}")
        if encoder not in ENCODER_MODES:
            raise ValueError(f"encoder must be one of {sorted(ENCODER_MODES)}")
        if quality_mode not in QUALITY_MODES:
            raise ValueError(f"quality_mode must be one of {sorted(QUALITY_MODES)}")
        known = {
            "path",
            "target_size_mb",
            "label",
            "profile",
            "image_profile",
            "encoder",
            "quality_mode",
            "video_ssim_threshold",
            "image_ssim_threshold",
            "preset",
            "reserve_mb",
        }
        extra = {
            key: value
            for key, value in entry.items()
            if key not in known
            and key.startswith("meta_")
            and isinstance(value, (str, int, float, bool))
        }
        return cls(
            path=path,
            target_size_mb=target_size_mb,
            label=entry.get("label"),
            profile=profile,
            image_profile=image_profile,
            encoder=encoder,
            quality_mode=quality_mode,
            video_ssim_threshold=video_ssim_threshold,
            image_ssim_threshold=image_ssim_threshold,
            preset=entry.get("preset", "medium"),
            reserve_mb=reserve_mb,
            extra=extra,
        )

    def label_or_name(self) -> str:
        return self.label or Path(self.path).name


def load_manifest(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise SystemExit(f"Manifest not found: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Manifest is not valid JSON ({path}): {exc}") from exc
    if not isinstance(data, dict):
        raise SystemExit("Manifest root must be a JSON object")
    samples = data.get("samples")
    if not isinstance(samples, list) or not samples:
        raise SystemExit("Manifest must contain a non-empty 'samples' array")
    return data


def build_namespace(spec: SampleSpec, output_dir: Path) -> argparse.Namespace:
    """Construct the same argparse.Namespace the compact CLI builds per file."""
    input_path = Path(spec.path).expanduser().resolve()
    if not input_path.is_file():
        raise SystemExit(f"Sample file not found: {input_path}")
    # Output beside the sample would pollute the user's sample directory. Use a
    # content-derived name to avoid collisions and avoid leaking the source path/name.
    input_hash = _sha256(input_path)
    if input_hash is None:
        raise SystemExit(f"Unable to hash sample file: {input_path}")
    target_label = ""
    if spec.target_size_mb is not None:
        target_label = f"_{spec.target_size_mb:g}MB"
    output_path = (
        output_dir / f"sample_{input_hash[:12]}{target_label}_bench{input_path.suffix}"
    )
    return argparse.Namespace(
        input_pptx=input_path,
        target_size_mb=spec.target_size_mb,
        config=None,
        profile=spec.profile,
        image_profile=spec.image_profile,
        output=output_path,
        video_output_dir=None,
        slide_render_width=1920,
        slide_render_height=1080,
        min_height=480,
        max_height=1080,
        overscan=DEFAULT_OVERSCAN,
        reserve_mb=spec.reserve_mb,
        preset=spec.preset,
        encoder=spec.encoder,
        video_ssim_threshold=spec.video_ssim_threshold,
        image_ssim_threshold=spec.image_ssim_threshold,
        quality_mode=spec.quality_mode,
        work_dir=None,
        keep_work_dir=False,
        keep_artifacts=False,
        dry_run=False,
    )


def _count_correction_rounds(report: dict[str, Any]) -> int:
    attempts = report.get("presentation", {}).get("target_capacity_attempts", [])
    if not isinstance(attempts, list):
        return 0
    return sum(
        1
        for item in attempts
        if isinstance(item, dict) and item.get("kind") == "correction"
    )


def _count_giveback_rounds(report: dict[str, Any]) -> int:
    attempts = report.get("presentation", {}).get("target_capacity_attempts", [])
    if not isinstance(attempts, list):
        return 0
    return sum(
        1
        for item in attempts
        if isinstance(item, dict) and item.get("kind") == "quality_giveback"
    )


def _summarize_assets(report: dict[str, Any]) -> dict[str, Any]:
    videos = report.get("videos") or []
    images = report.get("images") or []
    entries = [*videos, *images]

    def tally(field_name: str) -> dict[str, int]:
        counts: dict[str, int] = {}
        for item in entries:
            if not isinstance(item, dict):
                continue
            value = item.get(field_name)
            if value is None:
                continue
            counts[str(value)] = counts.get(str(value), 0) + 1
        return counts

    statuses = tally("status")
    quality_statuses = tally("quality_status")
    gpu_used = any(
        item.get("status") == "encoded_gpu"
        for item in entries
        if isinstance(item, dict)
    )
    restored = sum(
        1
        for item in entries
        if isinstance(item, dict) and item.get("quality_status") == "restored_original"
    )
    below_threshold = sum(
        1
        for item in entries
        if isinstance(item, dict) and item.get("quality_status") == "below_threshold"
    )
    return {
        "asset_count": len(entries),
        "video_count": len(videos),
        "image_count": len(images),
        "status_counts": statuses,
        "quality_status_counts": quality_statuses,
        "gpu_encoded_assets": statuses.get("encoded_gpu", 0),
        "gpu_used": gpu_used,
        "restored_original_assets": restored,
        "below_threshold_assets": below_threshold,
    }


def run_one_sample(
    spec: SampleSpec,
    output_dir: Path,
    quiet: bool,
) -> dict[str, Any]:
    """Run the compression core for one sample and collect benchmark fields."""
    output_dir.mkdir(parents=True, exist_ok=True)
    input_path = Path(spec.path).expanduser().resolve()
    ns: argparse.Namespace | None = None
    started = time.perf_counter()
    error: str | None = None
    result: dict[str, Any] = {}
    report_json: dict[str, Any] = {}
    report_path: Path | None = None

    def logger(message: str) -> None:
        if not quiet:
            print(f"  [{spec.label_or_name()}] {message}")

    try:
        ns = build_namespace(spec, output_dir)
        result = compact_input_path(ns, logger=logger)
        raw_report_path = result.get("report_path")
        report_path = Path(raw_report_path) if raw_report_path else None
        if report_path and report_path.is_file():
            report_json = json.loads(report_path.read_text(encoding="utf-8"))
    except SystemExit as exc:
        error = f"SystemExit: {exc}"
    except Exception as exc:  # noqa: BLE001 - benchmark must record, not crash, per-sample failures
        error = f"{type(exc).__name__}: {exc}"
    elapsed = time.perf_counter() - started

    target = report_json.get("target") or {}
    actual_bytes = target.get("actual_bytes")
    target_bytes = target.get("target_bytes")
    delta_bytes = target.get("delta_bytes")
    output_pptx_str = result.get("output_pptx") or report_json.get("output_pptx")
    output_path = Path(output_pptx_str) if output_pptx_str else None
    measured_bytes = (
        output_path.stat().st_size
        if output_path and output_path.is_file()
        else actual_bytes
    )

    entry: dict[str, Any] = {
        "label": spec.label_or_name(),
        "input_path": input_path.name,
        "input_sha256": _sha256(input_path),
        "output_path": output_path.name if output_path else None,
        "output_sha256": _sha256(output_path) if output_path else None,
        "target_size_mb": spec.target_size_mb,
        "target_bytes": target_bytes,
        "actual_bytes": measured_bytes,
        "delta_bytes": delta_bytes
        if delta_bytes is not None
        else (
            None
            if measured_bytes is None or target_bytes is None
            else measured_bytes - target_bytes
        ),
        "target_ratio": target.get("target_ratio"),
        "target_status": target.get("status"),
        "correction_rounds": _count_correction_rounds(report_json),
        "quality_giveback_rounds": _count_giveback_rounds(report_json),
        "elapsed_sec": round(elapsed, 3),
        "encoder_requested": spec.encoder,
        "assets": _summarize_assets(report_json),
        "skipped": bool(result.get("skipped")),
        "skip_reason": result.get("reason")
        or report_json.get("presentation", {}).get("reason"),
        "report_path": report_path.name if report_path else None,
        "error": error,
    }
    # Preserve manifest-declared extras for traceability without leaking secrets.
    if spec.extra:
        entry["manifest_extra"] = spec.extra
    return entry


def _aggregate(entries: list[dict[str, Any]]) -> dict[str, Any]:
    completed = [e for e in entries if e.get("error") is None and not e.get("skipped")]
    ratios = [e["target_ratio"] for e in completed if e.get("target_ratio") is not None]
    deltas = [e["delta_bytes"] for e in completed if e.get("delta_bytes") is not None]
    elapsed = [e["elapsed_sec"] for e in entries if e.get("elapsed_sec") is not None]
    gpu_any = any(e.get("assets", {}).get("gpu_used") for e in entries)
    total_restored = sum(
        e.get("assets", {}).get("restored_original_assets", 0) for e in entries
    )
    total_below = sum(
        e.get("assets", {}).get("below_threshold_assets", 0) for e in entries
    )
    total_correction = sum(e.get("correction_rounds", 0) for e in entries)
    return {
        "sample_count": len(entries),
        "completed_count": len(completed),
        "skipped_count": sum(1 for e in entries if e.get("skipped")),
        "error_count": sum(1 for e in entries if e.get("error")),
        "gpu_used_any": gpu_any,
        "total_correction_rounds": total_correction,
        "total_restored_original_assets": total_restored,
        "total_below_threshold_assets": total_below,
        "target_ratio_mean": round(statistics.fmean(ratios), 6) if ratios else None,
        "target_ratio_min": round(min(ratios), 6) if ratios else None,
        "target_ratio_max": round(max(ratios), 6) if ratios else None,
        "delta_bytes_mean": round(statistics.fmean(deltas)) if deltas else None,
        "elapsed_sec_total": round(sum(elapsed), 3) if elapsed else None,
        "elapsed_sec_mean": round(statistics.fmean(elapsed), 3) if elapsed else None,
    }


def build_report(
    manifest: dict[str, Any],
    entries: list[dict[str, Any]],
    args_ns: argparse.Namespace,
) -> dict[str, Any]:
    return {
        "schema": "pptx-tools.compression-benchmark",
        "version": BENCHMARK_VERSION,
        "generated_at": _utc_now_iso(),
        "host": {
            "platform": platform.platform(),
            "python": sys.version.split()[0],
            "machine": platform.machine(),
        },
        "manifest_notes": manifest.get("notes"),
        "encoder_default": args_ns.encoder,
        "quality_mode_default": args_ns.quality_mode,
        "samples": entries,
        "aggregate": _aggregate(entries),
    }


def write_markdown(report: dict[str, Any], path: Path) -> None:
    agg = report["aggregate"]
    lines = [
        "# 压缩基准报告",
        "",
        f"- 生成时间：`{report['generated_at']}`",
        f"- 主机：`{report['host']['platform']}` / Python {report['host']['python']} / {report['host']['machine']}",
        f"- 样本数：{agg['sample_count']}（完成 {agg['completed_count']}，跳过 {agg['skipped_count']}，错误 {agg['error_count']}）",
        f"- 默认编码器：`{report.get('encoder_default', 'auto')}`，质量模式：`{report.get('quality_mode_default', 'safe')}`",
        "",
        "## 汇总",
        "",
        f"- 是否使用 GPU：`{agg['gpu_used_any']}`",
        f"- 纠偏轮数合计：{agg['total_correction_rounds']}",
        f"- 恢复原件素材数：{agg['total_restored_original_assets']}",
        f"- below_threshold 素材数：{agg['total_below_threshold_assets']}",
        f"- 目标容量比率：均值 {agg['target_ratio_mean']}，最小 {agg['target_ratio_min']}，最大 {agg['target_ratio_max']}",
        f"- 容量差值均值（bytes）：{agg['delta_bytes_mean']}",
        f"- 耗时：合计 {agg['elapsed_sec_total']} s，均值 {agg['elapsed_sec_mean']} s",
        "",
        "## 逐样本",
        "",
        "| 样本 | 目标(MB) | 实际(bytes) | 差值 | 比率 | 状态 | 纠偏 | 耗时(s) | GPU | 错误 |",
        "| --- | ---: | ---: | ---: | ---: | --- | ---: | ---: | :---: | --- |",
    ]
    for entry in report["samples"]:
        assets = entry.get("assets") or {}
        lines.append(
            "| {label} | {target} | {actual} | {delta} | {ratio} | {status} | {rounds} | {elapsed} | {gpu} | {error} |".format(
                label=entry.get("label", ""),
                target=entry.get("target_size_mb")
                if entry.get("target_size_mb") is not None
                else "-",
                actual=entry.get("actual_bytes")
                if entry.get("actual_bytes") is not None
                else "-",
                delta=entry.get("delta_bytes")
                if entry.get("delta_bytes") is not None
                else "-",
                ratio=entry.get("target_ratio")
                if entry.get("target_ratio") is not None
                else "-",
                status=entry.get("target_status") or "-",
                rounds=entry.get("correction_rounds", 0),
                elapsed=entry.get("elapsed_sec", "-"),
                gpu="是" if assets.get("gpu_used") else "否",
                error=entry.get("error") or "",
            )
        )
    lines.extend(
        ["", "## 说明", "", "- 样本文件不在仓库内；本报告仅记录路径哈希与结果。"]
    )
    if any(e.get("error") for e in report["samples"]):
        lines.append("- 存在错误样本：见 JSON 中各 `error` 字段。")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run a local manifest-driven compression benchmark over sanitized samples. "
            "Sample files are referenced by path and never committed to the repository."
        )
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        help="Path to the benchmark manifest JSON (expected outside the working tree).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for benchmark outputs and per-sample compressed files.",
    )
    parser.add_argument(
        "--encoder",
        choices=["auto", "cpu", "gpu"],
        default="auto",
        help="Default video encoder mode when the manifest does not override it.",
    )
    parser.add_argument(
        "--quality-mode",
        choices=["safe", "forced"],
        default="safe",
        help="Default quality mode when the manifest does not override it.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress per-sample progress lines.",
    )
    parser.add_argument(
        "--self-check",
        action="store_true",
        help="Do not run any sample. Verify imports and emit an empty result skeleton.",
    )
    return parser.parse_args(argv)


def _empty_skeleton(args_ns: argparse.Namespace) -> dict[str, Any]:
    return {
        "schema": "pptx-tools.compression-benchmark",
        "version": BENCHMARK_VERSION,
        "generated_at": _utc_now_iso(),
        "host": {
            "platform": platform.platform(),
            "python": sys.version.split()[0],
            "machine": platform.machine(),
        },
        "manifest_notes": "self-check: no samples run",
        "encoder_default": args_ns.encoder,
        "quality_mode_default": args_ns.quality_mode,
        "samples": [],
        "aggregate": _aggregate([]),
    }


def main(argv: list[str] | None = None) -> int:
    args_ns = parse_args(argv)

    if args_ns.self_check:
        # Verify the import chain and DEFAULT_OVERSCAN are reachable without samples.
        _ = compact_input_path
        _ = DEFAULT_OVERSCAN
        skeleton = _empty_skeleton(args_ns)
        args_ns.output_dir.mkdir(parents=True, exist_ok=True)
        json_path = args_ns.output_dir / "benchmark-self-check.json"
        md_path = args_ns.output_dir / "benchmark-self-check.md"
        json_path.write_text(
            json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        write_markdown(skeleton, md_path)
        print(f"self-check OK: imports resolved, skeleton written to {json_path}")
        return 0

    if args_ns.manifest is None:
        print("--manifest is required unless --self-check is used.", file=sys.stderr)
        return 2

    manifest = load_manifest(args_ns.manifest)
    try:
        specs = [
            SampleSpec.from_manifest_entry(
                entry,
                default_encoder=args_ns.encoder,
                default_quality_mode=args_ns.quality_mode,
            )
            for entry in manifest["samples"]
        ]
    except (TypeError, ValueError) as exc:
        raise SystemExit(f"Invalid benchmark manifest: {exc}") from exc

    args_ns.output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Benchmark: {len(specs)} sample(s) -> {args_ns.output_dir}")
    entries: list[dict[str, Any]] = []
    for spec in specs:
        entries.append(run_one_sample(spec, args_ns.output_dir, args_ns.quiet))

    report = build_report(manifest, entries, args_ns)
    json_path = args_ns.output_dir / "benchmark.json"
    md_path = args_ns.output_dir / "benchmark.md"
    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_markdown(report, md_path)
    print(f"Wrote {json_path}")
    print(f"Wrote {md_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
