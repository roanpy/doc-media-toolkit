#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import importlib
import json
import math
import mimetypes
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable
from zipfile import ZipFile, ZipInfo

from PIL import (
    Image,
    ImageChops,
    ImageFilter,
    ImageStat,
    PngImagePlugin,
    UnidentifiedImageError,
)
from defusedxml import ElementTree as SafeET

from pptx_output_watermark.process_utils import (
    finish_process,
    run_process,
    start_process,
    terminate_process,
)


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
STD_HEIGHTS = [480, 720, 1080]
DEFAULT_OVERSCAN = 1.2
BYTES_PER_MB = 1_000_000
GPU_ENCODERS = ("h264_videotoolbox", "h264_nvenc", "h264_qsv", "h264_amf", "h264_mf")
PROFILE_QUALITY_RULES = {
    "none": {
        "max_height": 1080,
        "min_height": 480,
        "dynamic_height": False,
        "bitrate_bias": 1.0,
        "estimate_ratio": 1.0,
        "fps_large": None,
        "fps_medium": None,
        "fps_small": None,
    },
    "high": {
        "max_height": 1080,
        "min_height": 480,
        "dynamic_height": False,
        "bitrate_bias": 0.90,
        "estimate_ratio": 0.76,
        "fps_large": None,
        "fps_medium": None,
        "fps_small": None,
    },
    "balanced": {
        "max_height": 1080,
        "min_height": 720,
        "dynamic_height": True,
        "bitrate_bias": 0.62,
        "estimate_ratio": 0.50,
        "fps_large": 30.0,
        "fps_medium": 30.0,
        "fps_small": 30.0,
    },
    "aggressive": {
        "max_height": 720,
        "min_height": 480,
        "dynamic_height": True,
        "bitrate_bias": 0.38,
        "estimate_ratio": 0.32,
        "fps_large": 24.0,
        "fps_medium": 24.0,
        "fps_small": 24.0,
    },
}
IMAGE_QUALITY_RULES = {
    "none": {
        "quality": 100,
        "target_min_quality": 100,
        "target_min_scale": 1.0,
        "estimate_ratio": 1.0,
    },
    "lossless": {
        "quality": 100,
        "target_min_quality": 100,
        "target_min_scale": 1.0,
        "estimate_ratio": 0.95,
    },
    "high": {
        "quality": 95,
        "target_min_quality": 85,
        "target_min_scale": 1.0,
        "estimate_ratio": 0.92,
    },
    "balanced": {
        "quality": 85,
        "target_min_quality": 75,
        "target_min_scale": 1.0,
        "estimate_ratio": 0.82,
    },
    "aggressive": {
        "quality": 75,
        "target_min_quality": 75,
        "target_min_scale": 0.8,
        "estimate_ratio": 0.72,
    },
}
VIDEO_EXTENSIONS = (
    ".mp4",
    ".m4v",
    ".mov",
    ".wmv",
    ".asf",
    ".avi",
    ".mpg",
    ".mpeg",
    ".mpe",
    ".webm",
    ".mkv",
    ".ts",
    ".m2ts",
    ".3gp",
    ".3g2",
)
MP4_CONTAINER_EXTENSIONS = {".mp4", ".m4v", ".mov"}
IMAGE_EXTENSIONS = (
    ".jpg",
    ".jpeg",
    ".jpe",
    ".png",
    ".webp",
    ".tif",
    ".tiff",
    ".bmp",
    ".gif",
)
JPEG_EXTENSIONS = {".jpg", ".jpeg", ".jpe"}
WEBP_EXTENSIONS = {".webp"}
PNG_EXTENSIONS = {".png"}


def runtime_root() -> Path:
    bundled_root = getattr(sys, "_MEIPASS", None)
    if bundled_root:
        return Path(bundled_root)
    source_root = Path(__file__).resolve().parents[1]
    return source_root if (source_root / "config").is_dir() else Path(sys.prefix)


def runtime_config_dir() -> Path:
    return runtime_root() / "config"


def default_config_path() -> Path:
    return runtime_config_dir() / "default.json"


def profile_config_paths() -> dict[str, Path]:
    config_dir = runtime_config_dir()
    return {
        "balanced": config_dir / "balanced.json",
        "high": config_dir / "high.json",
        "aggressive": config_dir / "aggressive.json",
    }


def binary_filename(name: str) -> str:
    return f"{name}.exe" if os.name == "nt" else name


def resolve_binary(name: str) -> str | None:
    env_names = (
        f"PPTX_TOOLS_{name.upper()}",
        f"PPTX_VIDEO_COMPACTOR_{name.upper()}",
        f"PPTX_OUTPUT_WATERMARK_{name.upper()}",
    )
    for env_name in env_names:
        override = os.environ.get(env_name, "").strip()
        if not override:
            continue
        candidate = Path(override).expanduser()
        if not candidate.exists():
            raise SystemExit(f"{env_name} points to a missing file: {candidate}")
        return str(candidate.resolve())

    filename = binary_filename(name)
    for root in (runtime_root(), Path(sys.executable).resolve().parent):
        for candidate in (
            root / filename,
            root / "assets" / "bin" / filename,
            root / "assets" / "tools" / filename,
            root / "tools" / filename,
        ):
            if candidate.exists():
                return str(candidate)

    return shutil.which(filename) or shutil.which(name)


@dataclass
class Rendition:
    max_height: int
    min_video_kbps: int
    best_video_kbps: int


@dataclass
class VideoOccurrence:
    slide_number: int
    slide_path: str
    shape_name: str
    media_path: str
    x: int
    y: int
    cx: int
    cy: int
    area_ratio: float
    width_ratio: float
    height_ratio: float


@dataclass
class ImageOccurrence:
    owner_path: str
    slide_number: int | None
    media_path: str
    area_ratio: float
    width_ratio: float
    height_ratio: float


@dataclass
class VideoAsset:
    media_path: str
    zip_size: int
    occurrences: list[VideoOccurrence] = field(default_factory=list)
    duration_sec: float = 0.0
    width: int = 0
    height: int = 0
    has_audio: bool = False
    audio_stream_usable: bool = True
    original_video_kbps: int = 0
    original_total_kbps: int = 0
    original_audio_kbps: int = 0
    original_fps: float = 0.0
    original_frame_count: int = 0
    target_fps: float = 0.0
    output_frame_count: int = 0
    display_width_px: int = 0
    display_height_px: int = 0
    max_area_ratio: float = 0.0
    max_width_ratio: float = 0.0
    max_height_ratio: float = 0.0
    min_allowed_height: int = 0
    selected_height: int = 0
    audio_kbps: int = 0
    min_video_kbps: int = 0
    max_video_kbps: int = 0
    target_video_kbps: int = 0
    target_total_kbps: int = 0
    target_bytes: int = 0
    extracted_path: str = ""
    output_path: str = ""
    output_media_path: str = ""
    status: str = "pending"
    source_ssim: float | None = None
    display_ssim: float | None = None
    applied_threshold: float | None = None
    quality_status: str = "not_checked"
    quality_reason: str = ""


@dataclass
class ImageAsset:
    media_path: str
    zip_size: int
    occurrences: list[ImageOccurrence] = field(default_factory=list)
    width: int = 0
    height: int = 0
    display_width_px: int = 0
    display_height_px: int = 0
    max_area_ratio: float = 1.0
    max_width_ratio: float = 1.0
    max_height_ratio: float = 1.0
    content_type: str = "unknown"
    image_format: str = ""
    mode: str = ""
    quality: int = 100
    scale: float = 1.0
    target_bytes: int = 0
    extracted_path: str = ""
    output_path: str = ""
    output_media_path: str = ""
    status: str = "pending"
    reason: str = ""
    source_ssim: float | None = None
    display_ssim: float | None = None
    edge_similarity: float | None = None
    alpha_similarity: float | None = None
    applied_threshold: float | None = None
    quality_status: str = "not_checked"
    quality_reason: str = ""
    metadata_preserved: bool | None = None


@dataclass
class RuntimeConfig:
    render_limits: dict[str, int]
    height_floor_rules: list[dict[str, float | int]]
    audio_rules: dict[str, Any]
    bitrate_ladder: dict[int, Rendition]


Logger = Callable[[str], None]
ProgressCallback = Callable[[int, int, str], None]
CancelCallback = Callable[[], bool]


class CancelledError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compress PPTX/DOCX/PDF/XLSX embedded media, or compress standalone "
            "video/image files using the same profile rules."
        )
    )
    parser.add_argument(
        "input_pptx",
        type=Path,
        nargs="+",
        help="Input PPTX / DOCX / PDF / XLSX / video / image path(s)",
    )
    parser.add_argument(
        "--target-size-mb",
        type=float,
        help="Desired output size in decimal MB (1 MB = 1,000,000 bytes).",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Path to JSON config. Overrides --profile if provided.",
    )
    parser.add_argument(
        "--profile",
        choices=["none", "high", "balanced", "aggressive"],
        default="high",
        help="Built-in video compression profile. Default: high. Use none to skip videos.",
    )
    parser.add_argument(
        "--image-profile",
        choices=["none", "lossless", "high", "balanced", "aggressive"],
        default="high",
        help=(
            "Built-in image compression profile. Default: high. "
            "Lossless only recompresses PNG; none skips images."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Output path. Default: generated beside the input using the target size or preset names.",
    )
    parser.add_argument(
        "--video-output-dir",
        type=Path,
        help=(
            "Directory used when preserving compressed videos. "
            "By default process videos stay in a temp folder and are only copied beside the source on failure."
        ),
    )
    parser.add_argument(
        "--slide-render-width",
        type=int,
        default=1920,
        help="Virtual slide render width used for display-size estimation",
    )
    parser.add_argument(
        "--slide-render-height",
        type=int,
        default=1080,
        help="Virtual slide render height used for display-size estimation",
    )
    parser.add_argument(
        "--min-height",
        type=int,
        default=480,
        help="Smallest allowed encoded height bucket",
    )
    parser.add_argument(
        "--max-height",
        type=int,
        default=1080,
        help="Largest allowed encoded height bucket",
    )
    parser.add_argument(
        "--overscan",
        type=float,
        default=DEFAULT_OVERSCAN,
        help="Extra resolution factor above on-slide display size",
    )
    parser.add_argument(
        "--reserve-mb",
        type=float,
        default=None,
        help="Optional explicit package reserve in decimal MB; default is dynamic.",
    )
    parser.add_argument(
        "--preset",
        default="medium",
        help="ffmpeg x264 preset, for example medium or slow",
    )
    parser.add_argument(
        "--encoder",
        choices=["auto", "cpu", "gpu"],
        default="auto",
        help="Video encoder mode. auto prefers GPU and falls back to CPU; cpu uses libx264 two-pass for better size accuracy.",
    )
    parser.add_argument(
        "--video-ssim-threshold",
        type=float,
        default=0.95,
        help="Base video SSIM quality floor.",
    )
    parser.add_argument(
        "--image-ssim-threshold",
        type=float,
        default=0.99,
        help="Base image SSIM quality floor.",
    )
    parser.add_argument(
        "--quality-mode",
        choices=["safe", "forced"],
        default="safe",
        help="Forced mode is explicit and still enforces absolute quality redlines.",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        help="Working directory. Default: auto temp directory",
    )
    parser.add_argument(
        "--keep-work-dir",
        action="store_true",
        help="Keep extracted and compressed intermediates",
    )
    parser.add_argument(
        "--keep-artifacts",
        action="store_true",
        help="Keep diagnostic sidecar artifacts such as compressed video files and report JSON.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Analyze only and print plan without encoding",
    )
    return parser.parse_args()


def ensure_binary(name: str) -> str:
    resolved = resolve_binary(name)
    if resolved is None:
        raise SystemExit(f"Required binary not found in PATH: {name}")
    return resolved


def hidden_subprocess_kwargs() -> dict[str, Any]:
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startupinfo,
    }


def run(
    cmd: list[str],
    capture: bool = False,
    cancel_callback: CancelCallback | None = None,
    progress_callback: Callable[[float], None] | None = None,
    duration_sec: float = 0.0,
) -> subprocess.CompletedProcess[str]:
    resolved_cmd = list(cmd)
    if resolved_cmd and resolved_cmd[0] in {"ffmpeg", "ffprobe"}:
        binary_path = resolve_binary(resolved_cmd[0])
        if binary_path is None:
            raise SystemExit(f"Required binary not found in PATH: {resolved_cmd[0]}")
        resolved_cmd[0] = binary_path

    import locale

    encoding = locale.getpreferredencoding(False) or "utf-8"
    if encoding.lower() == "ascii":
        encoding = "utf-8"
    if (
        resolved_cmd
        and "ffprobe" in Path(resolved_cmd[0]).name.lower()
        and "-print_format" in resolved_cmd
        and "json" in resolved_cmd
    ):
        encoding = "utf-8"

    if cancel_callback is None:
        return run_process(
            resolved_cmd,
            capture_output=capture,
            check=True,
            text=True,
            encoding=encoding,
            errors="replace",
            **hidden_subprocess_kwargs(),
        )

    capture_stdout = capture or progress_callback is not None
    process = start_process(
        resolved_cmd,
        text=True,
        encoding=encoding,
        errors="replace",
        stdout=subprocess.PIPE if capture_stdout else None,
        stderr=subprocess.PIPE if capture else None,
        **hidden_subprocess_kwargs(),
    )
    try:
        if progress_callback is not None and duration_sec > 0.0 and process.stdout:
            import threading
            import re

            time_re = re.compile(r"out_time_us=(\d+)")

            def read_progress() -> None:
                assert process.stdout is not None
                for line in process.stdout:
                    m = time_re.search(line)
                    if m:
                        us = int(m.group(1))
                        fraction = min(1.0, max(0.0, (us / 1000000.0) / duration_sec))
                        progress_callback(fraction)

            t = threading.Thread(target=read_progress, daemon=True)
            t.start()

        while process.poll() is None:
            if cancel_callback():
                terminate_process(process, grace_seconds=2)
                stdout, stderr = process.communicate()
                raise CancelledError("Cancelled")
            time.sleep(0.2)
        stdout, stderr = process.communicate()
    finally:
        finish_process(process)
    if process.returncode:
        raise subprocess.CalledProcessError(
            process.returncode, resolved_cmd, output=stdout, stderr=stderr
        )
    return subprocess.CompletedProcess(resolved_cmd, process.returncode, stdout, stderr)


@lru_cache(maxsize=1)
def available_encoder_names() -> set[str]:
    try:
        result = run(["ffmpeg", "-hide_banner", "-encoders"], capture=True)
    except (subprocess.CalledProcessError, SystemExit):
        return set()
    names: set[str] = set()
    output = (result.stdout or "") + (result.stderr or "")
    for line in output.splitlines():
        parts = line.split()
        if len(parts) >= 2 and parts[0].startswith("V"):
            names.add(parts[1])
    return names


def gpu_encoder_priority() -> tuple[str, ...]:
    if sys.platform == "darwin":
        return ("h264_videotoolbox",)
    if os.name == "nt":
        return ("h264_nvenc", "h264_qsv", "h264_amf", "h264_mf")
    return GPU_ENCODERS


def probe_gpu_encoder(encoder: str) -> bool:
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel",
        "error",
        "-f",
        "lavfi",
        "-i",
        "testsrc2=size=128x128:rate=24:duration=0.2",
        "-frames:v",
        "1",
        "-c:v",
        encoder,
        "-pix_fmt",
        "yuv420p",
        "-f",
        "null",
        "NUL" if os.name == "nt" else "/dev/null",
    ]
    if encoder == "h264_videotoolbox":
        cmd[-5:-5] = ["-profile:v", "main"]
    try:
        run(cmd, capture=True)
    except (subprocess.CalledProcessError, SystemExit):
        return False
    return True


@lru_cache(maxsize=1)
def usable_gpu_encoder_names() -> tuple[str, ...]:
    available = available_encoder_names()
    usable = []
    for encoder in gpu_encoder_priority():
        if encoder in available and probe_gpu_encoder(encoder):
            usable.append(encoder)
    return tuple(usable)


def select_gpu_encoder() -> str | None:
    usable = usable_gpu_encoder_names()
    return usable[0] if usable else None


def ceil_even(value: float) -> int:
    return max(2, int(math.ceil(value / 2.0) * 2))


def mb_to_bytes(value: float) -> int:
    if not math.isfinite(value) or value <= 0:
        raise ValueError("Target size must be a positive finite decimal MB value")
    return int(value * BYTES_PER_MB)


def is_experimental_runtime() -> bool:
    explicit = os.environ.get("PPTX_TOOLS_EXPERIMENTAL", "").strip().lower()
    return explicit in {"1", "true", "yes", "on"} or (
        bool(getattr(sys, "frozen", False))
        and "experimental" in Path(sys.executable).stem.lower()
    )


def experimental_output_stem(stem: str) -> str:
    if is_experimental_runtime() and not stem.endswith("_experimental"):
        return f"{stem}_experimental"
    return stem


def quality_variant_output_path(path: Path, quality_mode: str) -> Path:
    if quality_mode != "forced":
        return path
    stem = path.stem
    experimental_suffix = "_experimental" if stem.endswith("_experimental") else ""
    if experimental_suffix:
        stem = stem[: -len(experimental_suffix)]
    return path.with_name(f"{stem}_forced{experimental_suffix}{path.suffix}")


def runtime_temp_prefix(prefix: str) -> str:
    return f"{prefix}experimental_" if is_experimental_runtime() else prefix


def dynamic_package_reserve_bytes(target_total_bytes: int, non_media_bytes: int) -> int:
    """Small bounded allowance for ZIP re-pack variance; real output closes the loop."""
    media_budget = max(0, target_total_bytes - non_media_bytes)
    return min(1_000_000, max(64_000, int(media_budget * 0.005)))


def target_report_fields(
    target_size_mb: float | None,
    output_path: Path | None,
) -> dict[str, Any]:
    target_bytes = mb_to_bytes(target_size_mb) if target_size_mb is not None else None
    actual_bytes = (
        output_path.stat().st_size
        if output_path is not None and output_path.is_file()
        else None
    )
    if target_bytes is None:
        status = "not_requested"
        delta_bytes = None
        ratio = None
    elif actual_bytes is None:
        status = "not_measured"
        delta_bytes = None
        ratio = None
    else:
        delta_bytes = actual_bytes - target_bytes
        ratio = actual_bytes / max(1, target_bytes)
        status = "met" if actual_bytes <= target_bytes else "quality_limited"
    return {
        "unit": "decimal_MB",
        "requested_mb": target_size_mb,
        "target_bytes": target_bytes,
        "actual_bytes": actual_bytes,
        "delta_bytes": delta_bytes,
        "target_ratio": round(ratio, 6) if ratio is not None else None,
        "status": status,
    }


def next_target_media_budget(
    *,
    actual_bytes: int,
    target_bytes: int,
    current_media_budget: int,
    maximum_media_budget: int,
    correction_rounds: int,
    giveback_used: bool,
) -> tuple[int, str] | None:
    """Choose the next bounded capacity attempt, if one is still useful."""
    if actual_bytes > target_bytes:
        if correction_rounds >= 2:
            return None
        adjusted = int(current_media_budget * target_bytes / actual_bytes * 0.98)
        adjusted = max(1, min(current_media_budget - 1, adjusted))
        if adjusted >= current_media_budget:
            return None
        return adjusted, "correction"
    if actual_bytes < int(target_bytes * 0.95) and not giveback_used:
        adjusted = int(current_media_budget * target_bytes * 0.975 / actual_bytes)
        adjusted = min(maximum_media_budget, max(current_media_budget + 1, adjusted))
        if adjusted > current_media_budget:
            return adjusted, "quality_giveback"
    return None


def media_plan_signature(
    video_assets: dict[str, VideoAsset], image_assets: dict[str, ImageAsset]
) -> tuple[tuple[object, ...], ...]:
    """Return the encoding choices that can change output bytes."""
    videos = tuple(
        (
            asset.media_path,
            asset.selected_height,
            asset.target_video_kbps,
            asset.audio_kbps,
            round(asset.target_fps, 3),
        )
        for asset in sorted(video_assets.values(), key=lambda item: item.media_path)
    )
    images = tuple(
        (asset.media_path, asset.quality, round(asset.scale, 4))
        for asset in sorted(image_assets.values(), key=lambda item: item.media_path)
    )
    return videos + images


def write_target_skip_report(
    input_path: Path, target_size_mb: float, reason: str
) -> Path:
    report_path = input_path.with_name(f"{input_path.stem}.target.report.json")
    target = target_report_fields(target_size_mb, input_path)
    target["status"] = "source_already_meets"
    report = {
        "input_pptx": str(input_path),
        "output_pptx": str(input_path),
        "target_size_mb": target_size_mb,
        "target": target,
        "presentation": {"skipped": True, "reason": reason},
        "videos": [],
        "images": [],
    }
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_markdown_report(report_path, report)
    return report_path


def parse_fps(value: str | None) -> float:
    if not value or value == "0/0":
        return 0.0
    if "/" not in value:
        return float(value)
    numerator, denominator = value.split("/", 1)
    den = float(denominator)
    if den == 0:
        return 0.0
    return float(numerator) / den


def parse_json_payload(
    payload: str | bytes | bytearray | None, *, source: str
) -> dict[str, Any]:
    if payload is None:
        raise SystemExit(f"{source} did not return JSON output.")
    if isinstance(payload, (bytes, bytearray)):
        payload = payload.decode("utf-8", errors="replace")
    text = str(payload).strip()
    if not text:
        raise SystemExit(f"{source} returned empty JSON output.")
    try:
        data = json.loads(text)
    except TypeError as exc:
        raise SystemExit(f"{source} returned invalid JSON payload.") from exc
    except json.JSONDecodeError as exc:
        raise SystemExit(f"{source} returned invalid JSON: {exc.msg}") from exc
    if not isinstance(data, dict):
        raise SystemExit(f"{source} returned an unexpected JSON structure.")
    return data


def load_json_file(path: Path, *, source: str) -> dict[str, Any]:
    try:
        payload = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise SystemExit(f"{source} JSON file could not be read: {path}") from exc
    return parse_json_payload(payload, source=f"{source} ({path})")


def load_runtime_config(config_path: Path | None, profile: str) -> RuntimeConfig:
    if config_path:
        resolved = config_path.expanduser().resolve()
    else:
        resolved = profile_config_paths().get(profile, default_config_path()).resolve()
    data = load_json_file(resolved, source="Runtime config")
    ladder = {
        int(height): Rendition(
            max_height=int(cfg["max_height"]),
            min_video_kbps=int(cfg["min_video_kbps"]),
            best_video_kbps=int(cfg.get("best_video_kbps", cfg.get("max_video_kbps"))),
        )
        for height, cfg in data["bitrate_ladder"].items()
    }
    return RuntimeConfig(
        render_limits={
            "max_output_height": int(data["render_limits"]["max_output_height"]),
            "max_long_edge": int(data["render_limits"]["max_long_edge"]),
        },
        height_floor_rules=list(data["height_floor_rules"]),
        audio_rules=dict(data["audio_rules"]),
        bitrate_ladder=ladder,
    )


def normalize_bucket(height: int, mode: str) -> int:
    if mode == "floor":
        candidates = [h for h in STD_HEIGHTS if h >= height]
        return candidates[0] if candidates else STD_HEIGHTS[-1]
    candidates = [h for h in STD_HEIGHTS if h <= height]
    return candidates[-1] if candidates else STD_HEIGHTS[0]


def pick_bucket(height_px: float, min_height: int, max_height: int) -> int:
    min_height = normalize_bucket(max(min_height, STD_HEIGHTS[0]), "floor")
    max_height = normalize_bucket(min(max_height, STD_HEIGHTS[-1]), "ceil")
    if min_height > max_height:
        min_height = max_height
    allowed = [h for h in STD_HEIGHTS if min_height <= h <= max_height]
    for bucket in allowed:
        if height_px <= bucket:
            return bucket
    return allowed[-1]


def lower_bucket(bucket: int, min_height: int) -> int:
    min_height = normalize_bucket(max(min_height, STD_HEIGHTS[0]), "floor")
    if min_height > bucket:
        min_height = bucket
    allowed = [h for h in STD_HEIGHTS if min_height <= h <= bucket]
    idx = allowed.index(bucket)
    return allowed[max(0, idx - 1)]


def bitrate_bucket_for_asset(asset: VideoAsset) -> int:
    return asset.selected_height


def rendition_for_asset(asset: VideoAsset, config: RuntimeConfig) -> Rendition:
    return config.bitrate_ladder[bitrate_bucket_for_asset(asset)]


def apply_video_bitrate_bounds(asset: VideoAsset, rendition: Rendition) -> None:
    source_cap = (
        asset.original_video_kbps
        if asset.original_video_kbps > 1
        else rendition.best_video_kbps
    )
    asset.min_video_kbps = min(rendition.min_video_kbps, source_cap)
    asset.max_video_kbps = min(rendition.best_video_kbps, source_cap)
    asset.max_video_kbps = max(asset.max_video_kbps, asset.min_video_kbps)


def hard_min_height_for_asset(asset: VideoAsset) -> int:
    if (
        asset.max_area_ratio >= 0.45
        or asset.max_width_ratio >= 0.5
        or asset.max_height_ratio >= 0.5
    ):
        return 720
    return STD_HEIGHTS[0]


def ratio_rule_matches(
    area_ratio: float,
    width_ratio: float,
    height_ratio: float,
    rule: dict[str, Any],
) -> bool:
    min_area = float(rule.get("min_area_ratio", 0.0))
    min_width = float(rule.get("min_width_ratio", 0.0))
    min_height = float(rule.get("min_height_ratio", 0.0))

    if min_area > 0.0 and area_ratio >= min_area:
        return True

    if (min_width > 0.0 and width_ratio >= min_width) or (
        min_height > 0.0 and height_ratio >= min_height
    ):
        return True

    return min_area <= 0.0 and min_width <= 0.0 and min_height <= 0.0


def recommend_height_floor(
    area_ratio: float,
    width_ratio: float,
    height_ratio: float,
    config: RuntimeConfig,
) -> int:
    for rule in config.height_floor_rules:
        if ratio_rule_matches(area_ratio, width_ratio, height_ratio, rule):
            return int(rule["min_height"])
    return STD_HEIGHTS[0]


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def recommend_audio_kbps(
    asset: VideoAsset,
    compression_ratio: float,
    config: RuntimeConfig,
) -> int:
    if not asset.has_audio:
        return 0

    target = int(config.audio_rules["min_kbps"])
    for tier in config.audio_rules["tiers"]:
        if ratio_rule_matches(
            asset.max_area_ratio,
            asset.max_width_ratio,
            asset.max_height_ratio,
            tier,
        ):
            target = int(tier["target_kbps"])
            break

    for threshold in config.audio_rules["pressure_adjustments"]:
        if compression_ratio < float(threshold["below_ratio"]):
            target -= int(threshold["reduce_by_kbps"])

    min_audio = int(config.audio_rules["min_kbps"])
    max_audio = int(config.audio_rules["max_kbps"])
    if asset.original_audio_kbps > 0:
        step = int(config.audio_rules["rounding_step_kbps"])
        rounded_original = max(1, int(asset.original_audio_kbps / step) * step)
        max_audio = min(max_audio, rounded_original)
        if max_audio < min_audio:
            return max_audio

    return min(max_audio, max(min_audio, target))


def asset_priority_score(
    asset: VideoAsset,
    median_size_bytes: float,
    median_video_kbps: float,
) -> float:
    area_score = max(0.04, asset.max_area_ratio) ** 0.6
    size_score = clamp(
        math.sqrt(asset.zip_size / max(1.0, median_size_bytes)), 0.7, 1.6
    )
    bitrate_score = clamp(
        math.sqrt(asset.original_video_kbps / max(1.0, median_video_kbps)),
        0.7,
        1.6,
    )
    return area_score * size_score * bitrate_score * asset.duration_sec


def resolve_zip_target(source_xml: str, target: str) -> str:
    base_dir = posixpath.dirname(source_xml)
    target = target.replace("\\", "/")
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def rel_target_to_zip_path(source_xml: str, target_zip_path: str) -> str:
    base_dir = posixpath.dirname(source_xml)
    return posixpath.relpath(target_zip_path, base_dir)


def source_path_from_rels_path(rels_path: str) -> str | None:
    marker = "/_rels/"
    if marker not in rels_path or not rels_path.endswith(".rels"):
        return None
    base_dir, rel_name = rels_path.split(marker, 1)
    source_name = rel_name.removesuffix(".rels")
    return posixpath.join(base_dir, source_name)


def media_needs_mp4(media_path: str) -> bool:
    return Path(media_path).suffix.lower() not in MP4_CONTAINER_EXTENSIONS


def media_is_supported_image(media_path: str) -> bool:
    return Path(media_path).suffix.lower() in IMAGE_EXTENSIONS


def zip_member_size(info: ZipInfo) -> int:
    return int(info.compress_size or info.file_size)


def choose_output_media_path(media_path: str, reserved_paths: set[str]) -> str:
    if not media_needs_mp4(media_path):
        return media_path

    stem, _ = posixpath.splitext(media_path)
    candidate = f"{stem}.mp4"
    if candidate not in reserved_paths:
        return candidate

    index = 1
    while True:
        candidate = f"{stem}_compact{index}.mp4"
        if candidate not in reserved_paths:
            return candidate
        index += 1


def parse_relationships(zf: ZipFile, slide_xml_path: str) -> dict[str, str]:
    slide_name = posixpath.basename(slide_xml_path)
    rel_path = f"ppt/slides/_rels/{slide_name}.rels"
    if rel_path not in zf.namelist():
        return {}
    root = SafeET.fromstring(zf.read(rel_path))
    rels: dict[str, str] = {}
    for rel in root.findall("pr:Relationship", REL_NS):
        rel_id = rel.attrib["Id"]
        rels[rel_id] = resolve_zip_target(slide_xml_path, rel.attrib["Target"])
    return rels


def referenced_ooxml_images(
    zf: ZipFile, owner_prefix: str, media_prefix: str
) -> dict[str, set[str]]:
    members = set(zf.namelist())
    references: dict[str, set[str]] = {}
    for rels_path in members:
        owner = source_path_from_rels_path(rels_path)
        if owner is None or owner not in members or not owner.startswith(owner_prefix):
            continue
        try:
            root = SafeET.fromstring(zf.read(rels_path))
        except (ET.ParseError, KeyError):
            continue
        for relationship in root.findall("pr:Relationship", REL_NS):
            if relationship.attrib.get("TargetMode") == "External":
                continue
            if not relationship.attrib.get("Type", "").endswith("/image"):
                continue
            target = relationship.attrib.get("Target", "")
            resolved = resolve_zip_target(owner, target)
            if resolved.startswith(media_prefix) and resolved in members:
                references.setdefault(resolved, set()).add(owner)
    return references


def referenced_pptx_images(zf: ZipFile) -> dict[str, set[str]]:
    return referenced_ooxml_images(zf, "ppt/", "ppt/media/")


def slide_image_occurrences(
    zf: ZipFile,
    slide_path: str,
    slide_cx: int,
    slide_cy: int,
) -> list[ImageOccurrence]:
    rels = parse_relationships(zf, slide_path)
    root = SafeET.fromstring(zf.read(slide_path))
    slide_number = int(Path(slide_path).stem.replace("slide", ""))
    occurrences: list[ImageOccurrence] = []
    for pic in root.findall(".//p:pic", NS):
        blip = pic.find(".//a:blip", NS)
        transform = pic.find("./p:spPr/a:xfrm", NS)
        extent = transform.find("a:ext", NS) if transform is not None else None
        if blip is None or extent is None:
            continue
        rel_id = blip.attrib.get(f"{R_NS}embed")
        media_path = rels.get(rel_id or "")
        if not media_path or not media_is_supported_image(media_path):
            continue
        width_ratio = int(extent.attrib.get("cx", 0)) / max(1, slide_cx)
        height_ratio = int(extent.attrib.get("cy", 0)) / max(1, slide_cy)
        occurrences.append(
            ImageOccurrence(
                owner_path=slide_path,
                slide_number=slide_number,
                media_path=media_path,
                area_ratio=width_ratio * height_ratio,
                width_ratio=width_ratio,
                height_ratio=height_ratio,
            )
        )
    return occurrences


def parse_pptx_assets(
    pptx_path: Path,
    render_width: int,
    render_height: int,
    overscan: float,
    min_height: int,
    max_height: int,
    config: RuntimeConfig,
    include_videos: bool = True,
) -> tuple[dict[str, VideoAsset], dict[str, ImageAsset], dict[str, Any]]:
    with ZipFile(pptx_path) as zf:
        presentation = SafeET.fromstring(zf.read("ppt/presentation.xml"))
        slide_size = presentation.find("p:sldSz", NS)
        if slide_size is None:
            raise SystemExit("Could not read slide size from presentation.xml")
        slide_cx = int(slide_size.attrib["cx"])
        slide_cy = int(slide_size.attrib["cy"])
        assets: dict[str, VideoAsset] = {}
        slide_paths = (
            sorted(
                (
                    name
                    for name in zf.namelist()
                    if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                ),
                key=lambda name: int(Path(name).stem.replace("slide", "")),
            )
            if include_videos
            else []
        )

        for slide_path in slide_paths:
            slide_number = int(Path(slide_path).stem.replace("slide", ""))
            rels = parse_relationships(zf, slide_path)
            root = SafeET.fromstring(zf.read(slide_path))
            for pic in root.findall(".//p:pic", NS):
                nv_pic = pic.find("./p:nvPicPr", NS)
                sp_pr = pic.find("./p:spPr/a:xfrm", NS)
                if nv_pic is None or sp_pr is None:
                    continue
                nv_pr = nv_pic.find("./p:nvPr", NS)
                c_nv_pr = nv_pic.find("./p:cNvPr", NS)
                if nv_pr is None or c_nv_pr is None:
                    continue

                rel_ids: list[str] = []
                video_file = nv_pr.find("a:videoFile", NS)
                if video_file is not None:
                    rel_id = video_file.attrib.get(f"{R_NS}link")
                    if rel_id:
                        rel_ids.append(rel_id)
                media_ref = nv_pr.find(".//p14:media", NS)
                if media_ref is not None:
                    rel_id = media_ref.attrib.get(f"{R_NS}embed")
                    if rel_id:
                        rel_ids.append(rel_id)
                if not rel_ids:
                    continue

                media_path = None
                for rel_id in rel_ids:
                    candidate = rels.get(rel_id)
                    if candidate and candidate.lower().endswith(VIDEO_EXTENSIONS):
                        media_path = candidate
                        break
                if not media_path:
                    continue

                ext = sp_pr.find("a:ext", NS)
                off = sp_pr.find("a:off", NS)
                if ext is None or off is None:
                    continue

                cx = int(ext.attrib["cx"])
                cy = int(ext.attrib["cy"])
                width_ratio = cx / slide_cx
                height_ratio = cy / slide_cy
                area_ratio = width_ratio * height_ratio
                occurrence = VideoOccurrence(
                    slide_number=slide_number,
                    slide_path=slide_path,
                    shape_name=c_nv_pr.attrib.get("name", f"slide{slide_number}_video"),
                    media_path=media_path,
                    x=int(off.attrib["x"]),
                    y=int(off.attrib["y"]),
                    cx=cx,
                    cy=cy,
                    area_ratio=area_ratio,
                    width_ratio=width_ratio,
                    height_ratio=height_ratio,
                )

                if media_path not in assets:
                    info = zf.getinfo(media_path)
                    assets[media_path] = VideoAsset(
                        media_path=media_path, zip_size=zip_member_size(info)
                    )
                assets[media_path].occurrences.append(occurrence)

        reserved_paths = set(zf.namelist())
        for asset in assets.values():
            asset.output_media_path = choose_output_media_path(
                asset.media_path, reserved_paths
            )
            reserved_paths.add(asset.output_media_path)

            max_occ = max(asset.occurrences, key=lambda item: item.area_ratio)
            asset.max_area_ratio = max_occ.area_ratio
            asset.max_width_ratio = max_occ.width_ratio
            asset.max_height_ratio = max_occ.height_ratio
            asset.display_width_px = min(
                render_width, int(render_width * max_occ.width_ratio * overscan)
            )
            asset.display_height_px = min(
                render_height, int(render_height * max_occ.height_ratio * overscan)
            )
            asset.min_allowed_height = recommend_height_floor(
                area_ratio=asset.max_area_ratio,
                width_ratio=asset.max_width_ratio,
                height_ratio=asset.max_height_ratio,
                config=config,
            )
            asset.selected_height = pick_bucket(
                asset.display_height_px,
                max(min_height, asset.min_allowed_height),
                max_height,
            )

        image_references = referenced_pptx_images(zf)
        image_occurrences: dict[str, list[ImageOccurrence]] = {}
        for slide_path in slide_paths or sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        ):
            for occurrence in slide_image_occurrences(
                zf, slide_path, slide_cx, slide_cy
            ):
                image_occurrences.setdefault(occurrence.media_path, []).append(
                    occurrence
                )

        image_assets: dict[str, ImageAsset] = {}
        all_media_images = {
            name
            for name in zf.namelist()
            if name.startswith("ppt/media/")
            and name not in assets
            and media_is_supported_image(name)
        }
        for name in sorted(image_references):
            if name not in all_media_images:
                continue
            info = zf.getinfo(name)
            occurrences = image_occurrences.get(name, [])
            asset = ImageAsset(
                media_path=name,
                zip_size=zip_member_size(info),
                occurrences=occurrences,
            )
            if occurrences:
                largest = max(occurrences, key=lambda item: item.area_ratio)
                asset.max_area_ratio = largest.area_ratio
                asset.max_width_ratio = largest.width_ratio
                asset.max_height_ratio = largest.height_ratio
            # Images referenced outside slides (masters, layouts, notes, charts)
            # stay conservative until their exact rendered geometry is known.
            asset.display_width_px = max(
                1, int(render_width * asset.max_width_ratio * overscan)
            )
            asset.display_height_px = max(
                1, int(render_height * asset.max_height_ratio * overscan)
            )
            image_assets[name] = asset

        orphan_image_paths = sorted(all_media_images - set(image_references))

        return (
            assets,
            image_assets,
            {
                "slide_cx": slide_cx,
                "slide_cy": slide_cy,
                "render_width": render_width,
                "render_height": render_height,
                "overscan": overscan,
                "orphan_image_paths": orphan_image_paths,
            },
        )


def ffprobe_json(path: Path) -> dict[str, Any]:
    result = run(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_streams",
            "-show_format",
            "-print_format",
            "json",
            str(path),
        ],
        capture=True,
    )
    if not result.stdout:
        stderr_msg = result.stderr or "No stderr output"
        raise SystemExit(
            f"ffprobe ({path}) returned empty JSON output. "
            f"Returncode: {result.returncode}. Stderr: {stderr_msg}"
        )
    return parse_json_payload(result.stdout, source=f"ffprobe ({path})")


def stream_frame_count(stream: dict[str, Any]) -> int:
    try:
        return max(0, int(stream.get("nb_frames") or 0))
    except (TypeError, ValueError):
        return 0


def audio_stream_is_usable(stream: dict[str, Any]) -> bool:
    """Return False only when ffprobe explicitly proves an empty audio track."""
    duration_known = stream.get("duration") not in (None, "", "N/A")
    frames_known = stream.get("nb_frames") not in (None, "", "N/A")
    if not duration_known or not frames_known:
        return True

    try:
        duration = float(stream.get("duration") or 0.0)
        frames = stream_frame_count(stream)
    except (TypeError, ValueError):
        return True
    return duration > 0.0 or frames > 1


def reduces_frame_rate(asset: VideoAsset) -> bool:
    return asset.target_fps > 0 and asset.original_fps > asset.target_fps + 0.5


def append_frame_rate_mode(cmd: list[str], asset: VideoAsset) -> None:
    if not reduces_frame_rate(asset):
        # Preserve VFR timestamps instead of letting FFmpeg silently drop frames
        # when the source average FPS differs from its nominal stream rate.
        cmd.extend(["-fps_mode", "passthrough"])


def validate_encoded_asset(
    asset: VideoAsset,
    target: Path,
    expected_width: int,
    expected_height: int,
) -> None:
    probe = ffprobe_json(target)
    streams = probe.get("streams", [])
    video = next((item for item in streams if item.get("codec_type") == "video"), None)
    audio = next((item for item in streams if item.get("codec_type") == "audio"), None)
    if video is None:
        raise ValueError(f"Encoded video stream is missing: {asset.media_path}")
    if (int(video.get("width") or 0), int(video.get("height") or 0)) != (
        expected_width,
        expected_height,
    ):
        raise ValueError(f"Encoded video resolution mismatch: {asset.media_path}")
    if (audio is not None) != asset.has_audio:
        raise ValueError(f"Encoded video audio stream mismatch: {asset.media_path}")

    duration = float(
        video.get("duration") or probe.get("format", {}).get("duration") or 0
    )
    fps = asset.target_fps or asset.original_fps
    tolerance = max(0.1, 2.0 / fps) if fps > 0 else 0.1
    if duration <= 0 or abs(duration - asset.duration_sec) > tolerance:
        raise ValueError(f"Encoded video duration mismatch: {asset.media_path}")

    asset.output_frame_count = stream_frame_count(video)
    if not asset.output_frame_count:
        return
    if reduces_frame_rate(asset):
        expected_frames = round(asset.duration_sec * asset.target_fps)
        if abs(asset.output_frame_count - expected_frames) > 1:
            raise ValueError(f"Encoded video frame count mismatch: {asset.media_path}")
    elif (
        asset.original_frame_count
        and abs(asset.output_frame_count - asset.original_frame_count) > 1
    ):
        raise ValueError(f"Encoded video frame count mismatch: {asset.media_path}")


def extract_videos(
    pptx_path: Path,
    assets: dict[str, VideoAsset],
    work_dir: Path,
    progress_callback: ProgressCallback | None = None,
) -> None:
    originals_dir = work_dir / "original_media"
    originals_dir.mkdir(parents=True, exist_ok=True)
    with ZipFile(pptx_path) as zf:
        asset_list = list(assets.values())
        for index, asset in enumerate(asset_list, start=1):
            if progress_callback is not None:
                progress_callback(
                    index, len(asset_list), f"正在提取视频 {index}/{len(asset_list)}"
                )
            target = originals_dir / Path(asset.media_path).name
            with zf.open(asset.media_path) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output, length=1024 * 1024)
            asset.extracted_path = str(target)
            probe = ffprobe_json(target)
            streams = probe.get("streams", [])
            video_stream = next(
                (s for s in streams if s.get("codec_type") == "video"), None
            )
            audio_stream = next(
                (s for s in streams if s.get("codec_type") == "audio"), None
            )
            if video_stream is None:
                raise SystemExit(f"No video stream found in {asset.media_path}")
            asset.width = int(video_stream.get("width", 0))
            asset.height = int(video_stream.get("height", 0))
            duration = video_stream.get("duration") or probe.get("format", {}).get(
                "duration"
            )
            asset.duration_sec = max(0.1, float(duration or 0.1))
            format_bit_rate = int(
                float(probe.get("format", {}).get("bit_rate", 0) or 0)
            )
            video_bit_rate = int(float(video_stream.get("bit_rate", 0) or 0))
            asset.original_total_kbps = (
                max(1, format_bit_rate // 1000)
                if format_bit_rate
                else max(1, int(asset.zip_size * 8 / asset.duration_sec / 1000))
            )
            asset.original_video_kbps = max(
                1,
                video_bit_rate // 1000
                if video_bit_rate
                else max(asset.original_total_kbps - 96, 1),
            )
            asset.original_fps = parse_fps(
                video_stream.get("avg_frame_rate") or video_stream.get("r_frame_rate")
            )
            asset.original_frame_count = stream_frame_count(video_stream)
            asset.has_audio = audio_stream is not None
            asset.audio_stream_usable = audio_stream is None or audio_stream_is_usable(
                audio_stream
            )
            if not asset.audio_stream_usable:
                # Do not transcode a declared-but-empty track: FFmpeg may silently
                # omit it, which would violate the no-audio-loss contract.
                asset.output_media_path = asset.media_path
                asset.quality_reason = "unusable_audio_stream_preserved"
            asset.original_audio_kbps = (
                int(float(audio_stream.get("bit_rate", 0) or 0)) // 1000
                if audio_stream is not None
                else 0
            )


def extract_images(
    pptx_path: Path,
    assets: dict[str, ImageAsset],
    work_dir: Path,
    progress_callback: ProgressCallback | None = None,
) -> None:
    originals_dir = work_dir / "original_images"
    originals_dir.mkdir(parents=True, exist_ok=True)
    with ZipFile(pptx_path) as zf:
        asset_list = list(assets.values())
        for index, asset in enumerate(asset_list, start=1):
            if progress_callback is not None:
                progress_callback(
                    index, len(asset_list), f"正在提取图片 {index}/{len(asset_list)}"
                )
            target = originals_dir / Path(asset.media_path).name
            with zf.open(asset.media_path) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output, length=1024 * 1024)
            asset.extracted_path = str(target)
            try:
                with Image.open(target) as image:
                    asset.width = int(image.width)
                    asset.height = int(image.height)
                    asset.image_format = str(image.format or "").upper()
                    asset.mode = image.mode
                    asset.content_type = classify_image_content(image)
            except (UnidentifiedImageError, OSError) as exc:
                asset.status = "unsupported"
                asset.reason = f"Cannot read image: {exc}"


def consolidate_exact_duplicate_images(
    assets: dict[str, ImageAsset],
) -> dict[str, str]:
    canonical_by_identity: dict[tuple[str, str], ImageAsset] = {}
    duplicate_map: dict[str, str] = {}
    for media_path, asset in list(assets.items()):
        source = Path(asset.extracted_path)
        if not source.is_file():
            continue
        with source.open("rb") as handle:
            digest = hashlib.file_digest(handle, "sha256").hexdigest()
        identity = (digest, source.suffix.lower())
        canonical = canonical_by_identity.get(identity)
        if canonical is None:
            canonical_by_identity[identity] = asset
            continue
        duplicate_map[media_path] = canonical.media_path
        canonical.occurrences.extend(asset.occurrences)
        canonical.max_area_ratio = max(canonical.max_area_ratio, asset.max_area_ratio)
        canonical.max_width_ratio = max(
            canonical.max_width_ratio, asset.max_width_ratio
        )
        canonical.max_height_ratio = max(
            canonical.max_height_ratio, asset.max_height_ratio
        )
        canonical.display_width_px = max(
            canonical.display_width_px, asset.display_width_px
        )
        canonical.display_height_px = max(
            canonical.display_height_px, asset.display_height_px
        )
        del assets[media_path]
    return duplicate_map


def classify_image_content(image: Image.Image) -> str:
    if "A" in image.getbands() or "transparency" in image.info:
        return "transparent"
    sample = image.convert("RGB")
    sample.thumbnail((256, 256), Image.Resampling.LANCZOS)
    colors = sample.getcolors(maxcolors=513)
    if colors is not None and len(colors) <= 32:
        return "line_art"
    gray = sample.convert("L")
    edges = gray.filter(ImageFilter.FIND_EDGES)
    edge_mean = float(ImageStat.Stat(edges).mean[0]) / 255.0
    if edge_mean >= 0.16:
        return "screenshot_or_text"
    return "photo"


def image_importance(asset: ImageAsset) -> float:
    area = clamp(asset.max_area_ratio, 0.01, 1.0)
    reuse_bonus = min(0.15, math.log2(max(1, len(asset.occurrences))) * 0.05)
    content_bonus = (
        0.1
        if asset.content_type
        in {
            "line_art",
            "screenshot_or_text",
            "transparent",
        }
        else 0.0
    )
    return clamp(
        0.65 + math.sqrt(area) * 0.35 + reuse_bonus + content_bonus, 0.65, 1.25
    )


def assign_image_plan(
    assets: dict[str, ImageAsset],
    profile: str,
    target_image_bytes: int | None = None,
    *,
    preserve_quality_fallbacks: bool = False,
) -> None:
    rules = IMAGE_QUALITY_RULES.get(profile, IMAGE_QUALITY_RULES["high"])
    base_quality = int(rules["quality"])
    min_quality = int(rules["target_min_quality"])
    min_scale = float(rules["target_min_scale"])
    pressure = 1.0
    current_image_bytes = sum(asset.zip_size for asset in assets.values())
    if target_image_bytes is not None and current_image_bytes > 0:
        pressure = clamp(target_image_bytes / current_image_bytes, 0.0, 1.0)

    for asset in assets.values():
        if preserve_quality_fallbacks and asset.quality_status == "restored_original":
            asset.quality = 100
            asset.scale = 1.0
            asset.target_bytes = asset.zip_size
            asset.status = "copy_requested"
            continue
        if profile == "none":
            asset.quality = 100
            asset.scale = 1.0
            asset.target_bytes = asset.zip_size
            asset.status = "copy_requested"
            continue
        if (
            profile == "lossless"
            and Path(asset.media_path).suffix.lower() not in PNG_EXTENSIONS
        ):
            asset.quality = 100
            asset.scale = 1.0
            asset.target_bytes = asset.zip_size
            asset.status = "copy_requested"
            continue
        if asset.status == "unsupported":
            asset.quality = 100
            asset.scale = 1.0
            asset.target_bytes = asset.zip_size
            continue

        if target_image_bytes is None or pressure >= 1.0:
            quality = base_quality
            scale = 1.0
        else:
            importance = image_importance(asset)
            effective_pressure = clamp(1.0 - ((1.0 - pressure) / importance), 0.0, 1.0)
            quality_factor = clamp((effective_pressure - 0.58) / 0.42, 0.0, 1.0)
            quality = int(
                round(min_quality + (base_quality - min_quality) * quality_factor)
            )
            if profile == "aggressive":
                scale_factor = clamp((pressure - 0.70) / 0.30, 0.0, 1.0)
                scale = min_scale + (1.0 - min_scale) * scale_factor
            else:
                scale = 1.0

        asset.quality = max(1, min(100, quality))
        asset.scale = max(min_scale, min(1.0, scale))
        estimate_ratio = float(rules["estimate_ratio"])
        scale_ratio = asset.scale * asset.scale
        asset.target_bytes = int(asset.zip_size * estimate_ratio * scale_ratio)
        asset.status = "planned"


def allocate_media_budgets(
    current_video_bytes: int,
    current_image_bytes: int,
    target_media_bytes: int,
    video_profile: str,
    image_profile: str,
) -> tuple[int, int]:
    """Share required savings across media types instead of budgeting serially."""
    current_total = current_video_bytes + current_image_bytes
    if target_media_bytes >= current_total:
        return current_video_bytes, current_image_bytes

    video_floor = (
        current_video_bytes
        if video_profile == "none"
        else int(
            current_video_bytes
            * float(PROFILE_QUALITY_RULES[video_profile]["estimate_ratio"])
        )
    )
    image_floor = (
        current_image_bytes
        if image_profile == "none"
        else int(
            current_image_bytes
            * float(IMAGE_QUALITY_RULES[image_profile]["estimate_ratio"])
        )
    )
    video_headroom = max(0, current_video_bytes - video_floor)
    image_headroom = max(0, current_image_bytes - image_floor)
    savings_needed = max(0, current_total - target_media_bytes)
    initial_headroom = video_headroom + image_headroom

    if initial_headroom:
        initial_savings = min(savings_needed, initial_headroom)
        video_savings = int(initial_savings * video_headroom / initial_headroom)
        image_savings = initial_savings - video_savings
    else:
        video_savings = image_savings = 0

    remaining = savings_needed - video_savings - image_savings
    compressible_video = (
        max(0, current_video_bytes - video_savings) if video_profile != "none" else 0
    )
    compressible_image = (
        max(0, current_image_bytes - image_savings) if image_profile != "none" else 0
    )
    compressible_total = compressible_video + compressible_image
    if remaining > 0 and compressible_total > 0:
        extra_video = int(remaining * compressible_video / compressible_total)
        video_savings += min(compressible_video, extra_video)
        image_savings += min(
            compressible_image, remaining - min(compressible_video, extra_video)
        )

    return (
        max(0, current_video_bytes - video_savings),
        max(0, current_image_bytes - image_savings),
    )


def applied_quality_threshold(
    base: float, area_ratio: float, reuse_count: int, redline: float
) -> float:
    area_adjustment = (math.sqrt(clamp(area_ratio, 0.01, 1.0)) - 0.5) * 0.04
    reuse_adjustment = min(0.015, math.log2(max(1, reuse_count)) * 0.005)
    return round(clamp(base + area_adjustment + reuse_adjustment, redline, 0.995), 4)


def measure_media_ssim(
    reference: Path,
    candidate: Path,
    *,
    is_video: bool,
    width: int,
    height: int,
) -> float:
    ffmpeg = ensure_binary("ffmpeg")
    width = ceil_even(width)
    height = ceil_even(height)
    if is_video:
        filter_cmd = (
            f"[0:v]setpts=PTS-STARTPTS,fps=1,scale={width}:{height}[candidate];"
            f"[1:v]setpts=PTS-STARTPTS,fps=1,scale={width}:{height}[reference];"
            "[candidate][reference]ssim"
        )
    else:
        filter_cmd = (
            f"[0:v]scale={width}:{height}[candidate];"
            f"[1:v]scale={width}:{height}[reference];"
            "[candidate][reference]ssim"
        )
    result = run(
        [
            ffmpeg,
            "-i",
            str(candidate),
            "-i",
            str(reference),
            "-lavfi",
            filter_cmd,
            "-f",
            "null",
            "-",
        ],
        capture=True,
    )
    match = re.search(r"All:([0-9.]+)", result.stderr or "")
    if match is None:
        raise ValueError("FFmpeg did not return an SSIM score")
    return float(match.group(1))


def image_detail_metrics(reference: Path, candidate: Path) -> tuple[float, float]:
    with (
        Image.open(reference) as reference_image,
        Image.open(candidate) as candidate_image,
    ):
        reference_rgba = reference_image.convert("RGBA")
        candidate_rgba = candidate_image.convert("RGBA").resize(
            reference_rgba.size, Image.Resampling.LANCZOS
        )
        reference_edge = reference_rgba.convert("L").filter(ImageFilter.FIND_EDGES)
        candidate_edge = candidate_rgba.convert("L").filter(ImageFilter.FIND_EDGES)
        edge_difference = float(
            ImageStat.Stat(ImageChops.difference(reference_edge, candidate_edge)).mean[
                0
            ]
        )
        alpha_difference = float(
            ImageStat.Stat(
                ImageChops.difference(
                    reference_rgba.getchannel("A"), candidate_rgba.getchannel("A")
                )
            ).mean[0]
        )
    return max(0.0, 1.0 - edge_difference / 255.0), max(
        0.0, 1.0 - alpha_difference / 255.0
    )


def images_pixel_identical(reference: Path, candidate: Path) -> bool:
    try:
        with Image.open(reference) as first, Image.open(candidate) as second:
            if first.size != second.size:
                return False
            extrema = ImageChops.difference(
                first.convert("RGBA"), second.convert("RGBA")
            ).getextrema()
            return all(low == high == 0 for low, high in extrema)
    except (UnidentifiedImageError, OSError, ValueError):
        return False


def image_metadata_signature(path: Path) -> tuple[Any, ...]:
    with Image.open(path) as image:
        dpi = image.info.get("dpi")
        normalized_dpi = (
            tuple(round(float(value), 2) for value in dpi)
            if isinstance(dpi, tuple)
            else dpi
        )
        text = tuple(sorted(getattr(image, "text", {}).items()))
        return (
            image.format,
            int(getattr(image, "n_frames", 1)),
            bool(getattr(image, "is_animated", False)),
            image.info.get("icc_profile"),
            image.info.get("exif") or bytes(image.getexif()),
            image.info.get("xmp"),
            normalized_dpi,
            image.info.get("transparency"),
            image.info.get("loop"),
            image.info.get("duration"),
            text,
        )


def image_quality_passes(asset: ImageAsset, *, preserve_image_metadata: bool) -> bool:
    reference = Path(asset.extracted_path)
    candidate = Path(asset.output_path)
    try:
        if images_pixel_identical(reference, candidate):
            asset.source_ssim = asset.display_ssim = 1.0
            asset.edge_similarity = asset.alpha_similarity = 1.0
        else:
            asset.source_ssim = measure_media_ssim(
                reference,
                candidate,
                is_video=False,
                width=asset.width,
                height=asset.height,
            )
            asset.display_ssim = measure_media_ssim(
                reference,
                candidate,
                is_video=False,
                width=asset.display_width_px or asset.width,
                height=asset.display_height_px or asset.height,
            )
            asset.edge_similarity, asset.alpha_similarity = image_detail_metrics(
                reference, candidate
            )
        edge_floor = (
            0.98 if asset.content_type in {"line_art", "screenshot_or_text"} else 0.90
        )
        if (
            min(asset.source_ssim, asset.display_ssim) < asset.applied_threshold
            or asset.edge_similarity < edge_floor
            or asset.alpha_similarity < 0.995
        ):
            asset.quality_reason = "image_quality_below_threshold"
            return False
        asset.metadata_preserved = not preserve_image_metadata or (
            image_metadata_signature(reference) == image_metadata_signature(candidate)
        )
        if not asset.metadata_preserved:
            asset.quality_reason = "image_metadata_changed"
            return False
        asset.quality_status = "passed"
        asset.quality_reason = ""
        return True
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        asset.quality_reason = f"image_quality_unavailable: {exc}"
        return False


def audit_encoded_assets(
    video_assets: dict[str, VideoAsset],
    image_assets: dict[str, ImageAsset],
    *,
    video_threshold: float,
    image_threshold: float,
    forced: bool = False,
    preserve_image_metadata: bool = True,
    logger: Logger | None = None,
    restore_failed: bool = True,
) -> None:
    for asset in video_assets.values():
        if asset.status not in {"encoded", "encoded_gpu", "reused"}:
            asset.quality_status = "unchanged"
            continue
        reference = Path(asset.extracted_path)
        candidate = Path(asset.output_path)
        asset.applied_threshold = applied_quality_threshold(
            0.90 if forced else video_threshold,
            asset.max_area_ratio,
            len(asset.occurrences),
            0.90,
        )
        try:
            asset.source_ssim = measure_media_ssim(
                reference,
                candidate,
                is_video=True,
                width=asset.width,
                height=asset.height,
            )
            asset.display_ssim = measure_media_ssim(
                reference,
                candidate,
                is_video=True,
                width=asset.display_width_px or asset.width,
                height=asset.display_height_px or asset.height,
            )
            if min(asset.source_ssim, asset.display_ssim) >= asset.applied_threshold:
                asset.quality_status = "passed"
                asset.quality_reason = ""
                continue
            asset.quality_reason = "video_ssim_below_threshold"
        except (OSError, ValueError, subprocess.CalledProcessError) as exc:
            asset.quality_reason = f"video_quality_unavailable: {exc}"
        if (
            not restore_failed
            and not forced
            and asset.quality_reason == "video_ssim_below_threshold"
        ):
            asset.quality_status = "below_threshold"
            if logger is not None:
                logger(f"Quality audit below threshold: {asset.media_path}")
            continue
        asset.output_path = asset.extracted_path
        asset.output_media_path = asset.media_path
        asset.target_bytes = asset.zip_size
        asset.status = "copied"
        asset.quality_status = "restored_original"
        if logger is not None:
            logger(f"Quality floor restored original video: {asset.media_path}")

    for asset in image_assets.values():
        if asset.status not in {"encoded", "reused"}:
            asset.quality_status = "unchanged"
            continue
        redline = (
            0.96
            if asset.content_type == "photo"
            else 0.98
            if asset.content_type in {"line_art", "screenshot_or_text"}
            else 0.97
        )
        asset.applied_threshold = applied_quality_threshold(
            redline if forced else image_threshold,
            asset.max_area_ratio,
            len(asset.occurrences),
            redline,
        )
        while not image_quality_passes(
            asset, preserve_image_metadata=preserve_image_metadata
        ):
            if not restore_failed:
                break
            suffix = Path(asset.media_path).suffix.lower()
            retryable = asset.quality_reason == "image_quality_below_threshold"
            if retryable and asset.scale < 0.999:
                asset.scale = 1.0
            elif (
                retryable
                and suffix in JPEG_EXTENSIONS | WEBP_EXTENSIONS
                and asset.quality < 100
            ):
                asset.quality = min(100, asset.quality + 3)
            else:
                break
            candidate_dir = Path(asset.output_path).parent
            asset.status = "planned"
            encode_image_asset(asset, candidate_dir)
            if asset.status != "encoded":
                break
        if asset.quality_status == "passed":
            continue
        if (
            not restore_failed
            and not forced
            and asset.quality_reason == "image_quality_below_threshold"
        ):
            asset.quality_status = "below_threshold"
            if logger is not None:
                logger(f"Quality audit below threshold: {asset.media_path}")
            continue
        asset.output_path = asset.extracted_path
        asset.target_bytes = asset.zip_size
        asset.status = "copied"
        asset.quality_status = "restored_original"
        if logger is not None:
            logger(f"Quality floor restored original image: {asset.media_path}")


def image_save_kwargs(asset: ImageAsset, image: Image.Image) -> dict[str, Any]:
    suffix = Path(asset.media_path).suffix.lower()
    kwargs: dict[str, Any] = {}
    if suffix in JPEG_EXTENSIONS:
        kwargs.update(
            {
                "format": "JPEG",
                "quality": asset.quality,
                "optimize": True,
                "progressive": True,
            }
        )
        icc_profile = image.info.get("icc_profile")
        if icc_profile:
            kwargs["icc_profile"] = icc_profile
        exif = image.info.get("exif")
        if exif:
            kwargs["exif"] = exif
        dpi = image.info.get("dpi")
        if dpi:
            kwargs["dpi"] = dpi
    elif suffix in WEBP_EXTENSIONS:
        kwargs.update({"format": "WEBP", "quality": asset.quality, "method": 6})
        icc_profile = image.info.get("icc_profile")
        if icc_profile:
            kwargs["icc_profile"] = icc_profile
        for key in ("exif", "xmp"):
            if image.info.get(key):
                kwargs[key] = image.info[key]
    elif suffix in PNG_EXTENSIONS:
        kwargs.update({"format": "PNG", "optimize": True, "compress_level": 9})
        for key in ("icc_profile", "exif", "dpi"):
            if image.info.get(key):
                kwargs[key] = image.info[key]
        if getattr(image, "text", None):
            pnginfo = PngImagePlugin.PngInfo()
            for key, value in image.text.items():
                pnginfo.add_text(key, value)
            kwargs["pnginfo"] = pnginfo
    return kwargs


def prepare_image_for_save(asset: ImageAsset, image: Image.Image) -> Image.Image:
    suffix = Path(asset.media_path).suffix.lower()
    working = image.copy()
    if asset.scale < 0.999 and working.width > 1 and working.height > 1:
        next_width = max(1, int(round(working.width * asset.scale)))
        next_height = max(1, int(round(working.height * asset.scale)))
        working = working.resize((next_width, next_height), Image.Resampling.LANCZOS)
    if suffix in JPEG_EXTENSIONS and working.mode not in {"RGB", "L"}:
        if "A" in working.mode:
            background = Image.new("RGB", working.size, (255, 255, 255))
            background.paste(
                working.convert("RGBA"), mask=working.convert("RGBA").getchannel("A")
            )
            working = background
        else:
            working = working.convert("RGB")
    return working


def encode_image_asset(
    asset: ImageAsset, output_dir: Path, asset_pre_encoded: dict[str, str] | None = None
) -> None:
    source = Path(asset.extracted_path)
    target = output_dir / Path(asset.media_path).name

    if asset_pre_encoded and asset.media_path in asset_pre_encoded:
        shutil.copy2(asset_pre_encoded[asset.media_path], target)
        asset.output_path = str(target)
        asset.status = "reused"
        return

    if asset.status in {"unsupported", "copy_requested"}:
        shutil.copy2(source, target)
        asset.output_path = str(target)
        asset.status = "copied"
        return

    suffix = Path(asset.media_path).suffix.lower()
    if suffix not in JPEG_EXTENSIONS | WEBP_EXTENSIONS | PNG_EXTENSIONS:
        shutil.copy2(source, target)
        asset.output_path = str(target)
        asset.status = "copied"
        asset.reason = "Unsupported image extension"
        return

    try:
        with Image.open(source) as image:
            working = prepare_image_for_save(asset, image)
            kwargs = image_save_kwargs(asset, image)
            if not kwargs:
                raise OSError(f"Unsupported image format: {asset.media_path}")
            working.save(target, **kwargs)
    except (UnidentifiedImageError, OSError) as exc:
        shutil.copy2(source, target)
        asset.output_path = str(target)
        asset.status = "copied"
        asset.reason = f"Image compression skipped: {exc}"
        return

    if not target.exists() or target.stat().st_size >= source.stat().st_size:
        shutil.copy2(source, target)
        asset.quality = 100
        asset.scale = 1.0
        asset.status = "copied"
        asset.reason = "Compressed image was not smaller"
    else:
        asset.status = "encoded"
    asset.output_path = str(target)
    asset.target_bytes = target.stat().st_size if target.exists() else asset.zip_size


def preserve_original_video_plan(asset: VideoAsset) -> None:
    asset.selected_height = asset.height
    asset.audio_kbps = asset.original_audio_kbps if asset.has_audio else 0
    asset.target_video_kbps = asset.original_video_kbps
    asset.target_total_kbps = asset.original_total_kbps
    asset.target_fps = asset.original_fps
    asset.target_bytes = asset.zip_size


def assign_quality_plan(
    assets: dict[str, VideoAsset],
    target_video_bytes: int,
    min_height: int,
    profile: str,
    config: RuntimeConfig,
) -> None:
    preserved_assets = [
        asset for asset in assets.values() if not asset.audio_stream_usable
    ]
    for asset in preserved_assets:
        preserve_original_video_plan(asset)
    video_list = [asset for asset in assets.values() if asset.audio_stream_usable]
    target_video_bytes = max(
        1, target_video_bytes - sum(asset.zip_size for asset in preserved_assets)
    )
    if not video_list:
        return
    current_video_bytes = sum(asset.zip_size for asset in video_list)
    compression_ratio = target_video_bytes / max(1, current_video_bytes)
    sorted_sizes = sorted(asset.zip_size for asset in video_list)
    sorted_kbps = sorted(asset.original_video_kbps for asset in video_list)
    median_size = sorted_sizes[len(sorted_sizes) // 2]
    median_kbps = sorted_kbps[len(sorted_kbps) // 2]

    for asset in video_list:
        asset.audio_kbps = recommend_audio_kbps(asset, compression_ratio, config)
        rendition = rendition_for_asset(asset, config)
        apply_video_bitrate_bounds(asset, rendition)

    def total_min_bytes() -> int:
        return sum(
            int(
                (asset.min_video_kbps + asset.audio_kbps)
                * 1000
                / 8
                * asset.duration_sec
            )
            for asset in video_list
        )

    while total_min_bytes() > target_video_bytes:
        candidates = [
            a
            for a in video_list
            if a.selected_height > max(min_height, a.min_allowed_height)
        ]
        if not candidates:
            break
        asset = min(
            candidates,
            key=lambda item: asset_priority_score(item, median_size, median_kbps),
        )
        asset.selected_height = lower_bucket(
            asset.selected_height,
            max(min_height, asset.min_allowed_height),
        )
        rendition = rendition_for_asset(asset, config)
        asset.audio_kbps = recommend_audio_kbps(asset, compression_ratio, config)
        apply_video_bitrate_bounds(asset, rendition)

    while total_min_bytes() > target_video_bytes:
        candidates = [
            a
            for a in video_list
            if a.selected_height > max(min_height, hard_min_height_for_asset(a))
        ]
        if not candidates:
            break
        asset = min(
            candidates,
            key=lambda item: asset_priority_score(item, median_size, median_kbps),
        )
        asset.selected_height = lower_bucket(
            asset.selected_height, max(min_height, hard_min_height_for_asset(asset))
        )
        rendition = rendition_for_asset(asset, config)
        asset.audio_kbps = recommend_audio_kbps(asset, compression_ratio, config)
        apply_video_bitrate_bounds(asset, rendition)

    min_bytes = total_min_bytes()
    extra_pool = max(0, target_video_bytes - min_bytes)
    remaining = {
        asset.media_path: max(0, asset.max_video_kbps - asset.min_video_kbps)
        for asset in video_list
    }
    weights = {
        asset.media_path: max(
            0.05, asset_priority_score(asset, median_size, median_kbps)
        )
        for asset in video_list
    }
    bonus_bytes: dict[str, int] = {asset.media_path: 0 for asset in video_list}

    while extra_pool > 0:
        active = [asset for asset in video_list if remaining[asset.media_path] > 0]
        if not active:
            break
        total_weight = sum(weights[asset.media_path] for asset in active)
        if total_weight <= 0:
            break
        progress = False
        for asset in active:
            share = extra_pool * (weights[asset.media_path] / total_weight)
            max_extra = int(remaining[asset.media_path] * 1000 / 8 * asset.duration_sec)
            grant = min(max_extra, max(1, int(share)))
            if grant <= 0:
                continue
            bonus_bytes[asset.media_path] += grant
            remaining_kbps_delta = int(grant * 8 / asset.duration_sec / 1000)
            remaining[asset.media_path] = max(
                0, remaining[asset.media_path] - remaining_kbps_delta
            )
            extra_pool -= grant
            progress = True
            if extra_pool <= 0:
                break
        if not progress:
            break

    for asset in video_list:
        base_bytes = int(
            (asset.min_video_kbps + asset.audio_kbps) * 1000 / 8 * asset.duration_sec
        )
        asset.target_bytes = base_bytes + bonus_bytes[asset.media_path]
        target_total_kbps = max(
            1, int(asset.target_bytes * 8 / asset.duration_sec / 1000)
        )
        max_total_kbps = max(asset.audio_kbps, int(asset.original_total_kbps * 0.98))
        asset.target_total_kbps = min(target_total_kbps, max_total_kbps)
        asset.target_video_kbps = max(1, asset.target_total_kbps - asset.audio_kbps)
        asset.target_video_kbps = min(asset.target_video_kbps, asset.max_video_kbps)
        asset.target_video_kbps = max(asset.target_video_kbps, asset.min_video_kbps)
        asset.target_total_kbps = asset.target_video_kbps + asset.audio_kbps
        asset.target_bytes = int(
            asset.target_total_kbps * 1000 / 8 * asset.duration_sec
        )


def height_bucket_for_source(asset: VideoAsset, max_height: int) -> int:
    source_height = max(2, min(asset.height, max_height))
    return normalize_bucket(source_height, "floor")


def profile_target_fps(asset: VideoAsset, profile: str) -> float:
    original_fps = asset.original_fps
    if original_fps <= 0:
        return 0.0

    rules = PROFILE_QUALITY_RULES.get(profile, PROFILE_QUALITY_RULES["balanced"])
    if not rules["dynamic_height"]:
        return original_fps

    if (
        asset.max_area_ratio >= 0.45
        or asset.max_width_ratio >= 0.78
        or asset.max_height_ratio >= 0.78
    ):
        cap = rules["fps_large"]
    elif asset.max_area_ratio >= 0.12:
        cap = rules["fps_medium"]
    else:
        cap = rules["fps_small"]

    if cap is None:
        return original_fps
    return min(original_fps, float(cap))


def assign_profile_plan(
    assets: dict[str, VideoAsset],
    profile: str,
    min_height: int,
    config: RuntimeConfig,
) -> None:
    rules = PROFILE_QUALITY_RULES.get(profile, PROFILE_QUALITY_RULES["balanced"])
    max_profile_height = int(rules["max_height"])
    min_profile_height = int(rules.get("min_height", STD_HEIGHTS[0]))
    bitrate_bias = float(rules["bitrate_bias"])

    for asset in assets.values():
        if not asset.audio_stream_usable:
            preserve_original_video_plan(asset)
            continue
        if bool(rules["dynamic_height"]):
            asset.selected_height = pick_bucket(
                asset.display_height_px,
                max(min_height, min_profile_height, asset.min_allowed_height),
                min(max_profile_height, int(config.render_limits["max_output_height"])),
            )
        else:
            asset.selected_height = height_bucket_for_source(
                asset,
                min(max_profile_height, int(config.render_limits["max_output_height"])),
            )

        rendition = rendition_for_asset(asset, config)
        asset.audio_kbps = recommend_audio_kbps(asset, 1.0, config)
        apply_video_bitrate_bounds(asset, rendition)
        dynamic_range = max(0, asset.max_video_kbps - asset.min_video_kbps)
        asset.target_video_kbps = asset.min_video_kbps + int(
            dynamic_range * bitrate_bias
        )
        asset.target_video_kbps = min(asset.target_video_kbps, asset.max_video_kbps)
        asset.target_video_kbps = max(asset.target_video_kbps, asset.min_video_kbps)
        asset.target_total_kbps = asset.target_video_kbps + asset.audio_kbps
        asset.target_bytes = int(
            asset.target_total_kbps * 1000 / 8 * asset.duration_sec
        )
        asset.target_fps = profile_target_fps(asset, profile)


def assign_video_copy_plan(
    assets: dict[str, VideoAsset],
    asset_pre_encoded: dict[str, str] | None = None,
) -> None:
    pre_encoded_paths = set(asset_pre_encoded or {})
    for asset in assets.values():
        if asset.audio_stream_usable:
            asset.selected_height = height_bucket_for_source(
                asset, max(STD_HEIGHTS[-1], asset.height)
            )
            asset.audio_kbps = asset.original_audio_kbps if asset.has_audio else 0
            asset.target_video_kbps = asset.original_video_kbps
            asset.target_total_kbps = asset.original_total_kbps
            asset.target_fps = asset.original_fps
            asset.target_bytes = asset.zip_size
        else:
            preserve_original_video_plan(asset)
        if asset.media_path in pre_encoded_paths:
            asset.status = "reuse_requested"
        else:
            asset.output_media_path = asset.media_path
            asset.status = "copy_requested"


def scale_dims(asset: VideoAsset, config: RuntimeConfig) -> tuple[int, int]:
    if asset.width <= 0 or asset.height <= 0:
        raise SystemExit(f"Invalid source resolution for {asset.media_path}")
    max_h = min(
        int(config.render_limits["max_output_height"]), max(2, asset.selected_height)
    )
    factor = min(
        1.0,
        max_h / asset.height,
        int(config.render_limits["max_long_edge"]) / max(asset.width, asset.height),
    )
    scaled_w = ceil_even(asset.width * factor)
    scaled_h = ceil_even(asset.height * factor)
    return scaled_w, scaled_h


def should_copy(asset: VideoAsset) -> bool:
    if not asset.audio_stream_usable:
        return True
    if media_needs_mp4(asset.media_path):
        return False
    if reduces_frame_rate(asset):
        return False
    if asset.selected_height < asset.height:
        return False
    same_height = asset.selected_height >= asset.height
    near_original_bitrate = asset.target_total_kbps >= int(
        asset.original_total_kbps * 0.98
    )
    near_original_size = asset.target_bytes >= int(asset.zip_size * 0.98)
    return near_original_size or (same_height and near_original_bitrate)


def append_audio_args(cmd: list[str], asset: VideoAsset) -> None:
    if asset.has_audio:
        cmd.extend(["-c:a", "aac", "-b:a", f"{asset.audio_kbps}k", "-ac", "2"])
    else:
        cmd.append("-an")


def encode_gpu(
    asset: VideoAsset,
    source: Path,
    target: Path,
    scale_filter: str,
    video_rate: str,
    max_rate: str,
    buf_size: str,
    encoder: str,
    cancel_callback: CancelCallback | None = None,
    progress_callback: Callable[[float], None] | None = None,
) -> None:
    cmd = [
        "ffmpeg",
        "-y",
        "-i",
        str(source),
        "-vf",
        scale_filter,
        "-c:v",
        encoder,
        "-pix_fmt",
        "yuv420p",
        "-b:v",
        video_rate,
        "-maxrate",
        max_rate,
        "-bufsize",
        buf_size,
    ]
    if encoder == "h264_videotoolbox":
        cmd.extend(["-profile:v", "main"])
    append_frame_rate_mode(cmd, asset)
    append_audio_args(cmd, asset)
    cmd.extend(["-movflags", "+faststart", str(target)])
    if progress_callback is not None:
        cmd.extend(["-progress", "-", "-nostats"])
    run(
        cmd,
        cancel_callback=cancel_callback,
        progress_callback=progress_callback,
        duration_sec=asset.duration_sec,
    )


def encode_asset(
    asset: VideoAsset,
    output_dir: Path,
    preset: str,
    config: RuntimeConfig,
    encoder_mode: str = "auto",
    logger: Logger | None = None,
    cancel_callback: CancelCallback | None = None,
    progress_callback: Callable[[float], None] | None = None,
    asset_pre_encoded: dict[str, str] | None = None,
) -> None:
    source = Path(asset.extracted_path)
    output_media_path = asset.output_media_path or asset.media_path
    target = output_dir / Path(output_media_path).name

    if asset_pre_encoded and asset.media_path in asset_pre_encoded:
        shutil.copy2(asset_pre_encoded[asset.media_path], target)
        asset.output_path = str(target)
        asset.status = "reused"
        if logger is not None:
            logger(f"Skipping {asset.media_path}, copied from pre-encoded asset")
        return

    if asset.status == "copy_requested":
        shutil.copy2(source, target)
        asset.output_path = str(target)
        asset.status = "copied"
        if logger is not None:
            logger(f"Keeping {asset.media_path}, copied from original asset")
        return

    if should_copy(asset):
        shutil.copy2(source, target)
        asset.output_path = str(target)
        asset.status = "copied"
        if not asset.audio_stream_usable:
            asset.quality_reason = "unusable_audio_stream_preserved"
            if logger is not None:
                logger(
                    f"Keeping {asset.media_path}, original bytes preserved because "
                    "its audio stream is empty or has invalid duration"
                )
        return

    scaled_w, scaled_h = scale_dims(asset, config)
    asset.status = "encoded"
    passlog = output_dir / f"{source.stem}.passlog"
    null_sink = "NUL" if os.name == "nt" else "/dev/null"
    video_rate = f"{asset.target_video_kbps}k"
    max_rate = (
        f"{max(int(asset.target_video_kbps * 1.25), asset.target_video_kbps + 100)}k"
    )
    buf_size = (
        f"{max(int(asset.target_video_kbps * 2), asset.target_video_kbps + 200)}k"
    )
    video_filters = [f"scale={scaled_w}:{scaled_h}:flags=lanczos"]
    if reduces_frame_rate(asset):
        video_filters.append(f"fps={asset.target_fps:g}")
    scale_filter = ",".join(video_filters)

    if encoder_mode in {"auto", "gpu"}:
        gpu_encoders = usable_gpu_encoder_names()
        if not gpu_encoders:
            if logger is not None:
                logger("No usable GPU H.264 encoder found; falling back to CPU x264.")
        else:
            for gpu_encoder in gpu_encoders:
                try:
                    if logger is not None:
                        logger(
                            f"Video encoder: GPU ({gpu_encoder}) for {asset.media_path}"
                        )
                    encode_gpu(
                        asset,
                        source,
                        target,
                        scale_filter,
                        video_rate,
                        max_rate,
                        buf_size,
                        gpu_encoder,
                        cancel_callback=cancel_callback,
                        progress_callback=progress_callback,
                    )
                    validate_encoded_asset(asset, target, scaled_w, scaled_h)
                    asset.output_path = str(target)
                    asset.status = "encoded_gpu"
                    return
                except CancelledError:
                    if target.exists():
                        target.unlink()
                    raise
                except (subprocess.CalledProcessError, ValueError) as exc:
                    if target.exists():
                        target.unlink()
                    if logger is not None:
                        logger(
                            f"GPU encoder {gpu_encoder} failed for {asset.media_path}; trying next encoder."
                        )
                    if encoder_mode == "gpu":
                        logger and logger(f"GPU failure detail: {exc}")
            if logger is not None:
                logger(
                    "All usable GPU encoders failed for this video; falling back to CPU x264."
                )

    if logger is not None:
        logger(f"Video encoder: CPU (libx264) for {asset.media_path}")

    pass1 = [
        "ffmpeg",
        "-y",
        "-i",
        str(source),
        "-vf",
        scale_filter,
        "-c:v",
        "libx264",
        "-preset",
        preset,
        "-pix_fmt",
        "yuv420p",
        "-b:v",
        video_rate,
        "-maxrate",
        max_rate,
        "-bufsize",
        buf_size,
        "-pass",
        "1",
        "-passlogfile",
        str(passlog),
        "-an",
    ]
    append_frame_rate_mode(pass1, asset)
    if progress_callback is not None:
        pass1.extend(["-progress", "-", "-nostats"])
    pass1.extend(["-f", "mp4", null_sink])

    def pass1_progress(f: float) -> None:
        if progress_callback is not None:
            progress_callback(f * 0.5)

    def pass2_progress(f: float) -> None:
        if progress_callback is not None:
            progress_callback(0.5 + f * 0.5)

    run(
        pass1,
        cancel_callback=cancel_callback,
        progress_callback=pass1_progress,
        duration_sec=asset.duration_sec,
    )

    pass2 = [
        "ffmpeg",
        "-y",
        "-i",
        str(source),
        "-vf",
        scale_filter,
        "-c:v",
        "libx264",
        "-preset",
        preset,
        "-pix_fmt",
        "yuv420p",
        "-b:v",
        video_rate,
        "-maxrate",
        max_rate,
        "-bufsize",
        buf_size,
        "-pass",
        "2",
        "-passlogfile",
        str(passlog),
    ]
    append_frame_rate_mode(pass2, asset)
    if asset.has_audio:
        pass2.extend(["-c:a", "aac", "-b:a", f"{asset.audio_kbps}k", "-ac", "2"])
    else:
        pass2.append("-an")
    pass2.extend(["-movflags", "+faststart", str(target)])
    if progress_callback is not None:
        pass2.extend(["-progress", "-", "-nostats"])
    run(
        pass2,
        cancel_callback=cancel_callback,
        progress_callback=pass2_progress,
        duration_sec=asset.duration_sec,
    )
    validate_encoded_asset(asset, target, scaled_w, scaled_h)
    asset.output_path = str(target)


def clone_zip_info(info: ZipInfo, filename: str) -> ZipInfo:
    cloned = ZipInfo(filename, date_time=info.date_time)
    cloned.compress_type = info.compress_type
    cloned.comment = info.comment
    cloned.extra = info.extra
    cloned.internal_attr = info.internal_attr
    cloned.external_attr = info.external_attr
    cloned.create_system = info.create_system
    return cloned


def rewrite_content_types(xml_bytes: bytes, assets: dict[str, VideoAsset]) -> bytes:
    required = {
        Path(asset.output_media_path or asset.media_path).suffix.lower().lstrip(".")
        for asset in assets.values()
        if Path(asset.output_media_path or asset.media_path).suffix
    }
    if not required:
        return xml_bytes

    root = SafeET.fromstring(xml_bytes)
    default_tag = f"{{{CONTENT_TYPES_NS}}}Default"
    existing: dict[str, ET.Element] = {}
    for default in root.findall(default_tag):
        existing[default.attrib.get("Extension", "").lower()] = default
    for extension in sorted(required):
        content_type = (
            mimetypes.guess_type(f"file.{extension}")[0] or "application/octet-stream"
        )
        if extension == "mp4":
            content_type = "video/mp4"
        default = existing.get(extension)
        if default is None:
            ET.SubElement(
                root, default_tag, {"Extension": extension, "ContentType": content_type}
            )
        elif default.attrib.get("ContentType") in {"", "application/octet-stream"}:
            default.attrib["ContentType"] = content_type
    ET.register_namespace("", CONTENT_TYPES_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rewrite_relationships(
    rels_path: str, xml_bytes: bytes, media_path_map: dict[str, str]
) -> bytes:
    source_xml = source_path_from_rels_path(rels_path)
    if source_xml is None:
        return xml_bytes

    root = SafeET.fromstring(xml_bytes)
    changed = False
    for rel in root.findall("pr:Relationship", REL_NS):
        if rel.attrib.get("TargetMode") == "External":
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        resolved = resolve_zip_target(source_xml, target)
        output_media_path = media_path_map.get(resolved)
        if output_media_path and output_media_path != resolved:
            rel.attrib["Target"] = rel_target_to_zip_path(source_xml, output_media_path)
            changed = True

    if not changed:
        return xml_bytes

    ET.register_namespace("", REL_NS["pr"])
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_output_pptx(
    input_pptx: Path,
    output_pptx: Path,
    video_assets: dict[str, VideoAsset],
    image_assets: dict[str, ImageAsset] | None = None,
    relationship_path_map: dict[str, str] | None = None,
    remove_paths: set[str] | None = None,
) -> None:
    image_assets = image_assets or {}
    replacements: dict[str, Path] = {}
    replacement_infos: dict[str, ZipInfo] = {}
    relationship_path_map = dict(relationship_path_map or {})
    remove_paths = set(remove_paths or set())

    with ZipFile(input_pptx, "r") as source_zip:
        for asset in video_assets.values():
            output_media_path = asset.output_media_path or asset.media_path
            replacements[output_media_path] = Path(asset.output_path)
            try:
                replacement_infos[output_media_path] = source_zip.getinfo(
                    asset.media_path
                )
            except KeyError:
                pass
            if output_media_path != asset.media_path:
                relationship_path_map[asset.media_path] = output_media_path
                remove_paths.add(asset.media_path)

        for asset in image_assets.values():
            replacements[asset.media_path] = Path(asset.output_path)
            try:
                replacement_infos[asset.media_path] = source_zip.getinfo(
                    asset.media_path
                )
            except KeyError:
                pass

    patch_output_pptx(
        input_pptx,
        output_pptx,
        replacements,
        relationship_path_map=relationship_path_map,
        replacement_infos=replacement_infos,
        remove_paths=remove_paths,
        video_assets=video_assets,
    )


def _copy_fileobj_to_zip(source, target_zip: ZipFile, info: ZipInfo) -> None:
    if info.is_dir():
        target_zip.writestr(info, b"")
        return
    with target_zip.open(info, "w") as target:
        shutil.copyfileobj(source, target, length=1024 * 1024)


def patch_output_pptx(
    base_pptx: Path,
    output_pptx: Path,
    replacements: dict[str, Path],
    *,
    relationship_path_map: dict[str, str] | None = None,
    replacement_infos: dict[str, ZipInfo] | None = None,
    remove_paths: set[str] | None = None,
    video_assets: dict[str, VideoAsset] | None = None,
) -> None:
    """Patch selected media files into an existing compressed PPTX."""
    relationship_path_map = relationship_path_map or {}
    replacement_infos = replacement_infos or {}
    remove_paths = remove_paths or set()
    video_assets = video_assets or {}
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_pptx.with_name(
        f".{output_pptx.stem}.tmp-{os.getpid()}-{int(time.time() * 1000)}{output_pptx.suffix}"
    )
    written: set[str] = set()

    try:
        with (
            ZipFile(base_pptx, "r") as zin,
            ZipFile(tmp_path, "w", allowZip64=True) as zout,
        ):
            for info in zin.infolist():
                if info.filename in remove_paths and info.filename not in replacements:
                    continue
                if info.filename == "[Content_Types].xml" and video_assets:
                    data = rewrite_content_types(zin.read(info.filename), video_assets)
                    zout.writestr(info, data)
                elif info.filename.endswith(".rels") and relationship_path_map:
                    data = rewrite_relationships(
                        info.filename, zin.read(info.filename), relationship_path_map
                    )
                    zout.writestr(info, data)
                elif info.filename in replacements:
                    with replacements[info.filename].open("rb") as source:
                        _copy_fileobj_to_zip(source, zout, info)
                    written.add(info.filename)
                else:
                    with zin.open(info, "r") as source:
                        _copy_fileobj_to_zip(source, zout, info)

            for media_path, file_path in replacements.items():
                if media_path in written:
                    continue
                source_info = replacement_infos.get(media_path)
                output_info = (
                    clone_zip_info(source_info, media_path)
                    if source_info
                    else ZipInfo(media_path)
                )
                with file_path.open("rb") as source:
                    _copy_fileobj_to_zip(source, zout, output_info)

        validate_pptx_playback_structure(base_pptx, tmp_path)
        tmp_path.replace(output_pptx)
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


def validate_pptx_playback_structure(source_path: Path, output_path: Path) -> None:
    """Fail if media replacement changes slide playback structure."""
    with ZipFile(source_path, "r") as source, ZipFile(output_path, "r") as output:
        slide_parts = [
            name
            for name in source.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        ]
        for name in slide_parts:
            if output.read(name) != source.read(name):
                raise RuntimeError(
                    f"PPTX slide XML changed during media replacement: {name}"
                )

        relationship_parts = [
            name
            for name in source.namelist()
            if name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels")
        ]
        for name in relationship_parts:
            before = SafeET.fromstring(source.read(name))
            after = SafeET.fromstring(output.read(name))

            def identities(root: ET.Element) -> dict[str, tuple[str, str]]:
                return {
                    rel.attrib["Id"]: (
                        rel.attrib.get("Type", ""),
                        rel.attrib.get("TargetMode", ""),
                    )
                    for rel in root.findall("pr:Relationship", REL_NS)
                }

            if identities(after) != identities(before):
                raise RuntimeError(
                    f"PPTX slide relationship identities changed during media replacement: {name}"
                )


def video_report_entry(asset: VideoAsset, config: RuntimeConfig) -> dict[str, Any]:
    scaled_w, scaled_h = scale_dims(asset, config)
    return {
        "media_path": asset.media_path,
        "output_media_path": asset.output_media_path or asset.media_path,
        "zip_size_bytes": asset.zip_size,
        "duration_sec": round(asset.duration_sec, 3),
        "source_resolution": [asset.width, asset.height],
        "display_resolution_estimate": [
            asset.display_width_px,
            asset.display_height_px,
        ],
        "selected_height": asset.selected_height,
        "target_resolution": [scaled_w, scaled_h],
        "original_total_kbps": asset.original_total_kbps,
        "target_video_kbps": asset.target_video_kbps,
        "has_audio": asset.has_audio,
        "audio_stream_usable": asset.audio_stream_usable,
        "audio_kbps": asset.audio_kbps,
        "original_fps": round(asset.original_fps, 3),
        "original_frame_count": asset.original_frame_count or None,
        "target_fps": round(asset.target_fps, 3) if asset.target_fps else None,
        "output_frame_count": asset.output_frame_count or None,
        "status": asset.status,
        "source_ssim": asset.source_ssim,
        "display_ssim": asset.display_ssim,
        "applied_threshold": asset.applied_threshold,
        "quality_status": asset.quality_status,
        "quality_reason": asset.quality_reason,
        "occurrences": [asdict(item) for item in asset.occurrences],
    }


def image_report_entry(asset: ImageAsset) -> dict[str, Any]:
    return {
        "media_path": asset.media_path,
        "output_media_path": asset.output_media_path or asset.media_path,
        "zip_size_bytes": asset.zip_size,
        "source_resolution": [asset.width, asset.height],
        "display_resolution_estimate": [
            asset.display_width_px,
            asset.display_height_px,
        ],
        "max_area_ratio": round(asset.max_area_ratio, 6),
        "content_type": asset.content_type,
        "format": asset.image_format,
        "mode": asset.mode,
        "quality": asset.quality,
        "scale": round(asset.scale, 3),
        "target_bytes": asset.target_bytes,
        "status": asset.status,
        "reason": asset.reason,
        "source_ssim": asset.source_ssim,
        "display_ssim": asset.display_ssim,
        "edge_similarity": asset.edge_similarity,
        "alpha_similarity": asset.alpha_similarity,
        "applied_threshold": asset.applied_threshold,
        "quality_status": asset.quality_status,
        "quality_reason": asset.quality_reason,
        "metadata_preserved": asset.metadata_preserved,
        "occurrences": [asdict(item) for item in asset.occurrences],
    }


def write_standalone_report(
    report_path: Path,
    input_path: Path,
    output_path: Path,
    *,
    video_asset: VideoAsset | None = None,
    image_asset: ImageAsset | None = None,
    target_size_mb: float | None = None,
    config: RuntimeConfig | None = None,
    capacity_attempts: list[dict[str, Any]] | None = None,
) -> None:
    input_kind = "standalone_video" if video_asset is not None else "standalone_image"
    report = {
        "input_kind": input_kind,
        "input_pptx": str(input_path),
        "output_pptx": str(output_path),
        "target_size_mb": target_size_mb,
        "target": target_report_fields(target_size_mb, output_path),
        "presentation": {
            "standalone": True,
            "source_name": input_path.name,
            "output_name": output_path.name,
            "target_capacity_attempts": capacity_attempts or [],
        },
        "videos": [],
        "images": [],
    }
    if video_asset is not None:
        if config is not None:
            report["videos"].append(video_report_entry(video_asset, config))
        else:
            report["videos"].append(
                {
                    "media_path": video_asset.media_path,
                    "output_media_path": video_asset.output_media_path
                    or video_asset.media_path,
                    "zip_size_bytes": video_asset.zip_size,
                    "duration_sec": round(video_asset.duration_sec, 3),
                    "source_resolution": [video_asset.width, video_asset.height],
                    "display_resolution_estimate": [
                        video_asset.display_width_px,
                        video_asset.display_height_px,
                    ],
                    "selected_height": video_asset.selected_height
                    or video_asset.height,
                    "target_resolution": [video_asset.width, video_asset.height],
                    "original_total_kbps": video_asset.original_total_kbps,
                    "target_video_kbps": video_asset.target_video_kbps
                    or video_asset.original_video_kbps,
                    "audio_kbps": video_asset.audio_kbps
                    or video_asset.original_audio_kbps,
                    "original_fps": round(video_asset.original_fps, 3),
                    "target_fps": round(video_asset.target_fps, 3)
                    if video_asset.target_fps
                    else round(video_asset.original_fps, 3),
                    "status": video_asset.status,
                    "occurrences": [],
                }
            )
    if image_asset is not None:
        report["images"].append(image_report_entry(image_asset))
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_markdown_report(report_path, report)


def write_report(
    report_path: Path,
    input_pptx: Path,
    output_pptx: Path | None,
    assets: dict[str, VideoAsset],
    image_assets: dict[str, ImageAsset],
    meta: dict[str, Any],
    target_size_mb: float | None,
    config: RuntimeConfig,
) -> None:
    report = {
        "input_pptx": str(input_pptx),
        "output_pptx": str(output_pptx) if output_pptx else None,
        "target_size_mb": target_size_mb,
        "target": target_report_fields(target_size_mb, output_pptx),
        "presentation": meta,
        "videos": [],
        "images": [],
    }
    for asset in sorted(assets.values(), key=lambda item: item.media_path):
        report["videos"].append(video_report_entry(asset, config))
    for asset in sorted(image_assets.values(), key=lambda item: item.media_path):
        report["images"].append(image_report_entry(asset))
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_markdown_report(report_path, report)


def write_markdown_report(report_path: Path, report: dict[str, Any]) -> Path:
    markdown_path = report_path.with_suffix(".md")
    target = report.get("target", {})
    lines = [
        "# 压缩报告",
        "",
        f"- 输入：`{Path(str(report.get('input_pptx') or '')).name}`",
        f"- 输出：`{Path(str(report.get('output_pptx') or '')).name}`",
        f"- 状态：`{target.get('status', 'not_requested')}`",
    ]
    if target.get("target_bytes") is not None:
        lines.extend(
            [
                f"- 目标：{target['target_bytes']:,} bytes",
                f"- 实际：{(target.get('actual_bytes') or 0):,} bytes",
                f"- 差值：{(target.get('delta_bytes') or 0):+,} bytes",
            ]
        )
    attempts = report.get("presentation", {}).get("target_capacity_attempts", [])
    if attempts:
        lines.extend(["", "## 容量闭环", ""])
        for index, attempt in enumerate(attempts, start=1):
            budget = attempt.get("media_budget_bytes")
            budget_text = f"，媒体预算 {budget:,} bytes" if budget is not None else ""
            lines.append(
                f"- 第 {index} 轮 `{attempt.get('kind', 'unknown')}`："
                f"实际 {int(attempt.get('actual_bytes') or 0):,} bytes{budget_text}"
            )
    lines.extend(["", "## 素材结果", ""])
    entries = [
        *(report.get("videos") or []),
        *(report.get("images") or []),
    ]
    if not entries:
        lines.append("无可压缩素材。")
    else:
        for item in entries:
            quality = item.get("quality_status", "not_checked")
            threshold = item.get("applied_threshold")
            threshold_text = (
                f"，阈值 {float(threshold):.4f}" if threshold is not None else ""
            )
            lines.append(
                f"- `{item.get('media_path', '')}`：{item.get('status', 'unknown')}，"
                f"质量 `{quality}`{threshold_text}"
            )
    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return markdown_path


def update_report_entries(
    report_path: Path,
    report_data: dict[str, Any],
    input_pptx: Path,
    output_pptx: Path,
    assets: dict[str, VideoAsset],
    image_assets: dict[str, ImageAsset],
    meta: dict[str, Any],
    config: RuntimeConfig,
) -> None:
    report_data["input_pptx"] = str(input_pptx)
    report_data["output_pptx"] = str(output_pptx)
    report_data["target"] = target_report_fields(
        report_data.get("target_size_mb"), output_pptx
    )
    report_data.setdefault("presentation", meta)

    videos = {
        item.get("media_path"): item
        for item in report_data.get("videos", [])
        if item.get("media_path")
    }
    for asset in assets.values():
        videos[asset.media_path] = video_report_entry(asset, config)
    report_data["videos"] = sorted(videos.values(), key=lambda item: item["media_path"])

    images = {
        item.get("media_path"): item
        for item in report_data.get("images", [])
        if item.get("media_path")
    }
    for asset in image_assets.values():
        images[asset.media_path] = image_report_entry(asset)
    report_data["images"] = sorted(images.values(), key=lambda item: item["media_path"])

    report_path.write_text(
        json.dumps(report_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_markdown_report(report_path, report_data)


def target_size_filename_label(target_size_mb: float) -> str:
    return f"{target_size_mb:.6f}".rstrip("0").rstrip(".").replace(".", "_")


def default_output_path(input_pptx: Path, target_size_mb: float) -> Path:
    suffix = target_size_filename_label(target_size_mb)
    stem = experimental_output_stem(f"{input_pptx.stem}_compressed_{suffix}MB")
    return input_pptx.with_name(f"{stem}.pptx")


def default_video_output_dir(input_pptx: Path, target_size_mb: float) -> Path:
    suffix = target_size_filename_label(target_size_mb)
    stem = experimental_output_stem(f"{input_pptx.stem}_compressed_{suffix}MB")
    return input_pptx.with_name(f"{stem}_videos")


def profile_suffix(video_profile: str, image_profile: str) -> str:
    if image_profile == "none":
        return video_profile
    if video_profile == "none":
        return f"images_{image_profile}"
    if video_profile == image_profile:
        return f"media_{video_profile}"
    return f"media_v{video_profile}_i{image_profile}"


def default_profile_output_path(
    input_pptx: Path, video_profile: str, image_profile: str
) -> Path:
    stem = experimental_output_stem(
        f"{input_pptx.stem}_compressed_{profile_suffix(video_profile, image_profile)}"
    )
    return input_pptx.with_name(f"{stem}.pptx")


def default_profile_video_output_dir(
    input_pptx: Path, video_profile: str, image_profile: str
) -> Path:
    stem = experimental_output_stem(
        f"{input_pptx.stem}_compressed_{profile_suffix(video_profile, image_profile)}"
    )
    return input_pptx.with_name(f"{stem}_videos")


def is_supported_standalone_video(path: Path) -> bool:
    return path.suffix.lower() in VIDEO_EXTENSIONS


def is_supported_standalone_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def standalone_output_suffix(input_path: Path) -> str:
    suffix = input_path.suffix.lower()
    if suffix in MP4_CONTAINER_EXTENSIONS or suffix in IMAGE_EXTENSIONS:
        return suffix
    if suffix in VIDEO_EXTENSIONS:
        return ".mp4"
    return suffix or ".bin"


def default_media_output_path(input_path: Path, target_size_mb: float) -> Path:
    suffix = target_size_filename_label(target_size_mb)
    stem = experimental_output_stem(f"{input_path.stem}_compressed_{suffix}MB")
    return input_path.with_name(f"{stem}{standalone_output_suffix(input_path)}")


def default_profile_media_output_path(
    input_path: Path,
    video_profile: str,
    image_profile: str,
) -> Path:
    chosen_suffix = (
        profile_suffix(video_profile, "none")
        if is_supported_standalone_video(input_path)
        else profile_suffix("none", image_profile)
    )
    stem = experimental_output_stem(f"{input_path.stem}_compressed_{chosen_suffix}")
    return input_path.with_name(f"{stem}{standalone_output_suffix(input_path)}")


def publish_artifacts(
    runtime_video_dir: Path,
    artifact_video_dir: Path,
    runtime_report_path: Path | None,
    artifact_report_path: Path,
) -> tuple[Path | None, Path | None]:
    published_video_dir: Path | None = None
    published_report_path: Path | None = None

    if runtime_video_dir.exists() and any(runtime_video_dir.iterdir()):
        if runtime_video_dir.resolve() != artifact_video_dir.resolve():
            if artifact_video_dir.exists():
                shutil.rmtree(artifact_video_dir, ignore_errors=True)
            shutil.copytree(runtime_video_dir, artifact_video_dir)
        published_video_dir = artifact_video_dir

    if runtime_report_path is not None and runtime_report_path.exists():
        artifact_report_path.parent.mkdir(parents=True, exist_ok=True)
        if runtime_report_path.resolve() != artifact_report_path.resolve():
            shutil.copy2(runtime_report_path, artifact_report_path)
        published_report_path = artifact_report_path

    return published_video_dir, published_report_path


def print_plan(
    assets: dict[str, VideoAsset],
    image_assets: dict[str, ImageAsset],
    input_pptx: Path,
    target_size_mb: float | None,
    target_video_bytes: int | None,
    non_video_bytes: int,
    config: RuntimeConfig,
    logger: Logger = print,
) -> None:
    logger(f"Input: {input_pptx}")
    if target_size_mb is None:
        logger("Target PPTX size: profile preset")
    else:
        logger(f"Target PPTX size: {target_size_mb:.2f} MB")
    logger(f"Non-media size: {non_video_bytes / BYTES_PER_MB:.2f} MB")
    if target_video_bytes is not None:
        logger(f"Video budget: {target_video_bytes / BYTES_PER_MB:.2f} MB")
    logger("")
    for asset in sorted(assets.values(), key=lambda item: item.media_path):
        dims = scale_dims(asset, config)
        if not asset.audio_stream_usable:
            planned_status = "copy original (audio preserved)"
        elif asset.status == "copy_requested":
            planned_status = "copy original"
        elif asset.status == "reuse_requested":
            planned_status = "reuse"
        else:
            planned_status = "copy" if should_copy(asset) else "encode"
        output_hint = (
            f" -> {asset.output_media_path}"
            if asset.output_media_path and asset.output_media_path != asset.media_path
            else ""
        )
        fps_hint = f", fps {asset.target_fps:g}" if asset.target_fps else ""
        logger(
            f"{asset.media_path}: "
            f"{asset.zip_size / 1024 / 1024:.2f} MiB -> "
            f"{asset.target_bytes / 1024 / 1024:.2f} MiB, "
            f"display~{asset.display_width_px}x{asset.display_height_px}, "
            f"floor {asset.min_allowed_height}p, "
            f"{planned_status}{output_hint} {dims[0]}x{dims[1]} @ {asset.target_video_kbps}k "
            f"(audio {asset.audio_kbps}k{fps_hint})"
        )
    for asset in sorted(image_assets.values(), key=lambda item: item.media_path):
        if asset.status == "copy_requested":
            planned_status = "copy"
        elif asset.status == "unsupported":
            planned_status = "skip"
        else:
            planned_status = "optimize"
        scale_hint = f", scale {asset.scale:.2f}" if asset.scale < 0.999 else ""
        logger(
            f"{asset.media_path}: "
            f"{asset.zip_size / 1024 / 1024:.2f} MiB -> "
            f"{asset.target_bytes / 1024 / 1024:.2f} MiB, "
            f"{planned_status} quality {asset.quality}{scale_hint}"
        )


def compact_pptx(
    args: argparse.Namespace,
    logger: Logger = print,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    asset_pre_encoded: dict[str, str] | None = None,
) -> dict[str, Any]:
    video_profile = getattr(args, "profile", "high") or "high"
    image_profile = getattr(args, "image_profile", "high") or "high"
    compress_videos = video_profile != "none"
    compress_images = image_profile != "none"
    incremental_mode = asset_pre_encoded is not None
    if not compress_videos and not compress_images and not incremental_mode:
        raise SystemExit("No media compression profile is enabled.")

    if compress_videos:
        ensure_binary("ffmpeg")
        ensure_binary("ffprobe")
    runtime_config = load_runtime_config(args.config, video_profile)

    input_pptx = Path(args.input_pptx).expanduser().resolve()
    if not input_pptx.exists():
        raise SystemExit(f"Input file not found: {input_pptx}")

    target_size_mb = args.target_size_mb
    target_mode = target_size_mb is not None
    source_size_bytes = input_pptx.stat().st_size
    source_size_mb = source_size_bytes / BYTES_PER_MB
    if target_mode and mb_to_bytes(target_size_mb) >= source_size_bytes:
        reason = (
            f"Target size ({target_size_mb:.2f} MB) is not smaller than source "
            f"({source_size_mb:.2f} MB); skipped."
        )
        logger(reason)
        skip_report_path = write_target_skip_report(input_pptx, target_size_mb, reason)
        if progress_callback is not None:
            progress_callback(1, 1, "目标大小不小于源文件，已跳过")
        return {
            "input_pptx": input_pptx,
            "output_pptx": input_pptx,
            "report_path": skip_report_path,
            "video_output_dir": None,
            "skipped": True,
            "reason": reason,
        }

    output_pptx = (
        args.output.expanduser().resolve()
        if args.output
        else (
            default_output_path(input_pptx, target_size_mb)
            if target_mode
            else default_profile_output_path(input_pptx, video_profile, image_profile)
        )
    )
    if args.output is None:
        output_pptx = quality_variant_output_path(
            output_pptx, getattr(args, "quality_mode", "safe")
        )
    work_dir = (
        args.work_dir.expanduser().resolve()
        if args.work_dir
        else Path(tempfile.mkdtemp(prefix=runtime_temp_prefix("pptx_compact_")))
    )
    work_dir.mkdir(parents=True, exist_ok=True)
    video_output_dir = work_dir / "compressed_videos"
    video_output_dir.mkdir(parents=True, exist_ok=True)
    artifact_video_output_dir = (
        args.video_output_dir.expanduser().resolve()
        if args.video_output_dir
        else (
            default_video_output_dir(input_pptx, target_size_mb)
            if target_mode
            else default_profile_video_output_dir(
                input_pptx, video_profile, image_profile
            )
        )
    )
    artifact_report_path = output_pptx.with_suffix(".report.json")
    report_path: Path | None = None
    assets: dict[str, VideoAsset] | None = None
    image_assets: dict[str, ImageAsset] | None = None
    meta: dict[str, Any] | None = None
    published_video_dir: Path | None = None
    published_report_path: Path | None = None

    def check_cancelled() -> None:
        if cancel_callback is not None and cancel_callback():
            raise CancelledError("Cancelled")

    try:
        check_cancelled()
        if progress_callback is not None:
            progress_callback(0, 1, "正在解析演示文稿")
        process_videos = compress_videos or incremental_mode
        process_images = compress_images or incremental_mode
        assets, image_assets, meta = parse_pptx_assets(
            input_pptx,
            render_width=args.slide_render_width,
            render_height=args.slide_render_height,
            overscan=args.overscan,
            min_height=args.min_height,
            max_height=min(
                args.max_height, int(PROFILE_QUALITY_RULES[video_profile]["max_height"])
            ),
            config=runtime_config,
            include_videos=process_videos,
        )
        if not process_videos:
            assets = {}
        if not process_images:
            image_assets = {}
        meta["quality_mode"] = getattr(args, "quality_mode", "safe")
        if not assets and not image_assets:
            reason = "No selected media assets were found in this PPTX."
            logger(reason)
            if progress_callback is not None:
                progress_callback(1, 1, "未找到可压缩媒体，已跳过")
            return {
                "input_pptx": input_pptx,
                "output_pptx": input_pptx,
                "report_path": None,
                "video_output_dir": None,
                "skipped": True,
                "reason": reason,
            }

        total_steps = len(assets) + len(image_assets) + 5
        check_cancelled()
        if assets and progress_callback is not None:
            progress_callback(1, total_steps, "正在提取内嵌视频")
        if assets:
            extract_videos(
                input_pptx, assets, work_dir, progress_callback=progress_callback
            )
        if image_assets and progress_callback is not None:
            progress_callback(2, total_steps, "正在提取内嵌图片")
        if image_assets:
            extract_images(
                input_pptx, image_assets, work_dir, progress_callback=progress_callback
            )
        current_video_bytes = sum(asset.zip_size for asset in assets.values())
        current_image_bytes = sum(asset.zip_size for asset in image_assets.values())
        non_media_bytes = (
            input_pptx.stat().st_size - current_video_bytes - current_image_bytes
        )
        duplicate_image_path_map = consolidate_exact_duplicate_images(image_assets)
        removable_media_paths = set(meta.get("orphan_image_paths", [])) | set(
            duplicate_image_path_map
        )
        meta["exact_duplicate_image_paths"] = duplicate_image_path_map
        meta["removed_orphan_image_paths"] = sorted(meta.get("orphan_image_paths", []))
        initial_video_heights = {
            asset.media_path: asset.selected_height for asset in assets.values()
        }

        def assign_target_media_plan(media_budget: int) -> int | None:
            target_video_budget, target_image_budget = allocate_media_budgets(
                current_video_bytes,
                current_image_bytes,
                media_budget,
                video_profile,
                image_profile,
            )
            if assets:
                if compress_videos:
                    for asset in assets.values():
                        asset.selected_height = initial_video_heights[asset.media_path]
                    video_budget = max(1, target_video_budget)
                    assign_quality_plan(
                        assets,
                        video_budget,
                        args.min_height,
                        video_profile,
                        runtime_config,
                    )
                    for asset in assets.values():
                        asset.target_fps = profile_target_fps(asset, video_profile)
                else:
                    video_budget = current_video_bytes
                    assign_video_copy_plan(assets, asset_pre_encoded)
            else:
                video_budget = None
            assign_image_plan(
                image_assets,
                image_profile,
                max(1, target_image_budget) if image_assets else None,
            )
            return video_budget

        check_cancelled()
        if progress_callback is not None:
            progress_callback(3, total_steps, "正在计算压缩计划")
        if target_mode:
            target_total_bytes = mb_to_bytes(target_size_mb)
            explicit_reserve_mb = getattr(args, "reserve_mb", None)
            reserve_bytes = (
                mb_to_bytes(explicit_reserve_mb)
                if explicit_reserve_mb is not None
                else dynamic_package_reserve_bytes(target_total_bytes, non_media_bytes)
            )
            target_media_bytes = target_total_bytes - non_media_bytes - reserve_bytes
            if target_media_bytes <= 0:
                raise SystemExit(
                    "Target size is too small. Non-media content already exceeds the requested output size."
                )
            target_video_bytes = assign_target_media_plan(target_media_bytes)
        else:
            target_video_bytes = None
            if assets:
                if compress_videos:
                    assign_profile_plan(
                        assets, video_profile, args.min_height, runtime_config
                    )
                else:
                    assign_video_copy_plan(assets, asset_pre_encoded)
            assign_image_plan(image_assets, image_profile)
        for asset in assets.values():
            if asset.status not in {"copy_requested", "reuse_requested"}:
                asset.status = "would_copy" if should_copy(asset) else "planned"
        print_plan(
            assets,
            image_assets,
            input_pptx,
            target_size_mb,
            target_video_bytes,
            non_media_bytes,
            runtime_config,
            logger=logger,
        )

        report_path = work_dir / artifact_report_path.name
        if args.dry_run:
            report_path = artifact_report_path
            write_report(
                report_path,
                input_pptx,
                None,
                assets,
                image_assets,
                meta,
                target_size_mb,
                runtime_config,
            )
            logger("")
            logger(f"Dry run report written to: {report_path}")
            return {
                "input_pptx": input_pptx,
                "output_pptx": output_pptx,
                "report_path": report_path,
                "video_output_dir": None,
            }

        total_asset_bytes = sum(asset.zip_size for asset in assets.values()) + sum(
            asset.zip_size for asset in image_assets.values()
        )
        base_weight = max(1000.0, float(total_asset_bytes) * 0.1)
        total_weight = total_asset_bytes + (base_weight * 2.0)
        image_output_dir = work_dir / "compressed_images"
        image_output_dir.mkdir(parents=True, exist_ok=True)

        def encode_and_package() -> None:
            current_weight = base_weight
            sorted_assets = sorted(assets.values(), key=lambda item: item.media_path)
            for index, asset in enumerate(sorted_assets, start=1):
                check_cancelled()
                label = f"正在处理视频 {index}/{len(sorted_assets)}: {Path(asset.media_path).name}"
                if progress_callback is not None:
                    progress_callback(current_weight, total_weight, label)

                def asset_progress(f: float) -> None:
                    if progress_callback is not None:
                        progress_callback(
                            current_weight + (asset.zip_size * f), total_weight, label
                        )

                encode_asset(
                    asset,
                    video_output_dir,
                    args.preset,
                    runtime_config,
                    encoder_mode=args.encoder,
                    logger=logger,
                    cancel_callback=cancel_callback,
                    progress_callback=asset_progress,
                    asset_pre_encoded=asset_pre_encoded,
                )
                current_weight += asset.zip_size

            sorted_image_assets = sorted(
                image_assets.values(), key=lambda item: item.media_path
            )
            for index, asset in enumerate(sorted_image_assets, start=1):
                check_cancelled()
                label = f"正在处理图片 {index}/{len(sorted_image_assets)}: {Path(asset.media_path).name}"
                if progress_callback is not None:
                    progress_callback(current_weight, total_weight, label)
                encode_image_asset(
                    asset, image_output_dir, asset_pre_encoded=asset_pre_encoded
                )
                current_weight += asset.zip_size

            audit_encoded_assets(
                assets,
                image_assets,
                video_threshold=float(getattr(args, "video_ssim_threshold", 0.95)),
                image_threshold=float(getattr(args, "image_ssim_threshold", 0.99)),
                forced=getattr(args, "quality_mode", "safe") == "forced",
                logger=logger,
                restore_failed=False,
            )

            check_cancelled()
            if progress_callback is not None:
                progress_callback(
                    total_weight - base_weight, total_weight, "正在重新打包 PPTX"
                )
            build_output_pptx(
                input_pptx,
                output_pptx,
                assets,
                image_assets,
                relationship_path_map=duplicate_image_path_map,
                remove_paths=removable_media_paths,
            )

        encode_and_package()
        capacity_attempts = [
            {
                "kind": "initial",
                "media_budget_bytes": target_media_bytes if target_mode else None,
                "actual_bytes": output_pptx.stat().st_size,
            }
        ]
        if target_mode:
            correction_rounds = 0
            giveback_used = False
            media_budget = target_media_bytes
            while True:
                next_attempt = next_target_media_budget(
                    actual_bytes=output_pptx.stat().st_size,
                    target_bytes=target_total_bytes,
                    current_media_budget=media_budget,
                    maximum_media_budget=current_video_bytes + current_image_bytes,
                    correction_rounds=correction_rounds,
                    giveback_used=giveback_used,
                )
                if next_attempt is None:
                    break
                media_budget, attempt_kind = next_attempt
                correction_rounds += attempt_kind == "correction"
                giveback_used |= attempt_kind == "quality_giveback"
                previous_plan = media_plan_signature(assets, image_assets)
                target_video_bytes = assign_target_media_plan(media_budget)
                if media_plan_signature(assets, image_assets) == previous_plan:
                    logger("Target capacity retry skipped: media plan is unchanged")
                    break
                logger(
                    f"Target capacity {attempt_kind}: retrying with "
                    f"{media_budget / BYTES_PER_MB:.2f} MB media budget"
                )
                encode_and_package()
                capacity_attempts.append(
                    {
                        "kind": attempt_kind,
                        "media_budget_bytes": media_budget,
                        "actual_bytes": output_pptx.stat().st_size,
                    }
                )
            meta["target_capacity_attempts"] = capacity_attempts
        write_report(
            report_path,
            input_pptx,
            output_pptx,
            assets,
            image_assets,
            meta,
            target_size_mb,
            runtime_config,
        )

        # Always publish the report file for quality audit purposes
        artifact_report_path.parent.mkdir(parents=True, exist_ok=True)
        if report_path.resolve() != artifact_report_path.resolve():
            shutil.copy2(report_path, artifact_report_path)
            markdown_report = report_path.with_suffix(".md")
            if markdown_report.is_file():
                shutil.copy2(markdown_report, artifact_report_path.with_suffix(".md"))
        published_report_path = artifact_report_path

        published_video_dir = None
        if args.keep_artifacts:
            published_video_dir, _ = publish_artifacts(
                video_output_dir,
                artifact_video_output_dir,
                None,
                artifact_report_path,
            )

        if progress_callback is not None:
            progress_callback(total_weight, total_weight, "压缩完成")
        logger("")
        logger(f"Output PPTX: {output_pptx}")
        logger(f"Report: {published_report_path}")
        if args.keep_artifacts and published_video_dir is not None:
            logger(f"Compressed videos: {published_video_dir}")

        return {
            "input_pptx": input_pptx,
            "output_pptx": output_pptx,
            "report_path": published_report_path,
            "video_output_dir": published_video_dir if args.keep_artifacts else None,
            "skipped": False,
        }
    except CancelledError:
        raise
    except Exception:
        if assets is not None and image_assets is not None and meta is not None:
            try:
                if report_path is None:
                    report_path = work_dir / artifact_report_path.name
                write_report(
                    report_path,
                    input_pptx,
                    output_pptx if output_pptx.exists() else None,
                    assets,
                    image_assets,
                    meta,
                    target_size_mb,
                    runtime_config,
                )
            except Exception:
                pass
        publish_artifacts(
            video_output_dir,
            artifact_video_output_dir,
            report_path,
            artifact_report_path,
        )
        raise
    finally:
        if not args.keep_work_dir and args.work_dir is None:
            shutil.rmtree(work_dir, ignore_errors=True)


def compact_standalone_video(
    args: argparse.Namespace,
    logger: Logger = print,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
) -> dict[str, Any]:
    video_profile = getattr(args, "profile", "high") or "high"
    input_path = Path(args.input_pptx).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")

    source_size_bytes = input_path.stat().st_size
    source_size_mb = source_size_bytes / BYTES_PER_MB
    target_size_mb = args.target_size_mb
    target_mode = target_size_mb is not None
    if target_mode and mb_to_bytes(target_size_mb) >= source_size_bytes:
        reason = (
            f"Target size ({target_size_mb:.2f} MB) is not smaller than source "
            f"({source_size_mb:.2f} MB); skipped."
        )
        logger(reason)
        skip_report_path = write_target_skip_report(input_path, target_size_mb, reason)
        if progress_callback is not None:
            progress_callback(1, 1, "目标大小不小于源文件，已跳过")
        return {
            "input_pptx": input_path,
            "output_pptx": input_path,
            "report_path": skip_report_path,
            "video_output_dir": None,
            "skipped": True,
            "reason": reason,
        }

    output_path = (
        args.output.expanduser().resolve()
        if args.output
        else (
            default_media_output_path(input_path, target_size_mb)
            if target_mode
            else default_profile_media_output_path(input_path, video_profile, "none")
        )
    )
    if args.output is None:
        output_path = quality_variant_output_path(
            output_path, getattr(args, "quality_mode", "safe")
        )
    expected_suffix = standalone_output_suffix(input_path)
    if output_path.suffix.lower() != expected_suffix:
        output_path = output_path.with_suffix(expected_suffix)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    artifact_report_path = output_path.with_suffix(".report.json")

    if video_profile == "none" and not target_mode:
        shutil.copy2(input_path, output_path)
        probe = ffprobe_json(input_path)
        streams = probe.get("streams", [])
        video_stream = next(
            (s for s in streams if s.get("codec_type") == "video"), None
        )
        audio_stream = next(
            (s for s in streams if s.get("codec_type") == "audio"), None
        )
        if video_stream is None:
            raise SystemExit(f"No video stream found in {input_path}")
        asset = VideoAsset(
            media_path=input_path.name,
            zip_size=source_size_bytes,
            duration_sec=max(
                0.1,
                float(
                    video_stream.get("duration")
                    or probe.get("format", {}).get("duration")
                    or 0.1
                ),
            ),
            width=int(video_stream.get("width", 0)),
            height=int(video_stream.get("height", 0)),
            has_audio=audio_stream is not None,
            audio_stream_usable=(
                audio_stream is None or audio_stream_is_usable(audio_stream)
            ),
            original_fps=parse_fps(
                video_stream.get("avg_frame_rate") or video_stream.get("r_frame_rate")
            ),
            original_frame_count=stream_frame_count(video_stream),
            display_width_px=int(video_stream.get("width", 0)),
            display_height_px=int(video_stream.get("height", 0)),
            min_allowed_height=0,
            extracted_path=str(input_path),
            output_path=str(output_path),
            output_media_path=output_path.name,
            status="copied",
        )
        asset.max_area_ratio = 1.0
        asset.max_width_ratio = 1.0
        asset.max_height_ratio = 1.0
        if not asset.audio_stream_usable:
            asset.quality_reason = "unusable_audio_stream_preserved"
        asset.selected_height = height_bucket_for_source(
            asset,
            max(STD_HEIGHTS[-1], asset.height),
        )
        format_bit_rate = int(float(probe.get("format", {}).get("bit_rate", 0) or 0))
        video_bit_rate = int(float(video_stream.get("bit_rate", 0) or 0))
        asset.original_total_kbps = max(
            1,
            format_bit_rate // 1000
            if format_bit_rate
            else int(source_size_bytes * 8 / asset.duration_sec / 1000),
        )
        asset.original_video_kbps = max(
            1,
            video_bit_rate // 1000
            if video_bit_rate
            else max(asset.original_total_kbps - 96, 1),
        )
        asset.original_audio_kbps = (
            int(float(audio_stream.get("bit_rate", 0) or 0)) // 1000
            if audio_stream is not None
            else 0
        )
        asset.target_video_kbps = asset.original_video_kbps
        asset.audio_kbps = asset.original_audio_kbps
        asset.target_fps = asset.original_fps
        asset.target_bytes = source_size_bytes
        write_standalone_report(
            artifact_report_path,
            input_path,
            output_path,
            video_asset=asset,
            target_size_mb=target_size_mb,
            config=load_runtime_config(args.config, video_profile),
        )
        if progress_callback is not None:
            progress_callback(1, 1, "视频未压缩，已复制输出")
        logger(f"Output media: {output_path}")
        logger(f"Report: {artifact_report_path}")
        return {
            "input_pptx": input_path,
            "output_pptx": output_path,
            "report_path": artifact_report_path,
            "video_output_dir": None,
            "skipped": False,
        }

    ensure_binary("ffmpeg")
    ensure_binary("ffprobe")
    runtime_config = load_runtime_config(args.config, video_profile)
    work_dir = (
        args.work_dir.expanduser().resolve()
        if args.work_dir
        else Path(tempfile.mkdtemp(prefix=runtime_temp_prefix("pptx_compact_media_")))
    )
    work_dir.mkdir(parents=True, exist_ok=True)
    output_dir = work_dir / "compressed_videos"
    output_dir.mkdir(parents=True, exist_ok=True)

    def check_cancelled() -> None:
        if cancel_callback is not None and cancel_callback():
            raise CancelledError("Cancelled")

    try:
        if progress_callback is not None:
            progress_callback(0, 1, "正在分析视频")
        probe = ffprobe_json(input_path)
        streams = probe.get("streams", [])
        video_stream = next(
            (s for s in streams if s.get("codec_type") == "video"), None
        )
        audio_stream = next(
            (s for s in streams if s.get("codec_type") == "audio"), None
        )
        if video_stream is None:
            raise SystemExit(f"No video stream found in {input_path}")

        asset = VideoAsset(
            media_path=input_path.name,
            zip_size=source_size_bytes,
            duration_sec=max(
                0.1,
                float(
                    video_stream.get("duration")
                    or probe.get("format", {}).get("duration")
                    or 0.1
                ),
            ),
            width=int(video_stream.get("width", 0)),
            height=int(video_stream.get("height", 0)),
            has_audio=audio_stream is not None,
            audio_stream_usable=(
                audio_stream is None or audio_stream_is_usable(audio_stream)
            ),
            original_fps=parse_fps(
                video_stream.get("avg_frame_rate") or video_stream.get("r_frame_rate")
            ),
            original_frame_count=stream_frame_count(video_stream),
            display_width_px=int(video_stream.get("width", 0)),
            display_height_px=int(video_stream.get("height", 0)),
            min_allowed_height=0,
            extracted_path=str(input_path),
            output_media_path=output_path.name,
        )
        asset.max_area_ratio = 1.0
        asset.max_width_ratio = 1.0
        asset.max_height_ratio = 1.0
        if not asset.audio_stream_usable:
            asset.quality_reason = "unusable_audio_stream_preserved"
        asset.selected_height = height_bucket_for_source(
            asset,
            min(args.max_height, max(STD_HEIGHTS[-1], asset.height)),
        )
        initial_height = asset.selected_height
        format_bit_rate = int(float(probe.get("format", {}).get("bit_rate", 0) or 0))
        video_bit_rate = int(float(video_stream.get("bit_rate", 0) or 0))
        asset.original_total_kbps = max(
            1,
            format_bit_rate // 1000
            if format_bit_rate
            else int(source_size_bytes * 8 / asset.duration_sec / 1000),
        )
        asset.original_video_kbps = max(
            1,
            video_bit_rate // 1000
            if video_bit_rate
            else max(asset.original_total_kbps - 96, 1),
        )
        asset.original_audio_kbps = (
            int(float(audio_stream.get("bit_rate", 0) or 0)) // 1000
            if audio_stream is not None
            else 0
        )

        asset_map = {asset.media_path: asset}
        if target_mode:
            assign_quality_plan(
                asset_map,
                mb_to_bytes(target_size_mb),
                args.min_height,
                video_profile,
                runtime_config,
            )
            asset.target_fps = profile_target_fps(asset, video_profile)
        elif video_profile == "none":
            assign_video_copy_plan(asset_map)
        else:
            assign_profile_plan(
                asset_map, video_profile, args.min_height, runtime_config
            )

        if asset.status not in {"copy_requested", "reuse_requested"}:
            asset.status = "would_copy" if should_copy(asset) else "planned"
        dims = scale_dims(asset, runtime_config)
        planned_status = (
            "copy original (audio preserved)"
            if not asset.audio_stream_usable
            else "copy original"
            if asset.status in {"copy_requested", "reuse_requested", "would_copy"}
            else "encode"
        )
        fps_hint = f", fps {asset.target_fps:g}" if asset.target_fps else ""
        logger(f"Input: {input_path}")
        if target_mode:
            logger(f"Target media size: {target_size_mb:.2f} MB")
        logger(
            f"{asset.media_path}: "
            f"{asset.zip_size / 1024 / 1024:.2f} MiB -> "
            f"{asset.target_bytes / 1024 / 1024:.2f} MiB, "
            f"{planned_status} {dims[0]}x{dims[1]} @ {asset.target_video_kbps}k "
            f"(audio {asset.audio_kbps}k{fps_hint})"
        )

        check_cancelled()
        if progress_callback is not None:
            progress_callback(0, 1, "正在处理视频 1/1")

        def asset_progress(fraction: float) -> None:
            if progress_callback is not None:
                progress_callback(fraction, 1, "正在处理视频 1/1")

        def encode_video_attempt() -> None:
            encode_asset(
                asset,
                output_dir,
                args.preset,
                runtime_config,
                encoder_mode=args.encoder,
                logger=logger,
                cancel_callback=cancel_callback,
                progress_callback=asset_progress,
            )
            audit_encoded_assets(
                asset_map,
                {},
                video_threshold=float(getattr(args, "video_ssim_threshold", 0.95)),
                image_threshold=float(getattr(args, "image_ssim_threshold", 0.99)),
                forced=getattr(args, "quality_mode", "safe") == "forced",
                logger=logger,
                restore_failed=False,
            )
            shutil.copy2(Path(asset.output_path), output_path)
            asset.output_path = str(output_path)

        encode_video_attempt()
        capacity_attempts = [
            {
                "kind": "initial",
                "media_budget_bytes": asset.target_bytes,
                "actual_bytes": output_path.stat().st_size,
            }
        ]
        if target_mode:
            media_budget = mb_to_bytes(target_size_mb)
            correction_rounds = 0
            giveback_used = False
            while True:
                next_attempt = next_target_media_budget(
                    actual_bytes=output_path.stat().st_size,
                    target_bytes=mb_to_bytes(target_size_mb),
                    current_media_budget=media_budget,
                    maximum_media_budget=source_size_bytes,
                    correction_rounds=correction_rounds,
                    giveback_used=giveback_used,
                )
                if next_attempt is None:
                    break
                media_budget, attempt_kind = next_attempt
                correction_rounds += attempt_kind == "correction"
                giveback_used |= attempt_kind == "quality_giveback"
                previous_plan = media_plan_signature(asset_map, {})
                asset.selected_height = initial_height
                assign_quality_plan(
                    asset_map,
                    media_budget,
                    args.min_height,
                    video_profile,
                    runtime_config,
                )
                asset.target_fps = profile_target_fps(asset, video_profile)
                if media_plan_signature(asset_map, {}) == previous_plan:
                    logger("Target capacity retry skipped: media plan is unchanged")
                    break
                encode_video_attempt()
                capacity_attempts.append(
                    {
                        "kind": attempt_kind,
                        "media_budget_bytes": media_budget,
                        "actual_bytes": output_path.stat().st_size,
                    }
                )
        write_standalone_report(
            artifact_report_path,
            input_path,
            output_path,
            video_asset=asset,
            target_size_mb=target_size_mb,
            config=runtime_config,
            capacity_attempts=capacity_attempts,
        )
        if progress_callback is not None:
            progress_callback(1, 1, "压缩完成")
        logger("")
        logger(f"Output media: {output_path}")
        logger(f"Report: {artifact_report_path}")
        return {
            "input_pptx": input_path,
            "output_pptx": output_path,
            "report_path": artifact_report_path,
            "video_output_dir": None,
            "skipped": False,
        }
    finally:
        if not args.keep_work_dir and args.work_dir is None:
            shutil.rmtree(work_dir, ignore_errors=True)


def compact_standalone_image(
    args: argparse.Namespace,
    logger: Logger = print,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
) -> dict[str, Any]:
    del cancel_callback
    image_profile = getattr(args, "image_profile", "high") or "high"
    input_path = Path(args.input_pptx).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")

    source_size_bytes = input_path.stat().st_size
    source_size_mb = source_size_bytes / BYTES_PER_MB
    target_size_mb = args.target_size_mb
    target_mode = target_size_mb is not None
    if target_mode and mb_to_bytes(target_size_mb) >= source_size_bytes:
        reason = (
            f"Target size ({target_size_mb:.2f} MB) is not smaller than source "
            f"({source_size_mb:.2f} MB); skipped."
        )
        logger(reason)
        skip_report_path = write_target_skip_report(input_path, target_size_mb, reason)
        if progress_callback is not None:
            progress_callback(1, 1, "目标大小不小于源文件，已跳过")
        return {
            "input_pptx": input_path,
            "output_pptx": input_path,
            "report_path": skip_report_path,
            "video_output_dir": None,
            "skipped": True,
            "reason": reason,
        }

    output_path = (
        args.output.expanduser().resolve()
        if args.output
        else (
            default_media_output_path(input_path, target_size_mb)
            if target_mode
            else default_profile_media_output_path(input_path, "none", image_profile)
        )
    )
    if args.output is None:
        output_path = quality_variant_output_path(
            output_path, getattr(args, "quality_mode", "safe")
        )
    if output_path.suffix.lower() != input_path.suffix.lower():
        output_path = output_path.with_suffix(input_path.suffix.lower())
    output_path.parent.mkdir(parents=True, exist_ok=True)
    artifact_report_path = output_path.with_suffix(".report.json")

    if image_profile == "none" and not target_mode:
        shutil.copy2(input_path, output_path)
        asset = ImageAsset(
            media_path=input_path.name,
            zip_size=source_size_bytes,
            extracted_path=str(input_path),
            output_path=str(output_path),
            output_media_path=output_path.name,
            status="copied",
            target_bytes=source_size_bytes,
            reason="profile_none_copy",
        )
        try:
            with Image.open(input_path) as image:
                asset.width = int(image.width)
                asset.height = int(image.height)
                asset.image_format = str(image.format or "").upper()
                asset.mode = image.mode
        except (UnidentifiedImageError, OSError) as exc:
            raise SystemExit(f"Cannot read image: {exc}") from exc
        write_standalone_report(
            artifact_report_path,
            input_path,
            output_path,
            image_asset=asset,
            target_size_mb=target_size_mb,
        )
        if progress_callback is not None:
            progress_callback(1, 1, "图片未压缩，已复制输出")
        logger(f"Output media: {output_path}")
        logger(f"Report: {artifact_report_path}")
        return {
            "input_pptx": input_path,
            "output_pptx": output_path,
            "report_path": artifact_report_path,
            "video_output_dir": None,
            "skipped": False,
        }

    work_dir = (
        args.work_dir.expanduser().resolve()
        if args.work_dir
        else Path(tempfile.mkdtemp(prefix=runtime_temp_prefix("pptx_compact_image_")))
    )
    work_dir.mkdir(parents=True, exist_ok=True)
    output_dir = work_dir / "compressed_images"
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        if progress_callback is not None:
            progress_callback(0, 1, "正在分析图片")
        asset = ImageAsset(
            media_path=input_path.name,
            zip_size=source_size_bytes,
            extracted_path=str(input_path),
            output_media_path=output_path.name,
        )
        try:
            with Image.open(input_path) as image:
                asset.width = int(image.width)
                asset.height = int(image.height)
                asset.image_format = str(image.format or "").upper()
                asset.mode = image.mode
        except (UnidentifiedImageError, OSError) as exc:
            raise SystemExit(f"Cannot read image: {exc}") from exc

        asset_map = {asset.media_path: asset}
        assign_image_plan(
            asset_map,
            image_profile,
            mb_to_bytes(target_size_mb) if target_mode else None,
        )
        if progress_callback is not None:
            progress_callback(0, 1, "正在处理图片 1/1")

        def encode_image_attempt() -> None:
            encode_image_asset(asset, output_dir)
            audit_encoded_assets(
                {},
                asset_map,
                video_threshold=float(getattr(args, "video_ssim_threshold", 0.95)),
                image_threshold=float(getattr(args, "image_ssim_threshold", 0.99)),
                forced=getattr(args, "quality_mode", "safe") == "forced",
                logger=logger,
                restore_failed=False,
            )
            shutil.copy2(Path(asset.output_path), output_path)
            asset.output_path = str(output_path)

        encode_image_attempt()
        capacity_attempts = [
            {
                "kind": "initial",
                "media_budget_bytes": (
                    mb_to_bytes(target_size_mb) if target_mode else asset.target_bytes
                ),
                "actual_bytes": output_path.stat().st_size,
            }
        ]
        if target_mode:
            media_budget = mb_to_bytes(target_size_mb)
            correction_rounds = 0
            giveback_used = False
            while True:
                next_attempt = next_target_media_budget(
                    actual_bytes=output_path.stat().st_size,
                    target_bytes=mb_to_bytes(target_size_mb),
                    current_media_budget=media_budget,
                    maximum_media_budget=source_size_bytes,
                    correction_rounds=correction_rounds,
                    giveback_used=giveback_used,
                )
                if next_attempt is None:
                    break
                media_budget, attempt_kind = next_attempt
                correction_rounds += attempt_kind == "correction"
                giveback_used |= attempt_kind == "quality_giveback"
                previous_plan = media_plan_signature({}, asset_map)
                assign_image_plan(asset_map, image_profile, media_budget)
                if media_plan_signature({}, asset_map) == previous_plan:
                    logger("Target capacity retry skipped: media plan is unchanged")
                    break
                encode_image_attempt()
                capacity_attempts.append(
                    {
                        "kind": attempt_kind,
                        "media_budget_bytes": media_budget,
                        "actual_bytes": output_path.stat().st_size,
                    }
                )
        write_standalone_report(
            artifact_report_path,
            input_path,
            output_path,
            image_asset=asset,
            target_size_mb=target_size_mb,
            capacity_attempts=capacity_attempts,
        )
        if progress_callback is not None:
            progress_callback(1, 1, "压缩完成")
        logger(
            f"{asset.media_path}: "
            f"{asset.zip_size / 1024 / 1024:.2f} MiB -> "
            f"{output_path.stat().st_size / 1024 / 1024:.2f} MiB"
        )
        logger("")
        logger(f"Output media: {output_path}")
        logger(f"Report: {artifact_report_path}")
        return {
            "input_pptx": input_path,
            "output_pptx": output_path,
            "report_path": artifact_report_path,
            "video_output_dir": None,
            "skipped": False,
        }
    finally:
        if not args.keep_work_dir and args.work_dir is None:
            shutil.rmtree(work_dir, ignore_errors=True)


_DOCUMENT_BACKEND_SUFFIXES = {
    ".docx": ("docx_image_compactor", "compact_docx"),
    ".docm": ("docx_image_compactor", "compact_docx"),
    ".xlsx": ("xlsx_image_compactor", "compact_xlsx"),
    ".xlsm": ("xlsx_image_compactor", "compact_xlsx"),
    ".pdf": ("pdf_image_compactor", "compact_pdf"),
}


def _compact_document_backend(
    source: Path,
    args: argparse.Namespace,
    logger: Logger,
    cancel_callback: CancelCallback | None,
) -> dict[str, Any]:
    """Route DOCX/XLSX/PDF inputs to their image-compactor backends.

    Backends import heavily from this module, so they are loaded lazily here
    to avoid a circular import at module load time. Validation failures raise
    ValueError (never SystemExit) so GUI worker threads can surface them.
    """
    if cancel_callback is not None and cancel_callback():
        raise CancelledError("Cancelled")
    if args.target_size_mb is None:
        raise ValueError(
            "DOCX/XLSX/PDF compression requires an explicit target size (MB)"
        )
    if args.image_profile == "none":
        raise ValueError(
            "DOCX/XLSX/PDF compression only re-encodes embedded images; "
            "image profile 'none' is not supported"
        )
    module_name, func_name = _DOCUMENT_BACKEND_SUFFIXES[source.suffix.lower()]
    module = importlib.import_module(module_name)
    compact_func = getattr(module, func_name)
    forced = args.quality_mode == "forced"
    safe_output = (
        module.default_output_path(source, args.target_size_mb, forced=False)
        if forced
        else None
    )
    result = compact_func(
        source,
        args.target_size_mb,
        output=args.output,
        image_profile=args.image_profile,
        image_ssim_threshold=args.image_ssim_threshold,
        forced=forced,
        safe_output=safe_output,
        confirm_forced=forced,
        logger=logger,
    )
    skipped = bool(result.get("skipped"))
    return {
        "input_pptx": result["input"],
        "output_pptx": result["output"],
        "report_path": result.get("report_path"),
        "skipped": skipped,
        "reason": "already_meets_target" if skipped else "",
    }


def compact_input_path(
    args: argparse.Namespace,
    logger: Logger = print,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    asset_pre_encoded: dict[str, str] | None = None,
) -> dict[str, Any]:
    input_path = Path(args.input_pptx).expanduser().resolve()
    if input_path.suffix.lower() == ".pptx":
        return compact_pptx(
            args,
            logger=logger,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
            asset_pre_encoded=asset_pre_encoded,
        )
    if is_supported_standalone_video(input_path):
        return compact_standalone_video(
            args,
            logger=logger,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
        )
    if is_supported_standalone_image(input_path):
        return compact_standalone_image(
            args,
            logger=logger,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
        )
    if input_path.suffix.lower() in _DOCUMENT_BACKEND_SUFFIXES:
        return _compact_document_backend(
            input_path,
            args,
            logger=logger,
            cancel_callback=cancel_callback,
        )
    raise SystemExit(
        f"Unsupported input file type: {input_path.suffix or input_path.name}"
    )


def compact_failed_assets_into_output(
    args: argparse.Namespace,
    failed_media_paths: set[str],
    base_output_pptx: Path,
    report_path: Path,
    logger: Logger = print,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
) -> dict[str, Any]:
    video_profile = getattr(args, "profile", "high") or "high"
    image_profile = getattr(args, "image_profile", "high") or "high"
    input_pptx = Path(args.input_pptx).expanduser().resolve()
    base_output_pptx = Path(base_output_pptx).expanduser().resolve()
    report_path = Path(report_path).expanduser().resolve()
    if not input_pptx.exists():
        raise SystemExit(f"Input file not found: {input_pptx}")
    if not base_output_pptx.exists():
        raise SystemExit(f"Previous compressed file not found: {base_output_pptx}")
    if not report_path.exists():
        raise SystemExit(f"Report JSON not found: {report_path}")
    if not failed_media_paths:
        return {
            "input_pptx": input_pptx,
            "output_pptx": base_output_pptx,
            "report_path": report_path,
            "video_output_dir": None,
            "skipped": True,
            "reason": "No failed assets to optimize.",
        }

    report_data = load_json_file(report_path, source="Report")
    output_pptx = (
        Path(report_data.get("output_pptx") or base_output_pptx).expanduser().resolve()
    )
    if not output_pptx.exists():
        output_pptx = base_output_pptx

    if progress_callback is not None:
        progress_callback(0, 1, "正在解析演示文稿")

    if video_profile != "none":
        ensure_binary("ffmpeg")
    runtime_config = load_runtime_config(args.config, video_profile)

    work_dir = (
        args.work_dir.expanduser().resolve()
        if args.work_dir
        else Path(tempfile.mkdtemp(prefix=runtime_temp_prefix("pptx_compact_")))
    )
    work_dir.mkdir(parents=True, exist_ok=True)
    video_output_dir = work_dir / "compressed_videos"
    image_output_dir = work_dir / "compressed_images"
    video_output_dir.mkdir(parents=True, exist_ok=True)
    image_output_dir.mkdir(parents=True, exist_ok=True)

    def check_cancelled() -> None:
        if cancel_callback is not None and cancel_callback():
            raise CancelledError("Cancelled")

    try:
        check_cancelled()
        assets, image_assets, meta = parse_pptx_assets(
            input_pptx,
            render_width=args.slide_render_width,
            render_height=args.slide_render_height,
            overscan=args.overscan,
            min_height=args.min_height,
            max_height=min(
                args.max_height, int(PROFILE_QUALITY_RULES[video_profile]["max_height"])
            ),
            config=runtime_config,
        )
        assets = {
            path: asset for path, asset in assets.items() if path in failed_media_paths
        }
        image_assets = {
            path: asset
            for path, asset in image_assets.items()
            if path in failed_media_paths
        }
        if not assets and not image_assets:
            reason = "No failed media assets were found in the source PPTX."
            logger(reason)
            return {
                "input_pptx": input_pptx,
                "output_pptx": output_pptx,
                "report_path": report_path,
                "video_output_dir": None,
                "skipped": True,
                "reason": reason,
            }

        old_video_entries = {
            item.get("media_path"): item
            for item in report_data.get("videos", [])
            if item.get("media_path")
        }
        old_output_paths = {
            media_path: item.get("output_media_path") or media_path
            for media_path, item in old_video_entries.items()
        }
        for media_path, asset in assets.items():
            asset.output_media_path = old_output_paths.get(
                media_path, asset.output_media_path or media_path
            )

        total_steps = len(assets) + len(image_assets) + 5
        logger(
            f"增量提档：仅处理 {len(assets)} 个视频、{len(image_assets)} 张图片，"
            f"并覆盖上一版压缩 PPTX。"
        )

        check_cancelled()
        if assets and progress_callback is not None:
            progress_callback(1, total_steps, "正在提取内嵌视频")
        if assets:
            ensure_binary("ffprobe")
            extract_videos(
                input_pptx, assets, work_dir, progress_callback=progress_callback
            )

        check_cancelled()
        if image_assets and progress_callback is not None:
            progress_callback(2, total_steps, "正在提取内嵌图片")
        if image_assets:
            extract_images(
                input_pptx, image_assets, work_dir, progress_callback=progress_callback
            )

        check_cancelled()
        if progress_callback is not None:
            progress_callback(3, total_steps, "正在计算压缩计划")
        if assets:
            if video_profile == "none":
                assign_video_copy_plan(assets)
            else:
                assign_profile_plan(
                    assets, video_profile, args.min_height, runtime_config
                )
                for asset in assets.values():
                    asset.status = "would_copy" if should_copy(asset) else "planned"
        if image_assets:
            assign_image_plan(image_assets, image_profile)

        current_video_bytes = sum(asset.zip_size for asset in assets.values())
        current_image_bytes = sum(asset.zip_size for asset in image_assets.values())
        non_media_bytes = (
            input_pptx.stat().st_size - current_video_bytes - current_image_bytes
        )
        print_plan(
            assets,
            image_assets,
            input_pptx,
            None,
            None,
            non_media_bytes,
            runtime_config,
            logger=logger,
        )

        total_asset_bytes = current_video_bytes + current_image_bytes
        base_weight = max(1000.0, float(total_asset_bytes) * 0.1)
        total_weight = total_asset_bytes + (base_weight * 2.0)
        current_weight = base_weight

        sorted_assets = sorted(assets.values(), key=lambda item: item.media_path)
        for index, asset in enumerate(sorted_assets, start=1):
            check_cancelled()
            label = f"正在处理视频 {index}/{len(sorted_assets)}: {Path(asset.media_path).name}"
            if progress_callback is not None:
                progress_callback(current_weight, total_weight, label)

            def asset_progress(f: float) -> None:
                if progress_callback is not None:
                    progress_callback(
                        current_weight + (asset.zip_size * f), total_weight, label
                    )

            encode_asset(
                asset,
                video_output_dir,
                args.preset,
                runtime_config,
                encoder_mode=args.encoder,
                logger=logger,
                cancel_callback=cancel_callback,
                progress_callback=asset_progress,
            )
            current_weight += asset.zip_size

        sorted_image_assets = sorted(
            image_assets.values(), key=lambda item: item.media_path
        )
        for index, asset in enumerate(sorted_image_assets, start=1):
            check_cancelled()
            label = f"正在处理图片 {index}/{len(sorted_image_assets)}: {Path(asset.media_path).name}"
            if progress_callback is not None:
                progress_callback(current_weight, total_weight, label)
            encode_image_asset(asset, image_output_dir)
            current_weight += asset.zip_size

        replacements: dict[str, Path] = {}
        replacement_infos: dict[str, ZipInfo] = {}
        relationship_path_map: dict[str, str] = {}
        remove_paths: set[str] = set()
        with ZipFile(input_pptx, "r") as source_zip:
            for asset in assets.values():
                new_media_path = asset.output_media_path or asset.media_path
                old_media_path = old_output_paths.get(asset.media_path, new_media_path)
                replacements[new_media_path] = Path(asset.output_path)
                try:
                    replacement_infos[new_media_path] = source_zip.getinfo(
                        asset.media_path
                    )
                except KeyError:
                    pass
                if old_media_path != new_media_path:
                    relationship_path_map[old_media_path] = new_media_path
                    remove_paths.add(old_media_path)
            for asset in image_assets.values():
                replacements[asset.media_path] = Path(asset.output_path)
                try:
                    replacement_infos[asset.media_path] = source_zip.getinfo(
                        asset.media_path
                    )
                except KeyError:
                    pass

        check_cancelled()
        if progress_callback is not None:
            progress_callback(
                total_weight - base_weight, total_weight, "正在重新打包 PPTX"
            )
        patch_output_pptx(
            output_pptx,
            output_pptx,
            replacements,
            relationship_path_map=relationship_path_map,
            replacement_infos=replacement_infos,
            remove_paths=remove_paths,
            video_assets=assets,
        )
        update_report_entries(
            report_path,
            report_data,
            input_pptx,
            output_pptx,
            assets,
            image_assets,
            meta,
            runtime_config,
        )

        if progress_callback is not None:
            progress_callback(total_weight, total_weight, "压缩完成")
        logger("")
        logger(f"Output PPTX: {output_pptx}")
        logger(f"Report: {report_path}")
        logger("增量提档完成：已在上一版压缩 PPTX 内原地替换低分素材。")
        return {
            "input_pptx": input_pptx,
            "output_pptx": output_pptx,
            "report_path": report_path,
            "video_output_dir": None,
            "skipped": False,
        }
    finally:
        if not args.keep_work_dir and args.work_dir is None:
            shutil.rmtree(work_dir, ignore_errors=True)


def main() -> int:
    args = parse_args()
    input_paths = [path.expanduser().resolve() for path in args.input_pptx]
    if len(input_paths) > 1 and args.output is not None:
        raise SystemExit("--output can only be used with one input file")
    if len(input_paths) > 1 and args.video_output_dir is not None:
        raise SystemExit("--video-output-dir can only be used with one input file")

    for input_pptx in input_paths:
        single_args = argparse.Namespace(
            **{
                **vars(args),
                "input_pptx": input_pptx,
                "target_size_mb": args.target_size_mb,
                "output": args.output if len(input_paths) == 1 else None,
                "video_output_dir": args.video_output_dir
                if len(input_paths) == 1
                else None,
            }
        )
        compact_input_path(single_args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
