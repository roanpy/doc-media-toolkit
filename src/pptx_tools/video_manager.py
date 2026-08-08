from __future__ import annotations

import argparse
import copy
import hashlib
import json
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import uuid
from array import array
from collections.abc import Callable, Iterable, Sequence
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import Any
from zipfile import BadZipFile, ZipFile

from PIL import Image

from pptx_output_watermark.ffmpeg_runtime import ensure_binary, run_binary
from pptx_output_watermark.process_utils import hidden_console_kwargs, run_process
from pptx_output_watermark.pptx_video_support import VideoAsset, scan_embedded_videos
from pptx_tools.project_lock import project_write_lock
from pptx_video_compactor import (
    VideoAsset as CompactVideoAsset,
    choose_output_media_path,
    compact_standalone_video,
    media_needs_mp4,
    patch_output_pptx,
)


SCHEMA_VERSION = 1
MANIFEST_NAME = "video-project.json"
BACKUP_MANIFEST_NAME = "video-project.json.bak"
LOCK_NAME = ".video-project.lock"
CLEANUP_LOCK_NAME = ".video-cleanup.lock"
FAMILY_MOVE_JOURNAL_NAME = ".video-family-move.json"
LOGGER = logging.getLogger("pptx_tools.video_manager")
ProgressCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]
PlaceholderBuilder = Callable[[Path, Path, dict[str, Any]], None]
VIDEO_SUFFIXES = {".mp4", ".mov", ".m4v", ".wmv", ".avi", ".mkv", ".webm"}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _sha256_file_prefix(path: Path, length: int) -> str:
    digest = hashlib.sha256()
    remaining = length
    with path.open("rb") as handle:
        while remaining > 0:
            chunk = handle.read(min(1024 * 1024, remaining))
            if not chunk:
                break
            digest.update(chunk)
            remaining -= len(chunk)
    if remaining:
        return ""
    return digest.hexdigest()


def _sha256_zip_member(archive: ZipFile, member: str) -> str:
    digest = hashlib.sha256()
    with archive.open(member, "r") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _copy_zip_member(archive: ZipFile, member: str, target: Path) -> str:
    target.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256()
    with archive.open(member, "r") as source, target.open("wb") as output:
        while chunk := source.read(1024 * 1024):
            digest.update(chunk)
            output.write(chunk)
    return digest.hexdigest()


def _safe_name(value: str, fallback: str = "video") -> str:
    value = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("._")
    return value[:80] or fallback


def _normalized_video_name(value: str) -> str:
    name = value.strip().casefold()
    while Path(name).suffix in VIDEO_SUFFIXES:
        name = Path(name).stem.strip()
    return name


def _normalized_cleanup_name(value: str) -> str:
    name = _normalized_video_name(value)
    return re.sub(
        r"(?:[\s._-]+(?:copy|副本)|\s*\(\d+\)|[_-]\d+)$",
        "",
        name,
        flags=re.IGNORECASE,
    ).strip(" ._-")


def _video_stem(value: str) -> str:
    value = value.strip()
    return Path(value).stem if Path(value).suffix.lower() in VIDEO_SUFFIXES else value


def _variant_filename(
    name: str, metadata: dict[str, Any], digest: str, suffix: str
) -> str:
    name = re.sub(
        r"_\[[^\]]+\]_[0-9a-f]{8}$", "", _video_stem(name), flags=re.IGNORECASE
    )
    width = int(metadata.get("width") or 0)
    height = int(metadata.get("height") or 0)
    duration = float(metadata.get("duration_sec") or 0)
    details = []
    if width > 0 and height > 0:
        details.append(f"{width}x{height}")
    if duration > 0:
        details.append(f"{duration:.1f}s")
    spec = f"_[{'_'.join(details)}]" if details else ""
    return f"{_safe_name(name)}{spec}_{digest[:8]}{suffix}"


def normalize_library_category(value: str) -> Path:
    raw = value.strip().replace("\\", "/")
    if not raw:
        return Path()
    if raw.startswith("/") or re.match(r"^[A-Za-z]:", raw):
        raise ValueError("Video library category must be a relative folder path")
    raw = raw.rstrip("/")
    if any(part in {"", ".", ".."} for part in raw.split("/")):
        raise ValueError("Video library category must be a relative folder path")
    parts = [_safe_name(part, "") for part in raw.split("/")]
    if any(not part for part in parts):
        raise ValueError("Video library category contains an invalid folder name")
    return Path(*parts)


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 10_000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Unable to choose a unique path near {path}")


def _check_cancelled(cancel_callback: CancelCallback | None) -> None:
    if cancel_callback is not None and cancel_callback():
        raise RuntimeError("Operation cancelled")


def probe_video(path: Path) -> dict[str, Any]:
    try:
        result = run_binary(
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
        payload = json.loads(result.stdout or "{}")
        streams = payload.get("streams", [])
        video = next(
            (stream for stream in streams if stream.get("codec_type") == "video"), {}
        )
        audio = next(
            (stream for stream in streams if stream.get("codec_type") == "audio"), None
        )
        format_data = payload.get("format", {})
        duration = video.get("duration") or format_data.get("duration") or 0
        bitrate = format_data.get("bit_rate") or video.get("bit_rate") or 0
        return {
            "width": int(video.get("width") or 0),
            "height": int(video.get("height") or 0),
            "duration_sec": round(float(duration or 0), 3),
            "bitrate_kbps": int(float(bitrate or 0)) // 1000,
            "video_codec": str(video.get("codec_name") or ""),
            "audio_codec": str(audio.get("codec_name") or "") if audio else "",
            "has_audio": audio is not None,
            "probe_error": "",
        }
    except Exception as exc:
        LOGGER.warning("Unable to probe video %s: %s", path, exc)
        return {
            "width": 0,
            "height": 0,
            "duration_sec": 0.0,
            "bitrate_kbps": 0,
            "video_codec": "",
            "audio_codec": "",
            "has_audio": False,
            "probe_error": str(exc),
        }


def _video_packet_error(path: Path) -> str:
    """Fast whole-file integrity check without re-encoding the media."""
    try:
        result = run_binary(
            [
                "ffmpeg",
                "-v",
                "error",
                "-i",
                str(path),
                "-map",
                "0:v:0",
                "-map",
                "0:a:0?",
                "-c",
                "copy",
                "-f",
                "null",
                "-",
            ],
            capture=True,
        )
        return str(result.stderr or "").strip()
    except Exception as exc:
        return str(exc)


def _aligned_file_similarity(
    first: Path,
    second: Path,
    *,
    sample_count: int = 48,
    sample_size: int = 64 * 1024,
) -> float:
    """Compare aligned samples, including truncated copies of a larger file."""
    length = min(first.stat().st_size, second.stat().st_size)
    if length <= 0:
        return 0.0
    window = min(sample_size, length)
    positions = (
        [0]
        if length == window
        else [
            round(index * (length - window) / (sample_count - 1))
            for index in range(sample_count)
        ]
    )
    matched = 0
    compared = 0
    with first.open("rb") as left, second.open("rb") as right:
        for position in positions:
            left.seek(position)
            right.seek(position)
            left_chunk = left.read(window)
            right_chunk = right.read(window)
            compared += min(len(left_chunk), len(right_chunk))
            matched += sum(a == b for a, b in zip(left_chunk, right_chunk))
    return matched / compared if compared else 0.0


def _decoded_audio_correlation(first: Path, second: Path) -> float | None:
    def samples(path: Path) -> array:
        result = run_process(
            [
                ensure_binary("ffmpeg"),
                "-v",
                "error",
                "-i",
                str(path),
                "-map",
                "0:a:0",
                "-ac",
                "1",
                "-ar",
                "500",
                "-f",
                "s16le",
                "-",
            ],
            capture_output=True,
            timeout=300,
            check=True,
            **hidden_console_kwargs(),
        )
        values = array("h")
        values.frombytes(result.stdout)
        return values

    try:
        left = samples(first)
        right = samples(second)
    except Exception:
        return None
    length = min(len(left), len(right))
    if not length:
        return None
    step = max(1, length // 20_000)
    left_values = left[:length:step]
    right_values = right[:length:step]
    left_mean = sum(left_values) / len(left_values)
    right_mean = sum(right_values) / len(right_values)
    numerator = sum(
        (a - left_mean) * (b - right_mean) for a, b in zip(left_values, right_values)
    )
    left_energy = sum((value - left_mean) ** 2 for value in left_values)
    right_energy = sum((value - right_mean) ** 2 for value in right_values)
    if not left_energy or not right_energy:
        return 1.0 if left_energy == right_energy else 0.0
    return numerator / math.sqrt(left_energy * right_energy)


BACKFILL_QUALITY_TIERS: dict[str, dict[str, Any]] = {
    # 上限语义：族源分辨率/码率均在档位上限内且容器编码兼容 → 原样嵌入；
    # 任一超限 → 按档位参数转码。bitrate_kbps = 0 表示不设限。
    "best": {
        "label": "最佳",
        "max_width": 1920,
        "max_height": 1080,
        "crf": 18,
        "bitrate_kbps": 0,
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


def _video_envelope(
    width: int, height: int, max_width: int = 1920, max_height: int = 1080
) -> tuple[int, int]:
    return (max_width, max_height) if width >= height else (max_height, max_width)


def _archived_dimensions(
    width: int, height: int, max_width: int = 1920, max_height: int = 1080
) -> tuple[int, int]:
    limit_width, limit_height = _video_envelope(width, height, max_width, max_height)
    scale = min(1.0, limit_width / width, limit_height / height)
    return (
        max(2, int(width * scale) // 2 * 2),
        max(2, int(height * scale) // 2 * 2),
    )


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
        return True  # No audio stream: the optional -map 0:a? is a no-op anyway.
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
    command = [
        "ffmpeg",
        "-y",
        "-i",
        str(source),
        "-map",
        "0:v:0",
    ]
    if has_audio:
        command += ["-map", "0:a?"]
    command += [
        "-vf",
        f"scale={output_width}:{output_height}:flags=lanczos",
        # Backfill never reduces the frame rate, so preserve VFR timestamps
        # instead of letting FFmpeg silently drop frames (ffmpeg >= 5.1).
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
        command += [
            "-maxrate",
            f"{bitrate_kbps}k",
            "-bufsize",
            f"{bitrate_kbps * 2}k",
        ]
    if has_audio:
        command += ["-c:a", "aac", "-b:a", audio_bitrate]
    command += ["-movflags", "+faststart"]
    if family_id:
        command.extend(["-metadata", f"comment=pptx-tools-family:{family_id}"])
    run_binary([*command, str(target)], capture=True)


def _video_fingerprint(
    path: Path, metadata: dict[str, Any] | None = None
) -> dict[str, Any] | None:
    metadata = metadata or probe_video(path)
    duration = float(metadata.get("duration_sec") or 0)
    width = int(metadata.get("width") or 0)
    height = int(metadata.get("height") or 0)
    if duration <= 0 or width <= 0 or height <= 0:
        return None
    work = Path(tempfile.mkdtemp(prefix="video-fingerprint-"))
    try:
        run_binary(
            [
                "ffmpeg",
                "-y",
                "-v",
                "error",
                "-i",
                str(path),
                "-vf",
                f"fps={5 / duration:.8f},scale=9:8:flags=area,format=gray",
                "-frames:v",
                "5",
                str(work / "frame-%02d.png"),
            ],
            capture=True,
        )
        hashes: list[str] = []
        luma: list[int] = []
        for frame in sorted(work.glob("frame-*.png")):
            with Image.open(frame) as image:
                resized = image.convert("L").resize((9, 8))
                pixels = list(
                    resized.get_flattened_data()
                    if hasattr(resized, "get_flattened_data")
                    else resized.getdata()
                )
            luma.append(round(sum(pixels) / len(pixels)))
            bits = 0
            for row in range(8):
                for column in range(8):
                    bits = (bits << 1) | int(
                        pixels[row * 9 + column] > pixels[row * 9 + column + 1]
                    )
            hashes.append(f"{bits:016x}")
        if len(hashes) < 4:
            return None
        fingerprint: dict[str, Any] = {
            "duration_ms": round(duration * 1000),
            "aspect_ppm": round(width / height * 1_000_000),
            "frames": hashes,
            "luma": luma,
            "has_audio": bool(metadata.get("has_audio")),
        }
        if fingerprint["has_audio"]:
            try:
                spectrum = work / "audio-spectrum.png"
                run_binary(
                    [
                        "ffmpeg",
                        "-y",
                        "-v",
                        "error",
                        "-i",
                        str(path),
                        "-filter_complex",
                        "[0:a:0]showspectrumpic=s=64x64:legend=disabled:color=channel[v]",
                        "-map",
                        "[v]",
                        "-frames:v",
                        "1",
                        str(spectrum),
                    ],
                    capture=True,
                )
                with Image.open(spectrum) as image:
                    resized = image.convert("L").resize(
                        (9, 8), Image.Resampling.LANCZOS
                    )
                    pixels = list(
                        resized.get_flattened_data()
                        if hasattr(resized, "get_flattened_data")
                        else resized.getdata()
                    )
                bits = 0
                for row in range(8):
                    for column in range(8):
                        bits = (bits << 1) | int(
                            pixels[row * 9 + column] > pixels[row * 9 + column + 1]
                        )
                fingerprint["audio_hash"] = f"{bits:016x}"
                fingerprint["audio_luma"] = round(sum(pixels) / len(pixels))
            except Exception:
                fingerprint["has_audio"] = False
        return fingerprint
    except Exception:
        return None
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _fingerprint_int(fingerprint: dict[str, Any], key: str) -> int | None:
    try:
        return int(fingerprint[key])
    except (KeyError, TypeError, ValueError):
        return None


def _visual_fingerprint_matches(
    distances: list[int], luma_differences: list[int]
) -> bool:
    if len(distances) < 4 or len(luma_differences) != len(distances):
        return False
    if max(distances) <= 6 and sum(distances) <= 20 and max(luma_differences) <= 10:
        return True
    if len(distances) < 5:
        return False
    # Re-encoding can move one sampled frame across a scene boundary.
    return (
        max(sorted(distances)[:-1]) <= 4
        and sum(sorted(distances)[:-1]) <= 12
        and max(sorted(luma_differences)[:-1]) <= 10
    )


def _fingerprints_match(first: dict[str, Any], second: dict[str, Any]) -> bool:
    # Corrupted or legacy fingerprints must degrade to "no match", never crash
    # and never widen into a false positive.
    try:
        first_duration = _fingerprint_int(first, "duration_ms")
        second_duration = _fingerprint_int(second, "duration_ms")
        first_aspect = _fingerprint_int(first, "aspect_ppm")
        second_aspect = _fingerprint_int(second, "aspect_ppm")
        if (
            first_duration is None
            or second_duration is None
            or first_aspect is None
            or second_aspect is None
        ):
            return False
        duration = max(first_duration, second_duration, 1)
        duration_tolerance = min(250, max(150, round(duration * 0.001)))
        if abs(first_duration - second_duration) > duration_tolerance:
            return False
        aspect = max(first_aspect, second_aspect, 1)
        # Encoders commonly crop 1080 to 1072 pixels at macroblock boundaries.
        if abs(first_aspect - second_aspect) > aspect * 0.015:
            return False
        left = first.get("frames", [])
        right = second.get("frames", [])
        if len(left) != len(right) or len(left) < 4:
            return False
        distances = [(int(a, 16) ^ int(b, 16)).bit_count() for a, b in zip(left, right)]
        left_luma = first.get("luma", [])
        right_luma = second.get("luma", [])
        luma_differences = [abs(int(a) - int(b)) for a, b in zip(left_luma, right_luma)]
        visual_match = (
            len(left_luma) == len(left)
            and len(right_luma) == len(right)
            and _visual_fingerprint_matches(distances, luma_differences)
        )
        if not visual_match:
            return False

        first_audio = first.get("has_audio")
        second_audio = second.get("has_audio")
        if first_audio is None and second_audio is None:
            return True
        if bool(first_audio) != bool(second_audio):
            return False
        if not first_audio:
            return True
        first_hash = first.get("audio_hash")
        second_hash = second.get("audio_hash")
        if not isinstance(first_hash, str) or not isinstance(second_hash, str):
            return False
        audio_distance = (int(first_hash, 16) ^ int(second_hash, 16)).bit_count()
        return audio_distance <= 13
    except (TypeError, ValueError, AttributeError):
        return False


def _ssim_videos(
    candidate: Path, reference: Path, *, timeout_seconds: int = 300
) -> float | None:
    """SSIM of `candidate` against `reference` (candidate is scaled to the
    reference canvas). Returns None when ffmpeg cannot score the pair."""
    from pptx_output_watermark.process_utils import (
        finish_process,
        kill_process,
        start_process,
    )
    from pptx_video_compactor import hidden_subprocess_kwargs, resolve_binary

    ffmpeg = resolve_binary("ffmpeg")
    if ffmpeg is None:
        return None
    filter_cmd = (
        "[0:v]setpts=PTS-STARTPTS,fps=1[v0];"
        "[1:v]setpts=PTS-STARTPTS,fps=1[v1];"
        "[v0][v1]scale2ref[v0s][v1s];[v0s][v1s]ssim"
    )
    process = start_process(
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
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
        errors="replace",
        **hidden_subprocess_kwargs(),
    )
    try:
        try:
            _, stderr = process.communicate(timeout=timeout_seconds)
        except subprocess.TimeoutExpired:
            kill_process(process)
            process.communicate()
            return None
    finally:
        finish_process(process)
    if process.returncode:
        return None
    match = re.search(r"All:([0-9.]+)", stderr or "")
    return float(match.group(1)) if match else None


def _fingerprint_confidence(
    first: dict[str, Any], second: dict[str, Any]
) -> dict[str, Any]:
    """Distance breakdown between two fingerprints. `matched` follows the same
    conservative rule as `_fingerprints_match`; `level` is high when every
    distance is comfortably inside the threshold, medium when it barely makes
    it, and low when the pair does not match."""
    result: dict[str, Any] = {
        "matched": False,
        "level": "low",
        "duration_consistent": False,
        "audio_consistent": False,
        "duration_diff_ms": None,
        "frame_max_distance": None,
        "frame_total_distance": None,
        "luma_max_difference": None,
        "aspect_consistent": False,
        "audio_distance": None,
    }
    try:
        first_duration = _fingerprint_int(first, "duration_ms")
        second_duration = _fingerprint_int(second, "duration_ms")
        first_aspect = _fingerprint_int(first, "aspect_ppm")
        second_aspect = _fingerprint_int(second, "aspect_ppm")
        if (
            first_duration is None
            or second_duration is None
            or first_aspect is None
            or second_aspect is None
        ):
            return result
        duration_diff = abs(first_duration - second_duration)
        duration = max(first_duration, second_duration, 1)
        duration_tolerance = min(250, max(150, round(duration * 0.001)))
        result["duration_diff_ms"] = duration_diff
        result["duration_consistent"] = duration_diff <= duration_tolerance
        aspect = max(first_aspect, second_aspect, 1)
        aspect_ok = abs(first_aspect - second_aspect) <= aspect * 0.015
        result["aspect_consistent"] = aspect_ok

        left = first.get("frames", [])
        right = second.get("frames", [])
        if len(left) != len(right) or len(left) < 4:
            return result
        distances = [(int(a, 16) ^ int(b, 16)).bit_count() for a, b in zip(left, right)]
        frame_max = max(distances)
        result["frame_max_distance"] = frame_max
        result["frame_total_distance"] = sum(distances)
        left_luma = first.get("luma", [])
        right_luma = second.get("luma", [])
        luma_differences = [abs(int(a) - int(b)) for a, b in zip(left_luma, right_luma)]
        if luma_differences:
            result["luma_max_difference"] = max(luma_differences)
        visual_ok = (
            len(left_luma) == len(left)
            and len(right_luma) == len(right)
            and _visual_fingerprint_matches(distances, luma_differences)
        )

        first_audio = first.get("has_audio")
        second_audio = second.get("has_audio")
        audio_distance: int | None = None
        if first_audio is None and second_audio is None:
            audio_ok = True
        elif bool(first_audio) != bool(second_audio):
            audio_ok = False
        elif not first_audio:
            audio_ok = True
        else:
            first_hash = first.get("audio_hash")
            second_hash = second.get("audio_hash")
            if not isinstance(first_hash, str) or not isinstance(second_hash, str):
                audio_ok = False
            else:
                audio_distance = (
                    int(first_hash, 16) ^ int(second_hash, 16)
                ).bit_count()
                result["audio_distance"] = audio_distance
                audio_ok = audio_distance <= 13
        result["audio_consistent"] = audio_ok
        matched = result["duration_consistent"] and aspect_ok and visual_ok and audio_ok
        result["matched"] = matched
        if matched:
            comfortable = (
                duration_diff <= duration_tolerance / 2
                and frame_max <= 2
                and (audio_distance is None or audio_distance <= 6)
            )
            result["level"] = "high" if comfortable else "medium"
        return result
    except (TypeError, ValueError, AttributeError):
        return result


def create_video_thumbnail(source: Path, target: Path) -> bool:
    """Create a three-frame review strip without changing the source video."""
    try:
        metadata = probe_video(source)
        duration = float(metadata.get("duration_sec") or 0)
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
            prefix="video-review-strip-", dir=target.parent
        ) as temporary:
            frames: list[Image.Image] = []
            for index, ratio in enumerate((0.1, 0.5, 0.9)):
                frame_path = Path(temporary) / f"{index}.jpg"
                run_binary(
                    [
                        "ffmpeg",
                        "-y",
                        "-v",
                        "error",
                        "-ss",
                        f"{max(0.0, duration * ratio):.3f}",
                        "-i",
                        str(source),
                        "-frames:v",
                        "1",
                        "-vf",
                        "scale=320:-2:flags=lanczos",
                        str(frame_path),
                    ],
                    capture=True,
                )
                with Image.open(frame_path) as image:
                    frames.append(image.convert("RGB"))
            height = max(image.height for image in frames)
            strip = Image.new("RGB", (sum(image.width for image in frames), height))
            left = 0
            for image in frames:
                strip.paste(image, (left, (height - image.height) // 2))
                left += image.width
            strip.save(target, format="JPEG", quality=88, optimize=True)
        return target.is_file() and target.stat().st_size > 0
    except Exception:
        target.unlink(missing_ok=True)
        return False


def _named_cleanup_match(
    first_name: str,
    second_name: str,
    first: dict[str, Any],
    second: dict[str, Any],
) -> bool:
    left = _normalized_cleanup_name(first_name)
    right = _normalized_cleanup_name(second_name)
    if left != right or left in {"video", "normal_video", "视频", "媒体1", "movie"}:
        return False
    if left.startswith("media"):
        return False
    confidence = _fingerprint_confidence(first, second)
    frame_max = confidence["frame_max_distance"]
    frame_total = confidence["frame_total_distance"]
    luma_max = confidence["luma_max_difference"]
    first_aspect = _fingerprint_int(first, "aspect_ppm")
    second_aspect = _fingerprint_int(second, "aspect_ppm")
    aspect_close = bool(
        first_aspect
        and second_aspect
        and abs(first_aspect - second_aspect) <= max(first_aspect, second_aspect) * 0.03
    )
    return bool(
        confidence["duration_consistent"]
        and aspect_close
        and confidence["audio_consistent"]
        and frame_max is not None
        and frame_max <= 16
        and frame_total is not None
        and frame_total <= 35
        and luma_max is not None
        and luma_max <= 15
    )


class VideoProject:
    def __init__(self, root: Path, data: dict[str, Any]) -> None:
        self.root = root.expanduser().resolve()
        self.data = data
        self._revision = int(data.get("revision", 0))
        self.recovered_from_backup = False
        self.recovery_detail = ""
        self._recover_pending_cleanup_moves()

    @property
    def manifest_path(self) -> Path:
        return self.root / MANIFEST_NAME

    @property
    def backup_manifest_path(self) -> Path:
        return self.root / BACKUP_MANIFEST_NAME

    @property
    def family_move_journal_path(self) -> Path:
        return self.root / FAMILY_MOVE_JOURNAL_NAME

    @staticmethod
    def _read_manifest(path: Path) -> dict[str, Any]:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Video project manifest must be an object")
        if data.get("schema_version") != SCHEMA_VERSION:
            raise ValueError(
                f"Unsupported video project schema: {data.get('schema_version')}"
            )
        if not isinstance(data.get("project_id"), str):
            raise ValueError("Video project is missing project_id")
        if not isinstance(data.get("families"), list) or not isinstance(
            data.get("decks"), list
        ):
            raise ValueError("Video project families/decks must be lists")

        def validate_stored_path(value: str, label: str) -> None:
            stored = Path(value).expanduser()
            if stored.is_absolute():
                return
            try:
                (path.parent / stored).resolve().relative_to(path.parent.resolve())
            except ValueError:
                raise ValueError(
                    f"{label} escapes the video project directory"
                ) from None

        family_ids: set[str] = set()
        variant_ids: set[str] = set()
        variant_family_ids: dict[str, str] = {}
        for family in data["families"]:
            if not isinstance(family, dict):
                raise ValueError("Video project family must be an object")
            for key in ("id", "name", "active_variant_id"):
                if not isinstance(family.get(key), str) or not family[key]:
                    raise ValueError(f"Video family is missing {key}")
            if family["id"] in family_ids:
                raise ValueError(f"Duplicate video family id: {family['id']}")
            category = family.get("category", "")
            if not isinstance(category, str):
                raise ValueError("Video family category must be a string")
            normalize_library_category(category)
            family_ids.add(family["id"])
            variants = family.get("variants")
            if not isinstance(variants, list) or not variants:
                raise ValueError("Video family variants must be a non-empty list")
            family_variant_ids: set[str] = set()
            for variant in variants:
                if not isinstance(variant, dict):
                    raise ValueError("Video variant must be an object")
                for key in ("id", "path", "sha256"):
                    if not isinstance(variant.get(key), str) or not variant[key]:
                        raise ValueError(f"Video variant is missing {key}")
                if variant["id"] in variant_ids:
                    raise ValueError(f"Duplicate video variant id: {variant['id']}")
                validate_stored_path(variant["path"], "Video variant path")
                origin_paths = variant.get("origin_paths", [])
                if not isinstance(origin_paths, list) or not all(
                    isinstance(path, str) and path for path in origin_paths
                ):
                    raise ValueError("Video variant origin paths must be strings")
                variant_ids.add(variant["id"])
                variant_family_ids[variant["id"]] = family["id"]
                family_variant_ids.add(variant["id"])
            if family["active_variant_id"] not in family_variant_ids:
                raise ValueError("Video family active variant does not exist")
            source_variant_id = family.get("source_variant_id")
            if (
                source_variant_id is not None
                and source_variant_id not in family_variant_ids
            ):
                raise ValueError("Video family source variant does not exist")
            known_hashes = family.get("known_hashes", [])
            if not isinstance(known_hashes, list) or not all(
                isinstance(digest, str) and digest for digest in known_hashes
            ):
                raise ValueError("Video family known hashes must be strings")
            source_hashes = family.get("source_hashes", [])
            if not isinstance(source_hashes, list) or not all(
                isinstance(digest, str) and digest for digest in source_hashes
            ):
                raise ValueError("Video family source hashes must be strings")
            fingerprint = family.get("content_fingerprint")
            if fingerprint is not None and (
                not isinstance(fingerprint, dict)
                or not isinstance(fingerprint.get("duration_ms"), int)
                or not isinstance(fingerprint.get("aspect_ppm"), int)
                or not isinstance(fingerprint.get("frames"), list)
                or len(fingerprint["frames"]) < 4
                or not all(isinstance(item, str) for item in fingerprint["frames"])
                or not isinstance(fingerprint.get("luma"), list)
                or len(fingerprint["luma"]) != len(fingerprint["frames"])
                or not all(isinstance(item, int) for item in fingerprint["luma"])
            ):
                raise ValueError("Video family content fingerprint is invalid")
            if fingerprint is not None:
                has_audio = fingerprint.get("has_audio")
                if has_audio is not None and not isinstance(has_audio, bool):
                    raise ValueError("Video fingerprint has_audio must be boolean")
                audio_hash = fingerprint.get("audio_hash")
                audio_luma = fingerprint.get("audio_luma")
                if audio_hash is not None and not isinstance(audio_hash, str):
                    raise ValueError("Video fingerprint audio hash is invalid")
                if audio_luma is not None and not isinstance(audio_luma, int):
                    raise ValueError("Video fingerprint audio luma is invalid")

        deck_ids: set[str] = set()
        for deck in data["decks"]:
            if not isinstance(deck, dict):
                raise ValueError("PPTX deck must be an object")
            for key in ("id", "name", "source_path", "source_sha256"):
                if not isinstance(deck.get(key), str) or not deck[key]:
                    raise ValueError(f"PPTX deck is missing {key}")
            if deck["id"] in deck_ids:
                raise ValueError(f"Duplicate PPTX deck id: {deck['id']}")
            validate_stored_path(deck["source_path"], "PPTX source path")
            source_aliases = deck.get("source_aliases", [])
            if not isinstance(source_aliases, list) or not all(
                isinstance(item, str) and item for item in source_aliases
            ):
                raise ValueError("PPTX source aliases must be paths")
            for source_alias in source_aliases:
                validate_stored_path(source_alias, "PPTX source alias")
            deck_ids.add(deck["id"])
            assets = deck.get("assets")
            if not isinstance(assets, list):
                raise ValueError("PPTX deck assets must be a list")
            for asset in assets:
                if not isinstance(asset, dict):
                    raise ValueError("PPTX video asset must be an object")
                for key in (
                    "part_path",
                    "placeholder_part",
                    "family_id",
                    "original_variant_id",
                ):
                    if not isinstance(asset.get(key), str) or not asset[key]:
                        raise ValueError(f"PPTX video asset is missing {key}")
                if asset["family_id"] not in family_ids:
                    raise ValueError("PPTX video asset references an unknown family")
                if asset["original_variant_id"] not in variant_ids:
                    raise ValueError("PPTX video asset references an unknown variant")
                if (
                    variant_family_ids[asset["original_variant_id"]]
                    != asset["family_id"]
                ):
                    raise ValueError(
                        "PPTX video asset variant belongs to another family"
                    )
                if not isinstance(asset.get("occurrences"), list):
                    raise ValueError("PPTX video occurrences must be a list")
                for occurrence in asset["occurrences"]:
                    if not isinstance(occurrence, dict):
                        raise ValueError("PPTX video occurrence must be an object")
                    if not isinstance(
                        occurrence.get("slide_path"), str
                    ) or not isinstance(occurrence.get("shape_id"), int):
                        raise ValueError("PPTX video occurrence is missing its anchor")
            for key in (
                "detached_outputs",
                "restored_outputs",
                "optimized_outputs",
            ):
                records = deck.get(key, [])
                if not isinstance(records, list):
                    raise ValueError(f"PPTX {key} must be a list")
                for record in records:
                    if not isinstance(record, dict):
                        raise ValueError("PPTX output record must be an object")
                    for required in ("path", "sha256"):
                        if (
                            not isinstance(record.get(required), str)
                            or not record[required]
                        ):
                            raise ValueError(f"PPTX output is missing {required}")
                    validate_stored_path(record["path"], "PPTX output path")
        return data

    def _fill_record_metadata(self) -> None:
        self.data.setdefault("revision", self._revision)
        for family in self.families():
            source = self.source_variant(family)
            family.setdefault("source_variant_id", source["id"])
            family.setdefault(
                "known_hashes", [variant["sha256"] for variant in family["variants"]]
            )
            family.setdefault("source_hashes", [source["sha256"]])
        for deck in self.decks():
            deck.setdefault("optimized_outputs", [])
            deck.setdefault("source_aliases", [])
            source = self.deck_source_path(deck)
            if source.is_file():
                deck.setdefault("source_size_bytes", source.stat().st_size)
                deck.setdefault("source_mtime_ns", source.stat().st_mtime_ns)
            for kind, key in (
                ("lightweight", "detached_outputs"),
                ("restored", "restored_outputs"),
                ("optimized", "optimized_outputs"),
            ):
                for record in deck.setdefault(key, []):
                    record.setdefault("id", str(uuid.uuid4()))
                    record.setdefault("kind", kind)
                    path = self.resolve_path(record["path"])
                    if path.is_file():
                        record.setdefault("size_bytes", path.stat().st_size)
                        record.setdefault("mtime_ns", path.stat().st_mtime_ns)

    @classmethod
    def create(cls, root: Path, name: str | None = None) -> VideoProject:
        root = root.expanduser().resolve()
        root.mkdir(parents=True, exist_ok=True)
        manifest = root / MANIFEST_NAME
        if manifest.exists():
            return cls.open(root)
        now = utc_now()
        project = cls(
            root,
            {
                "schema_version": SCHEMA_VERSION,
                "project_id": str(uuid.uuid4()),
                "revision": 0,
                "name": name or root.name,
                "created_at": now,
                "updated_at": now,
                "families": [],
                "decks": [],
            },
        )
        for directory in (
            "media",
            "posters",
            "decks/detached",
            "decks/restored",
            "reports",
        ):
            (root / directory).mkdir(parents=True, exist_ok=True)
        project.save()
        project.record("project_created", name=project.data["name"])
        return project

    @classmethod
    def open(cls, root: Path) -> VideoProject:
        root = root.expanduser().resolve()
        manifest = root if root.name == MANIFEST_NAME else root / MANIFEST_NAME
        if manifest.name == MANIFEST_NAME and manifest.is_file():
            root = manifest.parent
        backup = root / BACKUP_MANIFEST_NAME
        primary_error: Exception | None = None
        try:
            data = cls._read_manifest(manifest)
        except Exception as exc:
            primary_error = exc
            try:
                data = cls._read_manifest(backup)
            except Exception:
                if not manifest.exists() and not backup.exists():
                    raise FileNotFoundError(
                        f"Video project manifest not found: {manifest}"
                    ) from None
                raise primary_error
            LOGGER.warning("Recovered video project from %s: %s", backup, exc)
            project = cls(root, data)
            project.recovered_from_backup = True
            project.recovery_detail = str(primary_error)
            project._fill_record_metadata()
            project.save(recover_invalid_current=True)
            project.data = cls._recover_pending_family_move(root)
            project._revision = int(project.data.get("revision", 0))
            project.record("manifest_recovered", backup=str(backup))
            return project
        data = cls._recover_pending_family_move(root, data)
        project = cls(root, data)
        project._fill_record_metadata()
        return project

    @classmethod
    def _recover_pending_family_move(
        cls,
        root: Path,
        data: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        journal_path = root / FAMILY_MOVE_JOURNAL_NAME
        if not journal_path.is_file():
            return (
                data if data is not None else cls._read_manifest(root / MANIFEST_NAME)
            )
        with project_write_lock(root, LOCK_NAME):
            current = cls._read_manifest(root / MANIFEST_NAME)
            try:
                journal = json.loads(journal_path.read_text(encoding="utf-8"))
                family_id = str(journal["family_id"])
                moves = journal["moves"]
                family = next(
                    item for item in current["families"] if item["id"] == family_id
                )
                variants = {item["id"]: item for item in family["variants"]}
                if not isinstance(moves, list) or set(variants) != {
                    str(item["variant_id"]) for item in moves
                }:
                    raise ValueError("Video family move journal does not match family")

                source_matches = True
                target_matches = True
                resolved: list[tuple[Path, Path]] = []
                media_root = (root / "media").resolve()
                for item in moves:
                    variant = variants[str(item["variant_id"])]
                    source = Path(str(item["source"])).expanduser().resolve()
                    target = Path(str(item["target"])).expanduser().resolve()
                    target.relative_to(media_root)
                    manifest_path = (
                        Path(variant["path"]).expanduser().resolve()
                        if Path(variant["path"]).expanduser().is_absolute()
                        else (root / variant["path"]).resolve()
                    )
                    source_matches &= manifest_path == source
                    target_matches &= manifest_path == target
                    resolved.append((source, target))

                if target_matches:
                    journal_path.unlink()
                    return current
                if not source_matches:
                    raise ValueError(
                        "Video family move journal conflicts with the manifest"
                    )

                for source, target in reversed(resolved):
                    if source == target or source.exists():
                        if source != target and target.exists():
                            raise FileExistsError(
                                f"Both video move paths exist: {source}, {target}"
                            )
                        continue
                    if not target.exists():
                        raise FileNotFoundError(
                            f"Both video move paths are missing: {source}, {target}"
                        )
                    source.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(target, source)
                journal_path.unlink()
                LOGGER.warning("Recovered interrupted video family move: %s", family_id)
                return current
            except Exception as exc:
                raise RuntimeError(
                    "检测到未完成的视频族移动，自动恢复失败；"
                    f"请保留 {journal_path.name} 并检查视频库。"
                ) from exc

    def save(
        self,
        *,
        recover_invalid_current: bool = False,
        preserve_existing_backup: bool = False,
    ) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        with project_write_lock(self.root, LOCK_NAME):
            self._save_locked(
                recover_invalid_current=recover_invalid_current,
                preserve_existing_backup=preserve_existing_backup,
            )

    def _save_locked(
        self,
        *,
        recover_invalid_current: bool = False,
        preserve_existing_backup: bool = False,
    ) -> None:
        current: dict[str, Any] | None = None
        if self.manifest_path.is_file():
            try:
                current = self._read_manifest(self.manifest_path)
            except Exception as exc:
                if not recover_invalid_current:
                    raise RuntimeError(
                        "当前视频库清单已损坏或不可读；拒绝覆盖。"
                        "请重新打开视频库，让程序先从备份恢复。"
                    ) from exc
        if current is not None and int(current.get("revision", 0)) != self._revision:
            raise RuntimeError(
                "Video project changed in another window; reopen it before saving"
            )

        payload = copy.deepcopy(self.data)
        payload["revision"] = self._revision + 1
        payload["updated_at"] = utc_now()
        fd, temp_name = tempfile.mkstemp(
            prefix=".video-project-", suffix=".json", dir=self.root
        )
        os.close(fd)
        temp_path = Path(temp_name)
        backup_temp = self.root / f".{BACKUP_MANIFEST_NAME}.tmp"
        try:
            temp_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            self._read_manifest(temp_path)
            if current is not None and not preserve_existing_backup:
                shutil.copyfile(self.manifest_path, backup_temp)
                os.replace(backup_temp, self.backup_manifest_path)
            os.replace(temp_path, self.manifest_path)
            self._revision = payload["revision"]
            self.data["revision"] = self._revision
            self.data["updated_at"] = payload["updated_at"]
        finally:
            temp_path.unlink(missing_ok=True)
            backup_temp.unlink(missing_ok=True)

    def reload(self) -> None:
        refreshed = type(self).open(self.root)
        self.data = refreshed.data
        self._revision = refreshed._revision
        self.recovered_from_backup = refreshed.recovered_from_backup
        self.recovery_detail = refreshed.recovery_detail

    def record(self, action: str, **details: Any) -> None:
        LOGGER.info("%s %s", action, details)

    def encode_path(self, path: Path) -> str:
        resolved = path.expanduser().resolve()
        try:
            return resolved.relative_to(self.root).as_posix()
        except ValueError:
            return str(resolved)

    def resolve_path(self, value: str) -> Path:
        path = Path(value).expanduser()
        return path.resolve() if path.is_absolute() else (self.root / path).resolve()

    def deck_source_path(self, deck: dict[str, Any]) -> Path:
        return self.resolve_path(deck["source_path"])

    def families(self) -> list[dict[str, Any]]:
        return self.data["families"]

    def decks(self) -> list[dict[str, Any]]:
        return self.data["decks"]

    def family(self, family_id: str) -> dict[str, Any]:
        for family in self.families():
            if family["id"] == family_id:
                return family
        raise KeyError(f"Unknown video family: {family_id}")

    def deck(self, deck_id: str) -> dict[str, Any]:
        for deck in self.decks():
            if deck["id"] == deck_id:
                return deck
        raise KeyError(f"Unknown PPTX: {deck_id}")

    def find_variant(self, variant_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        for family in self.families():
            for variant in family["variants"]:
                if variant["id"] == variant_id:
                    return family, variant
        raise KeyError(f"Unknown video version: {variant_id}")

    def find_variant_by_hash(
        self, sha256: str
    ) -> tuple[dict[str, Any], dict[str, Any]] | None:
        for family in self.families():
            for variant in family["variants"]:
                if variant["sha256"] == sha256:
                    return family, variant
        return None

    def family_by_known_hash(self, digest: str) -> dict[str, Any] | None:
        matches = [
            family
            for family in self.families()
            if digest in family.get("known_hashes", [])
            or any(
                variant.get("sha256") == digest
                for variant in family.get("variants", [])
            )
        ]
        if len(matches) > 1:
            names = "、".join(
                str(family.get("name") or family["id"]) for family in matches
            )
            raise RuntimeError(
                f"视频哈希 {digest[:12]} 同时属于多个视频族（{names}）；"
                "已停止自动匹配，请先运行“库体检”并人工归并。"
            )
        return matches[0] if matches else None

    def hash_catalog(self) -> dict[str, Any]:
        return {
            "format": "doc-media-video-hash-catalog",
            "version": 1,
            "exported_at": utc_now(),
            "project_id": self.data["project_id"],
            "families": [
                {
                    "name": family["name"],
                    "hashes": sorted(
                        {
                            *family.get("known_hashes", []),
                            *family.get("source_hashes", []),
                            *(
                                variant["sha256"]
                                for variant in family.get("variants", [])
                            ),
                        }
                    ),
                }
                for family in self.families()
            ],
        }

    def merge_hash_catalog(self, catalog: Any) -> dict[str, int]:
        if (
            not isinstance(catalog, dict)
            or catalog.get("format") != "doc-media-video-hash-catalog"
            or catalog.get("version") != 1
            or not isinstance(catalog.get("families"), list)
        ):
            raise ValueError("不是受支持的视频哈希目录。")
        local_by_hash: dict[str, set[str]] = {}
        for family in self.families():
            hashes = {
                *family.get("known_hashes", []),
                *family.get("source_hashes", []),
                *(variant["sha256"] for variant in family.get("variants", [])),
            }
            for digest in hashes:
                local_by_hash.setdefault(digest, set()).add(family["id"])

        added = matched = skipped = conflicts = 0
        total_hashes = 0
        for incoming in catalog["families"][:10000]:
            if not isinstance(incoming, dict) or not isinstance(
                incoming.get("hashes"), list
            ):
                skipped += 1
                continue
            hashes = {
                str(digest).lower()
                for digest in incoming["hashes"]
                if isinstance(digest, str)
                and len(digest) == 64
                and all(character in "0123456789abcdefABCDEF" for character in digest)
            }
            total_hashes += len(hashes)
            if total_hashes > 100000:
                raise ValueError("哈希目录过大，已停止导入。")
            anchors = {
                family_id
                for digest in hashes
                for family_id in local_by_hash.get(digest, set())
            }
            if not anchors:
                skipped += 1
                continue
            if len(anchors) != 1:
                conflicts += 1
                continue
            family = self.family(anchors.pop())
            known = family.setdefault("known_hashes", [])
            new_hashes = sorted(hashes.difference(known))
            known.extend(new_hashes)
            added += len(new_hashes)
            matched += 1
            for digest in new_hashes:
                local_by_hash.setdefault(digest, set()).add(family["id"])
        if added:
            self.record("hash_catalog_merged", added=added, matched=matched)
            self.save()
        return {
            "matched": matched,
            "added": added,
            "skipped": skipped,
            "conflicts": conflicts,
        }

    def family_by_content_fingerprint(
        self, fingerprint: dict[str, Any]
    ) -> dict[str, Any] | None:
        matches = [
            family
            for family in self.families()
            if family.get("content_fingerprint")
            and _fingerprints_match(family["content_fingerprint"], fingerprint)
        ]
        return matches[0] if len(matches) == 1 else None

    def _suggest_families(
        self,
        fingerprint: dict[str, Any],
        *,
        limit: int = 8,
    ) -> list[dict[str, Any]]:
        suggestions: list[dict[str, Any]] = []
        source_duration = max(int(fingerprint.get("duration_ms") or 0), 1)
        source_aspect = max(int(fingerprint.get("aspect_ppm") or 0), 1)
        for family in self.families():
            family_fingerprint = self._family_fingerprint(family)
            if family_fingerprint is None:
                continue
            confidence = _fingerprint_confidence(fingerprint, family_fingerprint)
            duration_diff = int(confidence.get("duration_diff_ms") or source_duration)
            family_aspect = max(int(family_fingerprint.get("aspect_ppm") or 0), 1)
            aspect_diff = abs(source_aspect - family_aspect) / max(
                source_aspect, family_aspect
            )
            frame_total = confidence.get("frame_total_distance")
            audio_mismatch = not bool(confidence.get("audio_consistent"))
            score = 100.0
            score -= min(45.0, duration_diff / source_duration * 900)
            score -= min(20.0, aspect_diff * 400)
            score -= min(25.0, float(frame_total or 64) / 3)
            score -= 25.0 if audio_mismatch else 0.0
            if confidence.get("matched"):
                score = max(score, 95.0)
            source = self.source_variant(family)
            try:
                source_path = self.require_variant_path(source)
            except (FileNotFoundError, ValueError):
                continue
            suggestions.append(
                {
                    "family_id": family["id"],
                    "family_name": family["name"],
                    "source_variant_id": source["id"],
                    "source_path": str(source_path),
                    "source_sha256": source["sha256"],
                    "width": int(source.get("width") or 0),
                    "height": int(source.get("height") or 0),
                    "duration_sec": float(source.get("duration_sec") or 0),
                    "score": round(max(0.0, score), 1),
                    "strict_match": bool(confidence.get("matched")),
                    "confidence": confidence,
                }
            )
        suggestions.sort(
            key=lambda item: (
                item["strict_match"],
                item["score"],
                int(item["width"]) * int(item["height"]),
            ),
            reverse=True,
        )
        return suggestions[: max(1, limit)]

    def suggest_video_matches(self, source: Path, *, limit: int = 8) -> dict[str, Any]:
        source = source.expanduser().resolve()
        metadata = probe_video(source)
        fingerprint = _video_fingerprint(source, metadata)
        if fingerprint is None:
            raise RuntimeError(f"Cannot create a reliable video fingerprint: {source}")
        return {
            "source": str(source),
            "sha256": sha256_file(source),
            "metadata": metadata,
            "fingerprint": fingerprint,
            "candidates": self._suggest_families(fingerprint, limit=limit),
        }

    def _family_by_damaged_prefix(
        self,
        name: str,
        path: Path,
        digest: str,
        metadata: dict[str, Any],
    ) -> dict[str, Any] | None:
        if not metadata.get("probe_error") or path.stat().st_size < 1024 * 1024:
            return None
        normalized_name = _normalized_video_name(name)
        if not normalized_name:
            return None
        size = path.stat().st_size
        matches = []
        for family in self.families():
            if _normalized_video_name(family["name"]) != normalized_name:
                continue
            source = self.source_variant(family)
            try:
                source_path = self.require_variant_path(source)
            except (FileNotFoundError, ValueError):
                continue
            if source_path.stat().st_size <= size:
                continue
            if _sha256_file_prefix(source_path, size) == digest:
                matches.append(family)
        return matches[0] if len(matches) == 1 else None

    def source_variant(self, family: dict[str, Any]) -> dict[str, Any]:
        source_id = family.get("source_variant_id")
        for variant in family["variants"]:
            if variant["id"] == source_id:
                return variant
        return next(
            (
                variant
                for variant in family["variants"]
                if variant.get("profile") == "original"
            ),
            family["variants"][0],
        )

    def archive_pptx_videos(
        self,
        pptx_path: Path,
        *,
        source_quality: str = "1080p",
        category: str = "",
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        pptx_path = pptx_path.expanduser().resolve()
        if not pptx_path.is_file() or pptx_path.suffix.lower() != ".pptx":
            raise ValueError(f"Expected an existing PPTX file: {pptx_path}")
        if source_quality not in {"1080p", "mp4", "original"}:
            raise ValueError(
                f"Unsupported video library source quality: {source_quality}"
            )
        category_path = normalize_library_category(category)
        scanned = scan_embedded_videos(pptx_path)
        staging = Path(tempfile.mkdtemp(prefix="video-library-import-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        created_files: list[Path] = []
        media_families: dict[str, str] = {}
        added = 0
        reused = 0
        candidates_added = 0
        changed = False
        try:
            with ZipFile(pptx_path, "r") as archive:
                total = len(scanned)
                for index, asset in enumerate(
                    sorted(scanned.values(), key=lambda item: item.media_path), start=1
                ):
                    _check_cancelled(cancel_callback)
                    if progress_callback:
                        progress_callback(
                            f"正在归档视频 {index}/{total}: {Path(asset.media_path).name}"
                        )
                    suffix = Path(asset.media_path).suffix.lower() or ".bin"
                    staged = staging / f"asset-{index}{suffix}"
                    digest = _copy_zip_member(archive, asset.media_path, staged)
                    occurrence = asset.occurrences[0] if asset.occurrences else None
                    display_name = _safe_name(
                        occurrence.shape_name
                        if occurrence
                        else Path(asset.media_path).stem
                    )
                    family = self.family_by_known_hash(digest)
                    metadata: dict[str, Any] | None = None
                    if family is None:
                        metadata = probe_video(staged)
                        fingerprint = _video_fingerprint(staged, metadata)
                        if fingerprint is not None:
                            family = self.family_by_content_fingerprint(fingerprint)
                        if family is None:
                            family = self._family_by_damaged_prefix(
                                display_name, staged, digest, metadata
                            )
                        if family is not None:
                            known = family.setdefault("known_hashes", [])
                            if digest not in known:
                                known.append(digest)
                            changed = True
                    if family is not None:
                        reused += 1
                        current = self.source_variant(family)
                        best_quality = max(
                            self._variant_quality_key(item)
                            for item in family["variants"]
                        )
                        candidate_width = int((metadata or {}).get("width") or 0)
                        candidate_height = int((metadata or {}).get("height") or 0)
                        desired_width, desired_height = (
                            candidate_width,
                            candidate_height,
                        )
                        if (
                            source_quality == "1080p"
                            and candidate_width > 0
                            and candidate_height > 0
                        ):
                            desired_width, desired_height = _archived_dimensions(
                                candidate_width, candidate_height
                            )
                        candidate_pixels = desired_width * desired_height
                        candidate_bitrate = int(
                            (metadata or {}).get("bitrate_kbps") or 0
                        )
                        potentially_better = candidate_pixels > best_quality[0] or (
                            candidate_pixels == best_quality[0]
                            and candidate_bitrate > best_quality[1] * 1.05
                        )
                        if metadata is not None and potentially_better:
                            stored = staged
                            profile = "original"
                            normalize_mp4 = source_quality in {
                                "1080p",
                                "mp4",
                            } and media_needs_mp4(asset.media_path)
                            if (
                                candidate_width > desired_width
                                or candidate_height > desired_height
                                or normalize_mp4
                            ):
                                stored = staging / f"{digest}_{source_quality}.mp4"
                                _transcode_high_quality_mp4(
                                    staged,
                                    stored,
                                    candidate_width,
                                    candidate_height,
                                    family_id=family["id"],
                                    limit_1080p=source_quality == "1080p",
                                )
                                metadata = probe_video(stored)
                                suffix = ".mp4"
                                profile = f"{source_quality}_source"
                            stored_digest = sha256_file(stored)
                            target = _unique_path(
                                self.variant_path(current).parent
                                / _variant_filename(
                                    family["name"], metadata, stored_digest, suffix
                                )
                            )
                            shutil.move(stored, target)
                            created_files.append(target)
                            variant = {
                                "id": str(uuid.uuid4()),
                                "label": "source",
                                "profile": profile,
                                "path": self.encode_path(target),
                                "sha256": stored_digest,
                                "size_bytes": target.stat().st_size,
                                "mtime_ns": target.stat().st_mtime_ns,
                                "created_at": utc_now(),
                                "source_variant_id": current["id"],
                                **metadata,
                            }
                            family["variants"].append(variant)
                            candidates_added += 1
                            known = family.setdefault("known_hashes", [])
                            for known_digest in (digest, stored_digest):
                                if known_digest not in known:
                                    known.append(known_digest)
                            changed = True
                        staged.unlink(missing_ok=True)
                    else:
                        family_id = str(uuid.uuid4())
                        variant_id = str(uuid.uuid4())
                        metadata = metadata or probe_video(staged)
                        stored = staged
                        profile = "original"
                        width = int(metadata.get("width") or 0)
                        height = int(metadata.get("height") or 0)
                        if source_quality == "1080p" and width > 0 and height > 0:
                            max_width, max_height = _video_envelope(width, height)
                            if (
                                width > max_width
                                or height > max_height
                                or media_needs_mp4(asset.media_path)
                            ):
                                stored = staging / f"{digest}_1080p.mp4"
                                _transcode_high_quality_mp4(
                                    staged,
                                    stored,
                                    width,
                                    height,
                                    family_id=family_id,
                                )
                                metadata = probe_video(stored)
                                suffix = ".mp4"
                                profile = "1080p_source"
                        elif (
                            source_quality == "mp4"
                            and width > 0
                            and height > 0
                            and media_needs_mp4(asset.media_path)
                        ):
                            stored = staging / f"{digest}_mp4.mp4"
                            _transcode_high_quality_mp4(
                                staged,
                                stored,
                                width,
                                height,
                                family_id=family_id,
                                limit_1080p=False,
                            )
                            metadata = probe_video(stored)
                            suffix = ".mp4"
                            profile = "mp4_source"
                        stored_digest = sha256_file(stored)
                        family_dir = self.root / "media" / category_path
                        family_dir.mkdir(parents=True, exist_ok=True)
                        target = _unique_path(
                            family_dir
                            / _variant_filename(
                                display_name, metadata, stored_digest, suffix
                            )
                        )
                        shutil.move(stored, target)
                        created_files.append(target)
                        variant = {
                            "id": variant_id,
                            "label": "source",
                            "profile": profile,
                            "path": self.encode_path(target),
                            "sha256": stored_digest,
                            "size_bytes": target.stat().st_size,
                            "mtime_ns": target.stat().st_mtime_ns,
                            "created_at": utc_now(),
                            "source_variant_id": None,
                            **metadata,
                        }
                        family = {
                            "id": family_id,
                            "name": display_name,
                            "category": (
                                ""
                                if category_path == Path()
                                else category_path.as_posix()
                            ),
                            "source_variant_id": variant_id,
                            "active_variant_id": variant_id,
                            "known_hashes": list(
                                dict.fromkeys([digest, stored_digest])
                            ),
                            "source_hashes": list(
                                dict.fromkeys([digest, stored_digest])
                            ),
                            "variants": [variant],
                        }
                        fingerprint = _video_fingerprint(target, metadata)
                        if fingerprint is not None:
                            family["content_fingerprint"] = fingerprint
                        self.families().append(family)
                        added += 1
                        changed = True
                    media_families[asset.media_path] = family["id"]
            if changed:
                self.save()
            return {
                "media_families": media_families,
                "added": added,
                "reused": reused,
                "candidates_added": candidates_added,
            }
        except Exception:
            self.data = original_data
            for path in created_files:
                path.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(staging, ignore_errors=True)

    def register_compressed_pptx_hashes(
        self,
        output_pptx: Path,
        report_path: Path,
        media_families: dict[str, str],
    ) -> int:
        output_pptx = output_pptx.expanduser().resolve()
        report = json.loads(report_path.read_text(encoding="utf-8"))
        added = 0
        with ZipFile(output_pptx, "r") as archive:
            for item in report.get("videos", []):
                source_path = item.get("media_path")
                output_path = str(
                    item.get("output_media_path") or source_path or ""
                ).lstrip("/")
                family_id = media_families.get(str(source_path))
                if family_id is None or output_path not in archive.namelist():
                    continue
                digest = _sha256_zip_member(archive, output_path)
                family = self.family(family_id)
                known = family.setdefault("known_hashes", [])
                if digest not in known:
                    known.append(digest)
                    added += 1
        if added:
            self.save()
        return added

    def archive_and_register_pptx(
        self,
        pptx_path: Path,
        *,
        source_quality: str = "1080p",
        category: str = "",
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        """Archive media and register the PPTX shape mapping as one workflow.

        The lower-level ``archive_pptx_videos`` method intentionally only stores
        media. User-facing GUI, CLI, compression, and Agent workflows should use
        this method so an extracted video never appears as unlinked merely
        because the deck registration step was forgotten.
        """
        original_data = copy.deepcopy(self.data)
        original_revision = self._revision
        original_variant_ids = {
            variant["id"]
            for family in self.families()
            for variant in family["variants"]
        }
        try:
            archive = self.archive_pptx_videos(
                pptx_path,
                source_quality=source_quality,
                category=category,
                progress_callback=progress_callback,
                cancel_callback=cancel_callback,
            )
            if progress_callback:
                progress_callback(f"正在登记 PPTX 形状关联：{pptx_path.name}")
            deck = self.add_deck(
                pptx_path,
                progress_callback=progress_callback,
                cancel_callback=cancel_callback,
            )
            return {**archive, "deck": deck}
        except Exception:
            created_paths = [
                self.variant_path(variant)
                for family in self.families()
                for variant in family["variants"]
                if variant["id"] not in original_variant_ids
            ]
            manifest_changed = self._revision != original_revision
            self.data = original_data
            if manifest_changed:
                try:
                    # The current manifest contains the failed intermediate
                    # archive. Restore the snapshot without replacing the last
                    # known-good backup with that intermediate state.
                    self.save(preserve_existing_backup=True)
                except Exception as rollback_error:
                    raise RuntimeError(
                        "归档 PPTX 失败，且无法回滚视频库清单；"
                        "请停止写入并使用 video-project.json.bak 恢复。"
                    ) from rollback_error
            for path in created_paths:
                path.unlink(missing_ok=True)
                try:
                    path.parent.rmdir()
                except OSError:
                    pass
            raise

    def _delivery_master(
        self, family: dict[str, Any], work: Path, tier: str = DEFAULT_BACKFILL_TIER
    ) -> tuple[Path, str]:
        spec = _tier_spec(tier)
        variant = self.source_variant(family)
        source = self.require_variant_path(variant)
        bitrate_cap = int(spec.get("bitrate_kbps") or 0)
        metadata = (
            variant
            if variant.get("width")
            and (bitrate_cap <= 0 or variant.get("bitrate_kbps"))
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

    def upgrade_pptx_from_library(
        self,
        input_pptx: Path,
        *,
        output_path: Path | None = None,
        incompatible_only: bool = False,
        family_overrides: dict[str, str] | None = None,
        remember_manual_matches: set[str] | None = None,
        keep_current_media: set[str] | None = None,
        quality_tier: str = DEFAULT_BACKFILL_TIER,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        input_pptx = input_pptx.expanduser().resolve()
        if not input_pptx.is_file() or input_pptx.suffix.lower() != ".pptx":
            raise ValueError(f"Expected an existing PPTX file: {input_pptx}")
        spec = _tier_spec(quality_tier)
        output = (
            output_path.expanduser().resolve()
            if output_path
            else input_pptx.with_name(f"{input_pptx.stem}_{spec['suffix']}.pptx")
        )
        output = _unique_path(output)
        scanned = scan_embedded_videos(input_pptx)
        work = Path(tempfile.mkdtemp(prefix="video-library-upgrade-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        replacements: dict[str, Path] = {}
        replacement_infos: dict[str, Any] = {}
        relationship_map: dict[str, str] = {}
        remove_paths: set[str] = set()
        compact_assets: dict[str, CompactVideoAsset] = {}
        expected_hashes: dict[str, str] = {}
        masters: dict[tuple[str, str], tuple[Path, str]] = {}
        output_parts_by_hash: dict[str, str] = {}
        learned_aliases: list[tuple[dict[str, Any], str]] = []
        matched = 0
        content_matched = 0
        manual_matched = 0
        already_high_quality = 0
        kept_current = 0
        unmatched: list[str] = []
        family_overrides = family_overrides or {}
        keep_current_media = keep_current_media or set()
        reserved = set(scanned)
        try:
            with ZipFile(input_pptx, "r") as archive:
                total = len(scanned)
                for index, asset in enumerate(
                    sorted(scanned.values(), key=lambda item: item.media_path), start=1
                ):
                    _check_cancelled(cancel_callback)
                    if incompatible_only and not media_needs_mp4(asset.media_path):
                        continue
                    if asset.media_path in keep_current_media:
                        kept_current += 1
                        continue
                    if progress_callback:
                        progress_callback(
                            f"正在匹配视频 {index}/{total}: {Path(asset.media_path).name}"
                        )
                    digest = _sha256_zip_member(archive, asset.media_path)
                    family = self.family_by_known_hash(digest)
                    manual_family_id = family_overrides.get(asset.media_path)
                    if manual_family_id:
                        family = self.family(manual_family_id)
                        manual_matched += 1
                        if (
                            remember_manual_matches is None
                            or asset.media_path in remember_manual_matches
                        ):
                            learned_aliases.append((family, digest))
                    if family is None:
                        suffix = Path(asset.media_path).suffix.lower() or ".bin"
                        candidate = work / f"candidate-{index}{suffix}"
                        _copy_zip_member(archive, asset.media_path, candidate)
                        fingerprint = _video_fingerprint(candidate)
                        if fingerprint is not None:
                            family = self.family_by_content_fingerprint(fingerprint)
                        if family is None:
                            unmatched.append(asset.media_path)
                            continue
                        content_matched += 1
                        learned_aliases.append((family, digest))
                    master_key = (family["id"], quality_tier)
                    master = masters.get(master_key)
                    if master is None:
                        master = self._delivery_master(family, work, quality_tier)
                        masters[master_key] = master
                    master_path, master_digest = master
                    if master_digest not in family.setdefault("known_hashes", []):
                        learned_aliases.append((family, master_digest))
                    existing_part = output_parts_by_hash.get(master_digest)
                    if existing_part is not None:
                        relationship_map[asset.media_path] = existing_part
                        remove_paths.add(asset.media_path)
                        compact_assets[asset.media_path] = CompactVideoAsset(
                            media_path=asset.media_path,
                            zip_size=archive.getinfo(asset.media_path).file_size,
                            output_media_path=existing_part,
                        )
                        expected_hashes[existing_part] = master_digest
                        matched += 1
                        continue
                    if digest == master_digest and not media_needs_mp4(
                        asset.media_path
                    ):
                        output_parts_by_hash[master_digest] = asset.media_path
                        expected_hashes[asset.media_path] = master_digest
                        already_high_quality += 1
                        continue
                    output_part = asset.media_path
                    if media_needs_mp4(asset.media_path):
                        # The delivery master is always an MP4 (either a
                        # compatible source or a transcode). A non-MP4 part
                        # (wmv/avi/...) must be swapped for an .mp4 part so the
                        # package never carries MP4 bytes under a wrong
                        # extension and content type.
                        output_part = choose_output_media_path(
                            asset.media_path, reserved
                        )
                    if output_part != asset.media_path:
                        relationship_map[asset.media_path] = output_part
                        remove_paths.add(asset.media_path)
                        compact_assets[asset.media_path] = CompactVideoAsset(
                            media_path=asset.media_path,
                            zip_size=archive.getinfo(asset.media_path).file_size,
                            output_media_path=output_part,
                        )
                    reserved.add(output_part)
                    output_parts_by_hash[master_digest] = output_part
                    replacements[output_part] = master_path
                    replacement_infos[output_part] = archive.getinfo(asset.media_path)
                    expected_hashes[output_part] = master_digest
                    matched += 1
            if not replacements and not relationship_map:
                return {
                    "output_pptx": None,
                    "matched": 0,
                    "content_matched": 0,
                    "manual_matched": manual_matched,
                    "already_high_quality": already_high_quality,
                    "kept_current": kept_current,
                    "aliases_added": 0,
                    "unmatched": unmatched,
                    "quality_tier": quality_tier,
                }
            patch_output_pptx(
                input_pptx,
                output,
                replacements,
                relationship_path_map=relationship_map,
                replacement_infos=replacement_infos,
                remove_paths=remove_paths,
                video_assets=compact_assets,
            )
            with ZipFile(output, "r") as archive:
                if archive.testzip() is not None:
                    raise BadZipFile("Upgraded PPTX failed ZIP validation")
                for media_path, digest in expected_hashes.items():
                    if _sha256_zip_member(archive, media_path) != digest:
                        raise RuntimeError(
                            f"Upgraded video hash mismatch: {media_path}"
                        )
            aliases_added = 0
            for family, digest in learned_aliases:
                known = family.setdefault("known_hashes", [])
                if digest not in known:
                    known.append(digest)
                    aliases_added += 1
            if aliases_added:
                self.save()
            return {
                "output_pptx": output,
                "matched": matched,
                "content_matched": content_matched,
                "manual_matched": manual_matched,
                "already_high_quality": already_high_quality,
                "kept_current": kept_current,
                "aliases_added": aliases_added,
                "unmatched": unmatched,
                "quality_tier": quality_tier,
            }
        except Exception:
            self.data = original_data
            output.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def adopt_upgraded_deck_source(
        self,
        deck_id: str,
        source_pptx: Path,
        *,
        prefer_source_variant: bool = False,
    ) -> dict[str, Any]:
        """Rebind a registered deck after a validated in-place media upgrade.

        Matching uses slide/shape anchors. Every embedded media digest must
        resolve to a variant in the asset's existing family, so renames and
        media part extension changes cannot silently break an association.
        """
        deck = self.deck(deck_id)
        source = source_pptx.expanduser().resolve()
        if not source.is_file() or source.suffix.lower() != ".pptx":
            raise ValueError(f"Expected an existing PPTX file: {source}")

        scanned = scan_embedded_videos(source)
        by_anchor: dict[tuple[str, int], VideoAsset] = {}
        for asset in scanned.values():
            for occurrence in asset.occurrences:
                anchor = (occurrence.slide_path, occurrence.shape_id)
                if anchor in by_anchor:
                    raise ValueError(f"Duplicate video shape anchor: {anchor}")
                by_anchor[anchor] = asset

        original_data = copy.deepcopy(self.data)
        try:
            previous_source = self.deck_source_path(deck)
            rebound_assets: list[dict[str, Any]] = []
            with ZipFile(source, "r") as archive:
                if archive.testzip() is not None:
                    raise BadZipFile("Upgraded PPTX failed ZIP validation")
                for item in deck["assets"]:
                    matches = {
                        by_anchor[
                            (occurrence["slide_path"], occurrence["shape_id"])
                        ].media_path
                        for occurrence in item["occurrences"]
                        if (occurrence["slide_path"], occurrence["shape_id"])
                        in by_anchor
                    }
                    if len(matches) != 1:
                        raise ValueError(
                            "Upgraded PPTX no longer has a unique media part for "
                            f"{item['part_path']}"
                        )
                    media_path = matches.pop()
                    scanned_asset = scanned[media_path]
                    expected_anchors = {
                        (occurrence["slide_path"], occurrence["shape_id"])
                        for occurrence in item["occurrences"]
                    }
                    actual_anchors = {
                        (occurrence.slide_path, occurrence.shape_id)
                        for occurrence in scanned_asset.occurrences
                    }
                    if expected_anchors != actual_anchors:
                        raise ValueError(
                            "Upgraded PPTX changed the video occurrence mapping for "
                            f"{item['part_path']}"
                        )
                    digest = _sha256_zip_member(archive, media_path)
                    found = self.find_variant_by_hash(digest)
                    family = self.family(item["family_id"])
                    if found is not None and found[0]["id"] == item["family_id"]:
                        known = family.setdefault("known_hashes", [])
                        if digest not in known:
                            known.append(digest)
                        original_variant_id = (
                            family["source_variant_id"]
                            if prefer_source_variant
                            else found[1]["id"]
                        )
                    elif media_path == item["part_path"] and digest in family.get(
                        "known_hashes", []
                    ):
                        # Content-fingerprint imports can associate a deck's
                        # unchanged encoding without storing another physical
                        # variant. Preserve the configured restore policy.
                        original_variant_id = (
                            family["source_variant_id"]
                            if prefer_source_variant
                            else item["original_variant_id"]
                        )
                    else:
                        raise ValueError(
                            "Upgraded PPTX media does not belong to its registered "
                            f"video family: {media_path}"
                        )
                    rebound_assets.append(
                        {
                            **item,
                            "part_path": media_path,
                            "original_variant_id": original_variant_id,
                            "occurrences": [
                                asdict(occurrence)
                                for occurrence in scanned_asset.occurrences
                            ],
                        }
                    )

            anchors = [
                (item["part_path"], occurrence["slide_path"], occurrence["shape_id"])
                for item in rebound_assets
                for occurrence in item["occurrences"]
            ]
            stat = source.stat()
            encoded_source = self.encode_path(source)
            aliases = [
                value
                for value in deck.get("source_aliases", [])
                if value != encoded_source
            ]
            if previous_source != source:
                previous_encoded = self.encode_path(previous_source)
                if previous_encoded not in aliases:
                    aliases.append(previous_encoded)
            deck.update(
                {
                    "name": source.name,
                    "source_path": encoded_source,
                    "source_aliases": aliases,
                    "source_sha256": sha256_file(source),
                    "source_size_bytes": stat.st_size,
                    "source_mtime_ns": stat.st_mtime_ns,
                    "structure_sha256": hashlib.sha256(
                        json.dumps(sorted(anchors), ensure_ascii=False).encode("utf-8")
                    ).hexdigest(),
                    "assets": rebound_assets,
                }
            )
            self.save()
            self.record("deck_upgrade_adopted", deck_id=deck_id, source=str(source))
            return deck
        except Exception:
            self.data = original_data
            raise

    def variant_path(self, variant: dict[str, Any]) -> Path:
        return self.resolve_path(variant["path"])

    def require_variant_path(self, variant: dict[str, Any]) -> Path:
        path = self.variant_path(variant)
        if not path.is_file():
            raise FileNotFoundError(path)
        if (
            path.stat().st_size != variant["size_bytes"]
            or sha256_file(path) != variant["sha256"]
        ):
            raise ValueError(f"Video file changed since it was archived: {path}")
        return path

    def add_deck(
        self,
        pptx_path: Path,
        *,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        pptx_path = pptx_path.expanduser().resolve()
        if not pptx_path.is_file() or pptx_path.suffix.lower() != ".pptx":
            raise ValueError(f"Expected an existing PPTX file: {pptx_path}")
        source_sha = sha256_file(pptx_path)
        for deck in self.decks():
            known_deck_hashes = {
                deck["source_sha256"],
                *(
                    str(record.get("sha256") or "")
                    for record in self.deck_outputs(deck)
                ),
            }
            if source_sha in known_deck_hashes:
                encoded_path = self.encode_path(pptx_path)
                if encoded_path != deck["source_path"]:
                    aliases = deck.setdefault("source_aliases", [])
                    if encoded_path not in aliases:
                        aliases.append(encoded_path)
                        self.save()
                return deck

        scanned = scan_embedded_videos(pptx_path)
        deck_id = str(uuid.uuid4())
        staging = Path(tempfile.mkdtemp(prefix="video-import-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        created_files: list[Path] = []
        deck_assets: list[dict[str, Any]] = []
        try:
            with ZipFile(pptx_path, "r") as archive:
                total = len(scanned)
                for index, asset in enumerate(
                    sorted(scanned.values(), key=lambda item: item.media_path), start=1
                ):
                    _check_cancelled(cancel_callback)
                    if progress_callback:
                        progress_callback(
                            f"正在提取视频 {index}/{total}: {Path(asset.media_path).name}"
                        )
                    suffix = Path(asset.media_path).suffix.lower() or ".bin"
                    staged = staging / f"asset-{index}{suffix}"
                    digest = _copy_zip_member(archive, asset.media_path, staged)
                    family = self.family_by_known_hash(digest)
                    metadata: dict[str, Any] | None = None
                    fingerprint: dict[str, Any] | None = None
                    if family is None:
                        metadata = probe_video(staged)
                        fingerprint = _video_fingerprint(staged, metadata)
                        if fingerprint is not None:
                            family = self.family_by_content_fingerprint(fingerprint)
                    if family is not None:
                        variant = self.source_variant(family)
                        known = family.setdefault("known_hashes", [])
                        if digest not in known:
                            known.append(digest)
                        staged.unlink(missing_ok=True)
                    else:
                        occurrence = asset.occurrences[0] if asset.occurrences else None
                        display_name = _safe_name(
                            occurrence.shape_name
                            if occurrence
                            else Path(asset.media_path).stem
                        )
                        family_id = str(uuid.uuid4())
                        variant_id = str(uuid.uuid4())
                        family_dir = (
                            self.root / "media" / f"{display_name}_{family_id[:8]}"
                        )
                        family_dir.mkdir(parents=True, exist_ok=True)
                        metadata = metadata or probe_video(staged)
                        target = family_dir / _variant_filename(
                            display_name, metadata, digest, suffix
                        )
                        shutil.move(staged, target)
                        created_files.append(target)
                        variant = {
                            "id": variant_id,
                            "label": "original",
                            "profile": "original",
                            "path": self.encode_path(target),
                            "sha256": digest,
                            "size_bytes": target.stat().st_size,
                            "mtime_ns": target.stat().st_mtime_ns,
                            "created_at": utc_now(),
                            "source_variant_id": None,
                            **metadata,
                        }
                        family = {
                            "id": family_id,
                            "name": display_name,
                            "source_variant_id": variant_id,
                            "active_variant_id": variant_id,
                            "known_hashes": [digest],
                            "source_hashes": [digest],
                            "variants": [variant],
                        }
                        if fingerprint is not None:
                            family["content_fingerprint"] = fingerprint
                        self.families().append(family)

                    deck_assets.append(
                        {
                            "part_path": asset.media_path,
                            "placeholder_part": (
                                f"ppt/media/pptx_tools_{deck_id[:8]}_{index:03d}.mp4"
                            ),
                            "family_id": family["id"],
                            "original_variant_id": variant["id"],
                            "occurrences": [asdict(item) for item in asset.occurrences],
                        }
                    )

            anchors = [
                (item["part_path"], occurrence["slide_path"], occurrence["shape_id"])
                for item in deck_assets
                for occurrence in item["occurrences"]
            ]
            structure_sha = hashlib.sha256(
                json.dumps(sorted(anchors), ensure_ascii=False).encode("utf-8")
            ).hexdigest()
            deck = {
                "id": deck_id,
                "name": pptx_path.name,
                "source_path": self.encode_path(pptx_path),
                "source_sha256": source_sha,
                "source_aliases": [],
                "source_size_bytes": pptx_path.stat().st_size,
                "source_mtime_ns": pptx_path.stat().st_mtime_ns,
                "structure_sha256": structure_sha,
                "created_at": utc_now(),
                "assets": deck_assets,
                "detached_outputs": [],
                "restored_outputs": [],
                "optimized_outputs": [],
            }
            self.decks().append(deck)
            self.save()
            self.record(
                "deck_added",
                deck_id=deck_id,
                name=pptx_path.name,
                videos=len(deck_assets),
            )
            return deck
        except Exception:
            self.data = original_data
            for path in created_files:
                path.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(staging, ignore_errors=True)

    def rename_family(self, family_id: str, name: str) -> None:
        family = self.family(family_id)
        previous = family["name"]
        family["name"] = _safe_name(name)
        try:
            self.save()
        except Exception:
            family["name"] = previous
            raise
        self.record("family_renamed", family_id=family_id, name=family["name"])

    def rename_family_and_source(self, family_id: str, name: str) -> Path:
        family = self.family(family_id)
        variant = self.source_variant(family)
        source = self.require_variant_path(variant)
        clean = _safe_name(_video_stem(name))
        target = source.with_name(
            _variant_filename(clean, variant, variant["sha256"], source.suffix)
        )
        if target != source and target.exists():
            raise FileExistsError(target)
        previous = copy.deepcopy(family)
        if target != source:
            source.rename(target)
        family["name"] = clean
        variant["path"] = self.encode_path(target)
        variant["mtime_ns"] = target.stat().st_mtime_ns
        try:
            self.save()
        except Exception:
            family.clear()
            family.update(previous)
            if target != source:
                try:
                    target.rename(source)
                except OSError as rollback_error:
                    LOGGER.error(
                        "Unable to roll back video family rename: %s", rollback_error
                    )
            raise
        self.record("family_renamed", family_id=family_id, name=clean)
        return target

    def rename_variant_file(self, variant_id: str, filename: str) -> Path:
        _, variant = self.find_variant(variant_id)
        source = self.require_variant_path(variant)
        suffix = source.suffix
        clean = _video_stem(filename)
        target = source.with_name(
            _variant_filename(clean, variant, variant["sha256"], suffix)
        )
        if target != source and target.exists():
            raise FileExistsError(target)
        previous = copy.deepcopy(variant)
        source.rename(target)
        variant["path"] = self.encode_path(target)
        variant["mtime_ns"] = target.stat().st_mtime_ns
        try:
            self.save()
        except Exception:
            variant.clear()
            variant.update(previous)
            try:
                target.rename(source)
            except OSError as rollback_error:
                LOGGER.error("Unable to roll back video rename: %s", rollback_error)
            raise
        self.record("variant_renamed", variant_id=variant_id, path=variant["path"])
        return target

    def move_variant(self, variant_id: str, destination: Path) -> Path:
        _, variant = self.find_variant(variant_id)
        source = self.require_variant_path(variant)
        destination = destination.expanduser().resolve()
        destination.mkdir(parents=True, exist_ok=True)
        target = destination / source.name
        if target != source and target.exists():
            raise FileExistsError(target)
        previous = copy.deepcopy(variant)
        shutil.move(source, target)
        variant["path"] = self.encode_path(target)
        variant["mtime_ns"] = target.stat().st_mtime_ns
        try:
            self.save()
        except Exception:
            variant.clear()
            variant.update(previous)
            try:
                shutil.move(target, source)
            except OSError as rollback_error:
                LOGGER.error("Unable to roll back video move: %s", rollback_error)
            raise
        self.record("variant_moved", variant_id=variant_id, path=variant["path"])
        return target

    def move_family(self, family_id: str, category: str) -> list[Path]:
        """Move every version in a family and persist its library category."""
        family = self.family(family_id)
        category_path = normalize_library_category(category)
        destination = self.root / "media" / category_path
        targets: list[tuple[dict[str, Any], Path, Path]] = []
        seen_targets: set[Path] = set()
        for variant in family["variants"]:
            source = self.require_variant_path(variant)
            target = destination / source.name
            if target in seen_targets or (target != source and target.exists()):
                raise FileExistsError(target)
            seen_targets.add(target)
            targets.append((variant, source, target))

        previous = copy.deepcopy(family)
        moved: list[tuple[Path, Path]] = []
        with project_write_lock(self.root, LOCK_NAME):
            if self.family_move_journal_path.exists():
                raise RuntimeError(
                    "检测到上次视频族移动尚未恢复；请重新打开视频库后再操作。"
                )
            journal_temp = self.root / f"{FAMILY_MOVE_JOURNAL_NAME}.tmp"
            journal = {
                "family_id": family_id,
                "moves": [
                    {
                        "variant_id": variant["id"],
                        "source": str(source),
                        "target": str(target),
                    }
                    for variant, source, target in targets
                ],
            }
            try:
                journal_temp.write_text(
                    json.dumps(journal, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                os.replace(journal_temp, self.family_move_journal_path)
                destination.mkdir(parents=True, exist_ok=True)
                for variant, source, target in targets:
                    if target != source:
                        shutil.move(source, target)
                        moved.append((target, source))
                    variant["path"] = self.encode_path(target)
                    variant["mtime_ns"] = target.stat().st_mtime_ns
                family["category"] = (
                    "" if category_path == Path() else category_path.as_posix()
                )
                self._save_locked()
                try:
                    self.family_move_journal_path.unlink()
                except OSError as cleanup_error:
                    LOGGER.warning(
                        "Unable to remove completed video move journal: %s",
                        cleanup_error,
                    )
            except Exception as exc:
                family.clear()
                family.update(previous)
                rollback_errors: list[OSError] = []
                for target, source in reversed(moved):
                    try:
                        shutil.move(target, source)
                    except OSError as rollback_error:
                        rollback_errors.append(rollback_error)
                if not rollback_errors:
                    self.family_move_journal_path.unlink(missing_ok=True)
                if rollback_errors:
                    raise RuntimeError(
                        "视频族移动失败且回滚不完整；下次打开视频库时将再次恢复。"
                    ) from exc
                raise
            finally:
                journal_temp.unlink(missing_ok=True)
        self.record(
            "family_moved",
            family_id=family_id,
            category=family["category"],
            variants=len(targets),
        )
        return [target for _, _, target in targets]

    def _output_record(
        self, path: Path, kind: str, digest: str | None = None
    ) -> dict[str, Any]:
        path = path.expanduser().resolve()
        return {
            "id": str(uuid.uuid4()),
            "kind": kind,
            "path": self.encode_path(path),
            "sha256": digest or sha256_file(path),
            "size_bytes": path.stat().st_size,
            "mtime_ns": path.stat().st_mtime_ns,
            "created_at": utc_now(),
        }

    def deck_outputs(self, deck: dict[str, Any]) -> list[dict[str, Any]]:
        return [
            *deck.get("optimized_outputs", []),
            *deck.get("detached_outputs", []),
            *deck.get("restored_outputs", []),
        ]

    def find_output(self, output_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        for deck in self.decks():
            for record in self.deck_outputs(deck):
                if record.get("id") == output_id:
                    return deck, record
        raise KeyError(f"Unknown PPTX output: {output_id}")

    def register_optimized_output(
        self, source_path: Path, output_path: Path, digest: str | None = None
    ) -> dict[str, Any] | None:
        source = source_path.expanduser().resolve()
        output = output_path.expanduser().resolve()
        if not output.is_file() or output.suffix.lower() != ".pptx":
            raise ValueError(f"Expected an existing PPTX output: {output}")
        encoded_source = self.encode_path(source)
        deck = next(
            (
                item
                for item in self.decks()
                if item["source_path"] == encoded_source
                or encoded_source in item.get("source_aliases", [])
            ),
            None,
        )
        if deck is None:
            return None
        output_digest = digest or sha256_file(output)
        for record in self.deck_outputs(deck):
            if record.get("sha256") == output_digest:
                return record
            if self.resolve_path(record["path"]) == output:
                previous = copy.deepcopy(record)
                record.update(
                    {
                        "sha256": output_digest,
                        "size_bytes": output.stat().st_size,
                        "mtime_ns": output.stat().st_mtime_ns,
                        "created_at": utc_now(),
                        "kind": "optimized",
                    }
                )
                try:
                    self.save()
                except Exception:
                    record.clear()
                    record.update(previous)
                    raise
                self.record(
                    "optimized_output_updated",
                    deck_id=deck["id"],
                    path=record["path"],
                )
                return record
        optimized_outputs = deck.setdefault("optimized_outputs", [])
        if optimized_outputs:
            record = optimized_outputs[0]
            previous = copy.deepcopy(record)
            record.update(self._output_record(output, "optimized", output_digest))
            record["id"] = previous["id"]
            try:
                self.save()
            except Exception:
                record.clear()
                record.update(previous)
                raise
            self.record(
                "optimized_output_updated",
                deck_id=deck["id"],
                path=record["path"],
            )
            return record
        record = self._output_record(output, "optimized", output_digest)
        optimized_outputs.append(record)
        try:
            self.save()
        except Exception:
            deck["optimized_outputs"].remove(record)
            raise
        self.record(
            "optimized_output_registered", deck_id=deck["id"], path=record["path"]
        )
        return record

    def move_output(self, output_id: str, target_path: Path) -> Path:
        _, record = self.find_output(output_id)
        source = self.resolve_path(record["path"])
        if not source.is_file() or sha256_file(source) != record["sha256"]:
            raise ValueError(f"PPTX output is missing or changed: {source}")
        target = target_path.expanduser().resolve()
        if target.suffix.lower() != ".pptx":
            target = target.with_suffix(".pptx")
        target.parent.mkdir(parents=True, exist_ok=True)
        if target != source and target.exists():
            raise FileExistsError(target)
        previous = copy.deepcopy(record)
        shutil.move(source, target)
        record["path"] = self.encode_path(target)
        record["mtime_ns"] = target.stat().st_mtime_ns
        try:
            self.save()
        except Exception:
            record.clear()
            record.update(previous)
            try:
                shutil.move(target, source)
            except OSError as rollback_error:
                LOGGER.error("Unable to roll back PPTX output move: %s", rollback_error)
            raise
        self.record("pptx_output_moved", output_id=output_id, path=record["path"])
        return target

    def source_status(self, deck: dict[str, Any]) -> str:
        path = self.deck_source_path(deck)
        if not path.is_file():
            return "missing"
        if deck.get("source_size_bytes") not in (None, path.stat().st_size):
            return "modified"
        if deck.get("source_mtime_ns") not in (None, path.stat().st_mtime_ns):
            return "modified"
        return "available"

    def output_status(self, record: dict[str, Any]) -> str:
        path = self.resolve_path(record["path"])
        if not path.is_file():
            return "missing"
        if record.get("size_bytes") not in (None, path.stat().st_size):
            return "modified"
        if record.get("mtime_ns") not in (None, path.stat().st_mtime_ns):
            return "modified"
        return "available"

    def relink_missing_pptx(
        self,
        search_roots: Iterable[Path] = (),
        *,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> list[dict[str, str]]:
        missing: dict[str, list[tuple[str, dict[str, Any]]]] = {}
        expected_sizes: set[int] = set()
        has_unknown_size = False
        metadata_updated = False
        for deck in self.decks():
            source = self.deck_source_path(deck)
            source_size = deck.get("source_size_bytes")
            source_valid = source.is_file() and (
                source_size is None or source.stat().st_size == source_size
            )
            if source_valid and sha256_file(source) == deck["source_sha256"]:
                if deck.get("source_mtime_ns") != source.stat().st_mtime_ns:
                    metadata_updated = True
                deck["source_size_bytes"] = source.stat().st_size
                deck["source_mtime_ns"] = source.stat().st_mtime_ns
            else:
                missing.setdefault(deck["source_sha256"], []).append(("source", deck))
                if source_size is not None:
                    expected_sizes.add(int(source_size))
                else:
                    has_unknown_size = True
            for record in self.deck_outputs(deck):
                path = self.resolve_path(record["path"])
                size = record.get("size_bytes")
                valid = path.is_file() and (size is None or path.stat().st_size == size)
                if valid and sha256_file(path) == record["sha256"]:
                    if record.get("mtime_ns") != path.stat().st_mtime_ns:
                        metadata_updated = True
                    record["size_bytes"] = path.stat().st_size
                    record["mtime_ns"] = path.stat().st_mtime_ns
                    continue
                missing.setdefault(record["sha256"], []).append(("output", record))
                if size is not None:
                    expected_sizes.add(int(size))
                else:
                    has_unknown_size = True
        if not missing:
            if metadata_updated:
                self.save()
                self.record("pptx_metadata_refreshed")
            return []

        roots = [
            self.root,
            *(Path(item).expanduser().resolve() for item in search_roots),
        ]
        results: list[dict[str, str]] = []
        seen: set[Path] = set()
        for root in roots:
            if not root.is_dir():
                continue
            for candidate in root.rglob("*"):
                _check_cancelled(cancel_callback)
                try:
                    resolved = candidate.resolve()
                    if (
                        resolved in seen
                        or not candidate.is_file()
                        or candidate.suffix.lower() != ".pptx"
                    ):
                        continue
                    if (
                        not has_unknown_size
                        and expected_sizes
                        and candidate.stat().st_size not in expected_sizes
                    ):
                        continue
                except OSError:
                    continue
                seen.add(resolved)
                if progress_callback:
                    progress_callback(f"正在校验 {candidate.name}")
                digest = sha256_file(candidate)
                references = missing.pop(digest, [])
                for kind, item in references:
                    if kind == "source":
                        item["source_path"] = self.encode_path(resolved)
                        item["source_size_bytes"] = resolved.stat().st_size
                        item["source_mtime_ns"] = resolved.stat().st_mtime_ns
                    else:
                        item["path"] = self.encode_path(resolved)
                        item["size_bytes"] = resolved.stat().st_size
                        item["mtime_ns"] = resolved.stat().st_mtime_ns
                    results.append({"kind": kind, "path": str(resolved)})
                if not missing:
                    break
            if not missing:
                break
        if results or metadata_updated:
            self.save()
            self.record("pptx_files_relinked", count=len(results))
        return results

    def import_variant(
        self,
        family_id: str,
        source: Path,
        label: str = "custom",
        destination: Path | None = None,
        *,
        verify_identity: bool = True,
    ) -> dict[str, Any]:
        family = self.family(family_id)
        source = source.expanduser().resolve()
        if not source.is_file():
            raise FileNotFoundError(source)
        digest = sha256_file(source)
        existing_family = self.family_by_known_hash(digest)
        if existing_family is not None and existing_family["id"] != family_id:
            raise ValueError(
                f"This exact video already exists in '{existing_family['name']}'. "
                "Merge the video families instead of storing another copy."
            )
        for variant in family["variants"]:
            if variant["sha256"] == digest:
                origin = str(source)
                origins = variant.setdefault("origin_paths", [])
                if origin not in origins:
                    origins.append(origin)
                    self.save()
                return variant
        metadata = probe_video(source)
        fingerprint = _video_fingerprint(source, metadata)
        if verify_identity:
            family_fingerprint = self._family_fingerprint(family)
            if (
                fingerprint is None
                or family_fingerprint is None
                or not _fingerprints_match(family_fingerprint, fingerprint)
            ):
                raise ValueError(
                    "The selected video does not safely match this video family."
                )
        return self._store_external_variant(
            family,
            source,
            label=label,
            source_quality="original",
            destination=destination,
            metadata=metadata,
            fingerprint=fingerprint,
            identity_verified=verify_identity,
        )[0]

    def _family_fingerprint(self, family: dict[str, Any]) -> dict[str, Any] | None:
        fingerprint = family.get("content_fingerprint")
        if fingerprint:
            return fingerprint
        source = self.source_variant(family)
        try:
            return _video_fingerprint(self.require_variant_path(source), source)
        except (FileNotFoundError, ValueError):
            return None

    def _store_external_variant(
        self,
        family: dict[str, Any],
        source: Path,
        *,
        label: str,
        source_quality: str,
        destination: Path | None,
        metadata: dict[str, Any],
        fingerprint: dict[str, Any] | None,
        identity_verified: bool,
    ) -> tuple[dict[str, Any], bool]:
        if source_quality not in {"1080p", "mp4", "original"}:
            raise ValueError(
                f"Unsupported video library source quality: {source_quality}"
            )
        digest = sha256_file(source)
        new_family = not family["variants"]
        destination = (
            destination.expanduser().resolve()
            if destination
            else self.variant_path(family["variants"][0]).parent
        )
        destination.mkdir(parents=True, exist_ok=True)
        staging = Path(
            tempfile.mkdtemp(prefix="video-library-external-", dir=self.root)
        )
        original_data = copy.deepcopy(self.data)
        if not any(item is family for item in self.families()):
            self.families().append(family)
        target: Path | None = None
        try:
            stored = source
            stored_metadata = metadata
            profile = "original" if new_family else "external"
            width = int(metadata.get("width") or 0)
            height = int(metadata.get("height") or 0)
            if source_quality == "1080p" and width > 0 and height > 0:
                output_width, output_height = _archived_dimensions(width, height)
                if (
                    output_width < width
                    or output_height < height
                    or media_needs_mp4(source.name)
                ):
                    stored = staging / "external_1080p.mp4"
                    _transcode_high_quality_mp4(
                        source, stored, width, height, family_id=family["id"]
                    )
                    stored_metadata = probe_video(stored)
                    profile = "1080p_source"
            elif (
                source_quality == "mp4"
                and width > 0
                and height > 0
                and media_needs_mp4(source.name)
            ):
                stored = staging / "external_compatible.mp4"
                _transcode_high_quality_mp4(
                    source,
                    stored,
                    width,
                    height,
                    family_id=family["id"],
                    limit_1080p=False,
                )
                stored_metadata = probe_video(stored)
                profile = "mp4_source"
            stored_digest = sha256_file(stored)
            suffix = stored.suffix.lower() or source.suffix.lower()
            filename_base = (
                family["name"]
                if label in {"source", "original"}
                else f"{family['name']}_{_safe_name(label)}"
            )
            target = _unique_path(
                destination
                / _variant_filename(
                    filename_base,
                    stored_metadata,
                    stored_digest,
                    suffix,
                )
            )
            if stored == source:
                shutil.copy2(stored, target)
            else:
                shutil.move(stored, target)
            variant = {
                "id": str(uuid.uuid4()),
                "label": _safe_name(label, "custom"),
                "profile": profile,
                "path": self.encode_path(target),
                "sha256": stored_digest,
                "size_bytes": target.stat().st_size,
                "mtime_ns": target.stat().st_mtime_ns,
                "created_at": utc_now(),
                "source_variant_id": None,
                "origin_paths": [str(source)],
                **stored_metadata,
            }
            family["variants"].append(variant)
            if new_family or identity_verified:
                known = family.setdefault("known_hashes", [])
                for known_digest in (digest, stored_digest):
                    if known_digest not in known:
                        known.append(known_digest)
                if fingerprint is not None:
                    family.setdefault("content_fingerprint", fingerprint)
            promoted = new_family
            if promoted:
                family["source_variant_id"] = variant["id"]
                family["source_hashes"] = list(dict.fromkeys([digest, stored_digest]))
            if new_family:
                family["active_variant_id"] = variant["id"]
            self.save()
        except Exception:
            self.data = original_data
            if target is not None:
                target.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(staging, ignore_errors=True)
        self.record(
            "variant_imported", family_id=family["id"], variant_id=variant["id"]
        )
        return variant, promoted

    def import_external_video(
        self,
        source: Path,
        *,
        source_quality: str = "1080p",
        category: str = "",
        family_id: str | None = None,
        manual_confirmed: bool = False,
        force_new_family: bool = False,
        defer_suggestions: bool = False,
    ) -> dict[str, Any]:
        source = source.expanduser().resolve()
        if not source.is_file() or source.suffix.lower() not in VIDEO_SUFFIXES:
            raise ValueError(f"Expected a supported video file: {source}")
        category_path = normalize_library_category(category)
        digest = sha256_file(source)
        existing = self.family_by_known_hash(digest)
        if existing is not None:
            if family_id is not None and existing["id"] != family_id:
                raise ValueError(
                    f"This exact video already belongs to '{existing['name']}'."
                )
            matched_variant = self.find_variant_by_hash(digest)
            if matched_variant is not None:
                _, variant = matched_variant
                origin = str(source)
                origins = variant.setdefault("origin_paths", [])
                if origin not in origins:
                    origins.append(origin)
                    self.save()
            return {
                "status": "existing",
                "source": str(source),
                "family_id": existing["id"],
                "family_name": existing["name"],
                "promoted": False,
            }

        metadata = probe_video(source)
        fingerprint = _video_fingerprint(source, metadata)
        if fingerprint is None:
            raise RuntimeError(f"Cannot create a reliable video fingerprint: {source}")
        if family_id is not None:
            family = self.family(family_id)
            family_fingerprint = self._family_fingerprint(family)
            if not manual_confirmed and (
                family_fingerprint is None
                or not _fingerprints_match(family_fingerprint, fingerprint)
            ):
                raise ValueError(
                    "The selected video does not safely match this video family."
                )
            matches = [family]
        elif force_new_family:
            matches = []
        else:
            matches = [
                family
                for family in self.families()
                if (family_fingerprint := self._family_fingerprint(family))
                and _fingerprints_match(family_fingerprint, fingerprint)
            ]
            if len(matches) > 1:
                return {
                    "status": "ambiguous",
                    "source": str(source),
                    "metadata": metadata,
                    "sha256": digest,
                    "candidates": [
                        item
                        for item in self._suggest_families(
                            fingerprint, limit=len(self.families())
                        )
                        if item["family_id"] in {family["id"] for family in matches}
                    ],
                    "promoted": False,
                }
            if not matches and defer_suggestions:
                candidates = self._suggest_families(fingerprint)
                if candidates:
                    return {
                        "status": "suggested",
                        "source": str(source),
                        "metadata": metadata,
                        "sha256": digest,
                        "candidates": candidates,
                        "promoted": False,
                    }

        created = not matches
        if created:
            family_id = str(uuid.uuid4())
            name = _safe_name(source.stem)
            family = {
                "id": family_id,
                "name": name,
                "active_variant_id": "",
                "source_variant_id": "",
                "variants": [],
                "known_hashes": [],
                "source_hashes": [],
                "content_fingerprint": fingerprint,
            }
            destination = (
                self.root / "media" / category_path / f"{name}_{family_id[:8]}"
            )
            label = "source"
        else:
            family = matches[0]
            destination = None
            label = source.stem

        variant, promoted = self._store_external_variant(
            family,
            source,
            label=label,
            source_quality=source_quality,
            destination=destination,
            metadata=metadata,
            fingerprint=fingerprint,
            identity_verified=True,
        )
        return {
            "status": "created" if created else "matched",
            "source": str(source),
            "family_id": family["id"],
            "family_name": family["name"],
            "variant_id": variant["id"],
            "promoted": promoted,
        }

    def review_pptx_matches(
        self,
        input_pptx: Path,
        work_dir: Path,
        *,
        include_resolved: bool = False,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> list[dict[str, Any]]:
        """Extract PPTX media for explicit user review.

        The default preserves the legacy unresolved-only result. Full review
        includes exact and content matches without mutating the manifest.
        """
        input_pptx = input_pptx.expanduser().resolve()
        work_dir.mkdir(parents=True, exist_ok=True)
        scanned = scan_embedded_videos(input_pptx)
        reviewed: list[dict[str, Any]] = []
        with ZipFile(input_pptx, "r") as archive:
            for index, asset in enumerate(
                sorted(scanned.values(), key=lambda item: item.media_path), start=1
            ):
                _check_cancelled(cancel_callback)
                if progress_callback:
                    progress_callback(
                        f"正在分析待匹配视频 {index}/{len(scanned)}："
                        f"{Path(asset.media_path).name}"
                    )
                digest = _sha256_zip_member(archive, asset.media_path)
                family = self.family_by_known_hash(digest)
                match_kind = "exact" if family is not None else "unmatched"
                if family is not None and not include_resolved:
                    continue
                suffix = Path(asset.media_path).suffix.lower() or ".bin"
                candidate = _unique_path(
                    work_dir
                    / f"{input_pptx.stem}_{Path(asset.media_path).stem}{suffix}"
                )
                _copy_zip_member(archive, asset.media_path, candidate)
                metadata = probe_video(candidate)
                fingerprint = (
                    _video_fingerprint(candidate, metadata) if family is None else None
                )
                if family is None and fingerprint is not None:
                    family = self.family_by_content_fingerprint(fingerprint)
                    if family is not None:
                        match_kind = "content"
                if family is not None and not include_resolved:
                    candidate.unlink(missing_ok=True)
                    continue
                source_variant = None
                source_path = None
                target_error = ""
                if family is not None:
                    try:
                        source_variant = self.source_variant(family)
                        source_path = self.require_variant_path(source_variant)
                    except (FileNotFoundError, KeyError, ValueError) as exc:
                        target_error = str(exc)
                already_high_quality = bool(
                    source_variant
                    and digest == source_variant.get("sha256")
                    and not media_needs_mp4(asset.media_path)
                )
                reviewed.append(
                    {
                        "input_pptx": str(input_pptx),
                        "media_path": asset.media_path,
                        "source": str(candidate),
                        "sha256": digest,
                        "metadata": metadata,
                        "match_kind": match_kind,
                        "family_id": family["id"] if family else None,
                        "family_name": family["name"] if family else None,
                        "target_source": str(source_path) if source_path else None,
                        "target_sha256": (
                            source_variant.get("sha256") if source_variant else None
                        ),
                        "target_metadata": source_variant,
                        "target_error": target_error,
                        "already_high_quality": already_high_quality,
                        "occurrences": [
                            {
                                "slide_path": occurrence.slide_path,
                                "shape_id": occurrence.shape_id,
                            }
                            for occurrence in asset.occurrences
                        ],
                        "candidates": (
                            self._suggest_families(fingerprint)
                            if family is None and fingerprint is not None
                            else []
                        ),
                    }
                )
        return reviewed

    def compress_variant(
        self,
        variant_id: str,
        profile: str,
        *,
        destination: Path | None = None,
        activate: bool = True,
        progress_callback: Callable[[float, float, str], None] | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        if profile not in {"high", "balanced", "aggressive"}:
            raise ValueError(f"Unsupported compression profile: {profile}")
        family, source_variant = self.find_variant(variant_id)
        source = self.require_variant_path(source_variant)
        work = Path(tempfile.mkdtemp(prefix="video-version-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        persisted_target: Path | None = None
        output = work / f"compressed_{profile}.mp4"
        args = SimpleNamespace(
            input_pptx=source,
            profile=profile,
            target_size_mb=None,
            output=output,
            config=None,
            work_dir=None,
            max_height=720 if profile == "aggressive" else 1080,
            min_height=480,
            preset="medium",
            encoder="auto",
            keep_work_dir=False,
        )
        try:
            try:
                result = compact_standalone_video(
                    args,
                    logger=lambda message: LOGGER.info("compress %s", message),
                    progress_callback=progress_callback,
                    cancel_callback=cancel_callback,
                )
            except SystemExit as exc:
                raise RuntimeError(str(exc)) from None
            encoded = Path(result["output_pptx"])
            digest = sha256_file(encoded)
            for variant in family["variants"]:
                if variant["sha256"] == digest:
                    if activate:
                        family["active_variant_id"] = variant["id"]
                        self.save()
                        self.record(
                            "variant_activated",
                            family_id=family["id"],
                            variant_id=variant["id"],
                        )
                    return variant
            destination = (
                destination.expanduser().resolve() if destination else source.parent
            )
            destination.mkdir(parents=True, exist_ok=True)
            metadata = probe_video(encoded)
            target = _unique_path(
                destination
                / _variant_filename(
                    f"{family['name']}_{profile}",
                    metadata,
                    digest,
                    encoded.suffix.lower(),
                )
            )
            shutil.move(encoded, target)
            persisted_target = target
            variant = {
                "id": str(uuid.uuid4()),
                "label": profile,
                "profile": profile,
                "path": self.encode_path(target),
                "sha256": digest,
                "size_bytes": target.stat().st_size,
                "mtime_ns": target.stat().st_mtime_ns,
                "created_at": utc_now(),
                "source_variant_id": source_variant["id"],
                **metadata,
            }
            family["variants"].append(variant)
            if activate:
                family["active_variant_id"] = variant["id"]
            self.save()
            self.record(
                "variant_compressed",
                family_id=family["id"],
                variant_id=variant["id"],
                profile=profile,
            )
            return variant
        except Exception:
            self.data = original_data
            if persisted_target is not None:
                persisted_target.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def normalize_variant_to_mp4(
        self,
        variant_id: str,
        *,
        destination: Path | None = None,
    ) -> dict[str, Any]:
        family, source_variant = self.find_variant(variant_id)
        source = self.require_variant_path(source_variant)
        if source.suffix.lower() == ".mp4":
            return source_variant
        for variant in family["variants"]:
            if (
                variant.get("profile") == "mp4_high_fidelity"
                and variant.get("source_variant_id") == variant_id
            ):
                family["active_variant_id"] = variant["id"]
                self.save()
                return variant

        work = Path(tempfile.mkdtemp(prefix="video-normalize-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        persisted_target: Path | None = None
        encoded = work / "compatible.mp4"
        try:
            try:
                run_binary(
                    [
                        "ffmpeg",
                        "-y",
                        "-i",
                        str(source),
                        "-map",
                        "0:v:0",
                        "-map",
                        "0:a?",
                        "-c:v",
                        "libx264",
                        "-preset",
                        "medium",
                        "-crf",
                        "18",
                        "-profile:v",
                        "main",
                        "-pix_fmt",
                        "yuv420p",
                        "-c:a",
                        "aac",
                        "-b:a",
                        "256k",
                        "-movflags",
                        "+faststart",
                        str(encoded),
                    ],
                    capture=True,
                )
            except subprocess.CalledProcessError as exc:
                detail = (exc.stderr or "").strip().splitlines()
                raise RuntimeError(
                    f"FFmpeg could not create an MP4-compatible version: {detail[-1] if detail else exc}"
                ) from None
            digest = sha256_file(encoded)
            for variant in family["variants"]:
                if variant["sha256"] == digest:
                    family["active_variant_id"] = variant["id"]
                    self.save()
                    return variant
            destination = (
                destination.expanduser().resolve() if destination else source.parent
            )
            destination.mkdir(parents=True, exist_ok=True)
            metadata = probe_video(encoded)
            target = _unique_path(
                destination
                / _variant_filename(
                    f"{family['name']}_mp4_high_fidelity",
                    metadata,
                    digest,
                    ".mp4",
                )
            )
            shutil.move(encoded, target)
            persisted_target = target
            variant = {
                "id": str(uuid.uuid4()),
                "label": "mp4_high_fidelity",
                "profile": "mp4_high_fidelity",
                "path": self.encode_path(target),
                "sha256": digest,
                "size_bytes": target.stat().st_size,
                "mtime_ns": target.stat().st_mtime_ns,
                "created_at": utc_now(),
                "source_variant_id": source_variant["id"],
                **metadata,
            }
            family["variants"].append(variant)
            family["active_variant_id"] = variant["id"]
            self.save()
            self.record(
                "variant_normalized",
                family_id=family["id"],
                variant_id=variant["id"],
                source_variant_id=source_variant["id"],
            )
            return variant
        except Exception:
            self.data = original_data
            if persisted_target is not None:
                persisted_target.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def activate_variant(self, variant_id: str) -> None:
        family, _ = self.find_variant(variant_id)
        previous = family["active_variant_id"]
        family["active_variant_id"] = variant_id
        try:
            self.save()
        except Exception:
            family["active_variant_id"] = previous
            raise
        self.record("variant_activated", family_id=family["id"], variant_id=variant_id)

    def set_source_variant(self, variant_id: str) -> None:
        family, variant = self.find_variant(variant_id)
        fingerprint = _video_fingerprint(self.require_variant_path(variant), variant)
        previous = copy.deepcopy(family)
        family["source_variant_id"] = variant_id
        known = family.setdefault("known_hashes", [])
        if variant["sha256"] not in known:
            known.append(variant["sha256"])
        family["source_hashes"] = [variant["sha256"]]
        if fingerprint is not None:
            family["content_fingerprint"] = fingerprint
        else:
            family.pop("content_fingerprint", None)
        try:
            self.save()
        except Exception:
            family.clear()
            family.update(previous)
            raise
        self.record(
            "source_variant_selected", family_id=family["id"], variant_id=variant_id
        )

    def compatibility_warnings(self, variant_id: str) -> list[str]:
        family, variant = self.find_variant(variant_id)
        original = next(
            (item for item in family["variants"] if item.get("profile") == "original"),
            family["variants"][0],
        )
        warnings: list[str] = []
        original_duration = float(original.get("duration_sec") or 0)
        duration = float(variant.get("duration_sec") or 0)
        if (
            original_duration
            and duration
            and abs(duration - original_duration) > max(0.2, original_duration * 0.01)
        ):
            warnings.append(
                f"时长由 {original_duration:g}s 变为 {duration:g}s，依赖播放结束的动画可能改变。"
            )
        original_width = int(original.get("width") or 0)
        original_height = int(original.get("height") or 0)
        width = int(variant.get("width") or 0)
        height = int(variant.get("height") or 0)
        if original_width and original_height and width and height:
            original_ratio = original_width / original_height
            ratio = width / height
            if abs(ratio - original_ratio) / original_ratio > 0.01:
                warnings.append("宽高比与原片不同，回填后可能被拉伸或裁剪。")
        source_fingerprint = family.get("content_fingerprint")
        if source_fingerprint:
            candidate_fingerprint = _video_fingerprint(
                self.require_variant_path(variant), variant
            )
            if candidate_fingerprint is None:
                warnings.append("无法核实该版本的画面和音轨是否与当前高清源一致。")
            elif not _fingerprints_match(source_fingerprint, candidate_fingerprint):
                warnings.append(
                    "画面、音轨或时长与当前高清源不一致，可能是另一个视频。"
                )
        return warnings

    def merge_families(
        self,
        source_family_id: str,
        target_family_id: str,
        *,
        persist: bool = True,
        confirmed_same_content: bool = False,
    ) -> None:
        if source_family_id == target_family_id:
            return
        original_data = copy.deepcopy(self.data)
        source = self.family(source_family_id)
        target = self.family(target_family_id)
        if not confirmed_same_content and not self._families_have_same_content(
            source, target
        ):
            raise ValueError(
                "两个视频族未通过哈希、画面和音频指纹核实；"
                "必须人工确认是同一视频后才能归并"
            )
        hashes = {item["sha256"]: item for item in target["variants"]}
        variant_map: dict[str, str] = {}
        for variant in source["variants"]:
            duplicate = hashes.get(variant["sha256"])
            if duplicate:
                variant_map[variant["id"]] = duplicate["id"]
            else:
                hashes[variant["sha256"]] = variant
                variant_map[variant["id"]] = variant["id"]
        for deck in self.decks():
            for asset in deck["assets"]:
                if (
                    asset["family_id"] == source_family_id
                    and asset.get("original_variant_id") not in variant_map
                ):
                    raise ValueError(
                        "无法归并：演示文稿关联的原始视频版本不存在，请先修复库"
                    )
        target["variants"].extend(
            variant
            for variant in source["variants"]
            if variant_map[variant["id"]] == variant["id"]
        )
        target_known = target.setdefault("known_hashes", [])
        for digest in source.get("known_hashes", []):
            if digest not in target_known:
                target_known.append(digest)
        target_source = self.source_variant(target)
        target["source_hashes"] = [target_source["sha256"]]
        for deck in self.decks():
            for asset in deck["assets"]:
                if asset["family_id"] != source_family_id:
                    continue
                asset["family_id"] = target_family_id
                asset["original_variant_id"] = variant_map[asset["original_variant_id"]]
        self.data["families"] = [
            item for item in self.families() if item["id"] != source_family_id
        ]
        if persist:
            try:
                self.save()
            except Exception:
                self.data = original_data
                raise
            self.record(
                "families_merged", source=source_family_id, target=target_family_id
            )

    def family_merge_impact(
        self, source_family_id: str, target_family_id: str
    ) -> dict[str, Any]:
        source = self.family(source_family_id)
        self.family(target_family_id)
        affected_decks: list[str] = []
        reference_count = 0
        for deck in self.decks():
            count = sum(
                asset.get("family_id") == source_family_id
                for asset in deck.get("assets", [])
            )
            if count:
                affected_decks.append(str(deck.get("name") or deck["id"]))
                reference_count += count
        return {
            "deck_count": len(affected_decks),
            "deck_names": affected_decks,
            "reference_count": reference_count,
            "variant_count": len(source["variants"]),
            "known_hash_count": len(set(source.get("known_hashes", []))),
        }

    def relink_missing(
        self,
        search_roots: Iterable[Path] = (),
        *,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> list[dict[str, str]]:
        missing: dict[str, list[dict[str, Any]]] = {}
        expected_sizes: set[int] = set()
        metadata_updated = False
        for family in self.families():
            for variant in family["variants"]:
                path = self.variant_path(variant)
                valid = (
                    path.is_file()
                    and path.stat().st_size == variant["size_bytes"]
                    and sha256_file(path) == variant["sha256"]
                )
                if valid:
                    current_mtime = path.stat().st_mtime_ns
                    if variant.get("mtime_ns") != current_mtime:
                        variant["mtime_ns"] = current_mtime
                        metadata_updated = True
                    continue
                missing.setdefault(variant["sha256"], []).append(variant)
                expected_sizes.add(variant["size_bytes"])
        if not missing:
            if metadata_updated:
                self.save()
                self.record("variant_metadata_refreshed")
            return []

        roots = [
            self.root,
            *(Path(item).expanduser().resolve() for item in search_roots),
        ]
        results: list[dict[str, str]] = []
        seen: set[Path] = set()
        for root in roots:
            if not root.is_dir():
                continue
            for candidate in root.rglob("*"):
                _check_cancelled(cancel_callback)
                try:
                    resolved = candidate.resolve()
                    if (
                        resolved in seen
                        or not candidate.is_file()
                        or candidate.stat().st_size not in expected_sizes
                    ):
                        continue
                except OSError:
                    continue
                seen.add(resolved)
                if progress_callback:
                    progress_callback(f"正在校验 {candidate.name}")
                digest = sha256_file(candidate)
                variants = missing.pop(digest, [])
                for variant in variants:
                    variant["path"] = self.encode_path(candidate)
                    variant["mtime_ns"] = candidate.stat().st_mtime_ns
                    results.append(
                        {"variant_id": variant["id"], "path": variant["path"]}
                    )
                if not missing:
                    break
            if not missing:
                break
        if results or metadata_updated:
            self.save()
            self.record("variants_relinked", count=len(results))
        return results

    # ------------------------------------------------------------------
    # Library cleanup ("整理视频库")
    # ------------------------------------------------------------------

    @property
    def cleanup_dir(self) -> Path:
        return self.root / "_cleanup"

    @property
    def cleanup_index_path(self) -> Path:
        return self.cleanup_dir / "index.json"

    @staticmethod
    def _variant_quality_key(variant: dict[str, Any]) -> tuple[int, int, int]:
        width = int(variant.get("width") or 0)
        height = int(variant.get("height") or 0)
        return (
            width * height,
            int(variant.get("bitrate_kbps") or 0),
            int(variant.get("size_bytes") or 0),
        )

    def _cleanup_candidate(
        self,
        family: dict[str, Any],
        variant: dict[str, Any],
        *,
        ssim: float | None,
        confidence: dict[str, Any] | None,
        identity_known: bool,
        integrity_error: str = "",
    ) -> dict[str, Any]:
        try:
            self.require_variant_path(variant)
            file_ok = True
        except (FileNotFoundError, ValueError):
            file_ok = False
        block_reasons: list[str] = []
        if not file_ok:
            block_reasons.append("文件丢失或已被修改")
        if not identity_known and confidence is not None:
            if not confidence["duration_consistent"]:
                block_reasons.append("时长不一致（可能被剪辑）")
            if not confidence["audio_consistent"]:
                block_reasons.append("音轨不一致")
            if confidence["matched"] and confidence["level"] == "medium":
                block_reasons.append("匹配置信度中等")
            if not confidence["matched"]:
                block_reasons.append("置信度不足")
        return {
            "family_id": family["id"],
            "family_name": family["name"],
            "variant_id": variant["id"],
            "label": variant.get("label") or variant.get("profile") or "",
            "profile": variant.get("profile") or "",
            "path": variant["path"],
            "exists": file_ok,
            "sha256": variant["sha256"],
            "width": int(variant.get("width") or 0),
            "height": int(variant.get("height") or 0),
            "duration_sec": float(variant.get("duration_sec") or 0),
            "bitrate_kbps": int(variant.get("bitrate_kbps") or 0),
            "size_bytes": int(variant.get("size_bytes") or 0),
            "video_codec": str(variant.get("video_codec") or ""),
            "audio_codec": str(variant.get("audio_codec") or ""),
            "has_audio": bool(variant.get("has_audio")),
            "ssim_to_best": ssim,
            "confidence": confidence,
            "auto_allowed": not block_reasons,
            "can_keep": file_ok and not integrity_error,
            "integrity_error": integrity_error,
            "block_reasons": block_reasons,
        }

    def _named_damaged_copy_match(
        self,
        first: dict[str, Any],
        second: dict[str, Any],
        integrity_cache: dict[str, str] | None = None,
    ) -> tuple[bool, str, str]:
        if _normalized_cleanup_name(first["name"]) != _normalized_cleanup_name(
            second["name"]
        ):
            return False, "", ""
        first_variant = self.source_variant(first)
        second_variant = self.source_variant(second)
        if not all(
            int(first_variant.get(key) or 0) == int(second_variant.get(key) or 0) > 0
            for key in ("width", "height")
        ):
            return False, "", ""
        first_duration = float(first_variant.get("duration_sec") or 0)
        second_duration = float(second_variant.get("duration_sec") or 0)
        if not first_duration or abs(first_duration - second_duration) > 0.25:
            return False, "", ""
        try:
            first_path = self.require_variant_path(first_variant)
            second_path = self.require_variant_path(second_variant)
        except (FileNotFoundError, ValueError):
            return False, "", ""
        smaller = min(first_path.stat().st_size, second_path.stat().st_size)
        larger = max(first_path.stat().st_size, second_path.stat().st_size)
        if smaller < larger * 0.25:
            return False, "", ""
        cache = integrity_cache if integrity_cache is not None else {}

        def integrity(path: Path) -> str:
            key = str(path)
            if key not in cache:
                cache[key] = _video_packet_error(path)
            return cache[key]

        first_error = integrity(first_path)
        second_error = integrity(second_path)
        if bool(first_error) == bool(second_error):
            return False, first_error, second_error
        similarity = _aligned_file_similarity(first_path, second_path)
        return similarity >= 0.7, first_error, second_error

    def _named_deep_cleanup_match(
        self, first: dict[str, Any], second: dict[str, Any]
    ) -> bool:
        if _normalized_cleanup_name(first["name"]) != _normalized_cleanup_name(
            second["name"]
        ):
            return False
        first_fingerprint = self._family_fingerprint(first)
        second_fingerprint = self._family_fingerprint(second)
        if not first_fingerprint or not second_fingerprint:
            return False
        confidence = _fingerprint_confidence(first_fingerprint, second_fingerprint)
        frame_max = confidence["frame_max_distance"]
        frame_total = confidence["frame_total_distance"]
        luma_max = confidence["luma_max_difference"]
        first_aspect = _fingerprint_int(first_fingerprint, "aspect_ppm")
        second_aspect = _fingerprint_int(second_fingerprint, "aspect_ppm")
        if not bool(
            confidence["duration_consistent"]
            and first_aspect
            and second_aspect
            and abs(first_aspect - second_aspect)
            <= max(first_aspect, second_aspect) * 0.03
            and frame_max is not None
            and frame_max <= 25
            and frame_total is not None
            and frame_total <= 65
            and luma_max is not None
            and luma_max <= 15
        ):
            return False
        try:
            correlation = _decoded_audio_correlation(
                self.require_variant_path(self.source_variant(first)),
                self.require_variant_path(self.source_variant(second)),
            )
        except (FileNotFoundError, ValueError):
            return False
        return correlation is not None and correlation >= 0.98

    def _variant_matches_family_source(
        self, family: dict[str, Any], variant: dict[str, Any]
    ) -> bool:
        source = self.source_variant(family)
        if variant["id"] == source["id"] or variant["sha256"] == source["sha256"]:
            return True
        if variant.get("source_variant_id") in {
            item["id"] for item in family["variants"]
        }:
            return True
        source_fingerprint = family.get("content_fingerprint")
        if not source_fingerprint:
            return False
        try:
            candidate = _video_fingerprint(self.require_variant_path(variant), variant)
        except (FileNotFoundError, ValueError):
            return False
        return bool(candidate and _fingerprints_match(source_fingerprint, candidate))

    @staticmethod
    def _recommend_group(
        candidates: list[dict[str, Any]], best_id: str, ssim_threshold: float
    ) -> dict[str, Any]:
        best = next(item for item in candidates if item["variant_id"] == best_id)
        original = next(
            (item for item in candidates if item["profile"] == "original"), None
        )
        close_enough = []
        for item in candidates:
            if item["variant_id"] == best_id:
                continue
            if not item["auto_allowed"] or not item.get("can_keep", True):
                continue
            ssim = item["ssim_to_best"]
            if ssim is None or ssim < ssim_threshold:
                continue
            adequate_height = min(best["height"], 1080) if best["height"] else 0
            if adequate_height and item["height"] < adequate_height:
                continue
            close_enough.append(item)
        smallest_close = (
            min(close_enough, key=lambda item: item["size_bytes"])
            if close_enough
            else None
        )
        eligible = [item for item in candidates if item["auto_allowed"]]
        eligible = [item for item in eligible if item.get("can_keep", True)]
        return {
            "keep_variant_id": best_id,
            "strategy": "keep_best",
            "original_variant_id": (original["variant_id"] if original else None),
            "unify_available": len(eligible) >= 2,
            "alternatives": {
                "keep_best": best_id,
                "keep_original": (original["variant_id"] if original else None),
                "keep_smallest_close": (
                    smallest_close["variant_id"] if smallest_close else None
                ),
            },
        }

    def scan_cleanup_groups(
        self,
        *,
        ssim_threshold: float = 0.95,
        focus_family_id: str | None = None,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> list[dict[str, Any]]:
        """Find within-family redundant variants and cross-family duplicate
        families, score every candidate against the group's best version.
        Read-only: nothing is moved or deleted here."""
        groups: list[dict[str, Any]] = []

        def report(message: str) -> None:
            if progress_callback:
                progress_callback(message)

        # Pass 1: within-family groups. Imported custom versions still need
        # content verification; family membership alone is not identity.
        families = self.families()
        if focus_family_id is not None:
            self.family(focus_family_id)
        multi = [
            family
            for family in families
            if len(family["variants"]) >= 2
            and (focus_family_id is None or family["id"] == focus_family_id)
        ]
        for index, family in enumerate(multi, start=1):
            _check_cancelled(cancel_callback)
            report(f"正在评估族内版本 {index}/{len(multi)}：{family['name']}")
            best = self.source_variant(family)
            variants = [
                best,
                *(item for item in family["variants"] if item is not best),
            ]
            try:
                best_path = self.require_variant_path(best)
            except (FileNotFoundError, ValueError):
                best_path = None
            candidates: list[dict[str, Any]] = []
            for variant in variants:
                _check_cancelled(cancel_callback)
                identity_known = self._variant_matches_family_source(family, variant)
                ssim: float | None = None
                if variant["id"] == best["id"]:
                    ssim = 1.0
                elif identity_known and best_path is not None:
                    try:
                        candidate_path = self.require_variant_path(variant)
                        ssim = _ssim_videos(candidate_path, best_path)
                    except (FileNotFoundError, ValueError):
                        ssim = None
                candidates.append(
                    self._cleanup_candidate(
                        family,
                        variant,
                        ssim=ssim,
                        confidence=(
                            None
                            if identity_known
                            else {
                                "matched": False,
                                "level": "low",
                                "duration_consistent": False,
                                "audio_consistent": False,
                            }
                        ),
                        identity_known=identity_known,
                    )
                )
            safe_to_apply = all(item["auto_allowed"] for item in candidates)
            groups.append(
                {
                    "kind": "within_family",
                    "family_ids": [family["id"]],
                    "title": f"{family['name']}（{len(variants)} 个版本）",
                    "best_variant_id": best["id"],
                    "candidates": candidates,
                    "safe_to_apply": safe_to_apply,
                    "recommendation": self._recommend_group(
                        candidates, best["id"], ssim_threshold
                    ),
                }
            )

        # Pass 2: cross-family clusters via exact hashes + fingerprints.
        report("正在比对跨族重复")
        parent = {family["id"]: family["id"] for family in families}

        def find(node: str) -> str:
            while parent[node] != node:
                parent[node] = parent[parent[node]]
                node = parent[node]
            return node

        def union(a: str, b: str) -> None:
            root_a, root_b = find(a), find(b)
            if root_a != root_b:
                parent[root_b] = root_a

        by_hash: dict[str, list[str]] = {}
        for family in families:
            for digest in family.get("known_hashes", []):
                by_hash.setdefault(digest, []).append(family["id"])
        for ids in by_hash.values():
            for other in ids[1:]:
                union(ids[0], other)
        fingerprints = {
            family["id"]: self._family_fingerprint(family) for family in families
        }
        integrity_errors: dict[str, str] = {}
        fingerprinted = [family for family in families if fingerprints[family["id"]]]
        for index, first in enumerate(fingerprinted):
            _check_cancelled(cancel_callback)
            for second in fingerprinted[index + 1 :]:
                if find(first["id"]) == find(second["id"]):
                    continue
                if _fingerprints_match(
                    fingerprints[first["id"]], fingerprints[second["id"]]
                ) or _named_cleanup_match(
                    first["name"],
                    second["name"],
                    fingerprints[first["id"]],
                    fingerprints[second["id"]],
                ):
                    union(first["id"], second["id"])

        by_name: dict[str, list[dict[str, Any]]] = {}
        for family in families:
            name = _normalized_cleanup_name(family["name"])
            if name and name not in {"video", "normal_video", "视频", "媒体1", "movie"}:
                by_name.setdefault(name, []).append(family)
        for named_families in by_name.values():
            for index, first in enumerate(named_families):
                for second in named_families[index + 1 :]:
                    if find(first["id"]) == find(second["id"]):
                        continue
                    matched, _, _ = self._named_damaged_copy_match(
                        first, second, integrity_errors
                    )
                    if matched:
                        union(first["id"], second["id"])
                        continue
                    if self._named_deep_cleanup_match(first, second):
                        union(first["id"], second["id"])

        clusters: dict[str, list[dict[str, Any]]] = {}
        for family in families:
            clusters.setdefault(find(family["id"]), []).append(family)
        cross: list[list[dict[str, Any]]] = []
        for members in clusters.values():
            remaining = list(members)
            while remaining:
                anchor = max(
                    remaining,
                    key=lambda family: self._variant_quality_key(
                        self.source_variant(family)
                    ),
                )
                group = [
                    family
                    for family in remaining
                    if family is anchor
                    or self._families_have_same_content(
                        anchor, family, integrity_cache=integrity_errors
                    )
                ]
                remaining = [family for family in remaining if family not in group]
                if len(group) >= 2:
                    cross.append(group)
        if focus_family_id is not None:
            cross = [
                members
                for members in cross
                if any(family["id"] == focus_family_id for family in members)
            ]
        for group_index, members in enumerate(cross, start=1):
            _check_cancelled(cancel_callback)
            title = " ⇄ ".join(family["name"] for family in members)
            report(f"正在评估跨族重复 {group_index}/{len(cross)}：{title}")
            sources = [self.source_variant(family) for family in members]
            ordered = sorted(
                zip(members, sources),
                key=lambda pair: (
                    not bool(integrity_errors.get(str(self.variant_path(pair[1])))),
                    self._variant_quality_key(pair[1]),
                ),
                reverse=True,
            )
            best_family, best_variant = ordered[0]
            best_fingerprint = fingerprints[best_family["id"]]
            if best_fingerprint is None:
                try:
                    best_fingerprint = _video_fingerprint(
                        self.require_variant_path(best_variant), best_variant
                    )
                except (FileNotFoundError, ValueError):
                    best_fingerprint = None
            try:
                best_path = self.require_variant_path(best_variant)
            except (FileNotFoundError, ValueError):
                best_path = None
            candidates = []
            for family, variant in ordered:
                _check_cancelled(cancel_callback)
                if variant["id"] == best_variant["id"]:
                    confidence = {
                        "matched": True,
                        "level": "high",
                        "duration_consistent": True,
                        "audio_consistent": True,
                        "duration_diff_ms": 0,
                        "frame_max_distance": 0,
                        "audio_distance": 0,
                    }
                    ssim = 1.0
                else:
                    other_fingerprint = fingerprints[family["id"]]
                    if other_fingerprint is None:
                        try:
                            other_fingerprint = _video_fingerprint(
                                self.require_variant_path(variant), variant
                            )
                        except (FileNotFoundError, ValueError):
                            other_fingerprint = None
                    confidence = (
                        _fingerprint_confidence(best_fingerprint, other_fingerprint)
                        if best_fingerprint and other_fingerprint
                        else {
                            "matched": False,
                            "level": "low",
                            "duration_consistent": False,
                            "audio_consistent": False,
                            "duration_diff_ms": None,
                            "frame_max_distance": None,
                            "audio_distance": None,
                        }
                    )
                    identity_known = self._families_have_same_content(
                        best_family, family
                    )
                    ssim = None
                    if identity_known and best_path is not None:
                        try:
                            ssim = _ssim_videos(
                                self.require_variant_path(variant), best_path
                            )
                        except (FileNotFoundError, ValueError):
                            ssim = None
                candidates.append(
                    self._cleanup_candidate(
                        family,
                        variant,
                        ssim=ssim,
                        confidence=confidence,
                        identity_known=(
                            variant["id"] == best_variant["id"] or identity_known
                        ),
                        integrity_error=integrity_errors.get(
                            str(self.variant_path(variant)), ""
                        ),
                    )
                )
            groups.append(
                {
                    "kind": "cross_family",
                    "family_ids": [family["id"] for family, _ in ordered],
                    "title": f"跨族重复：{title}",
                    "best_variant_id": best_variant["id"],
                    "candidates": candidates,
                    "safe_to_apply": all(item["auto_allowed"] for item in candidates),
                    "recommendation": self._recommend_group(
                        candidates, best_variant["id"], ssim_threshold
                    ),
                }
            )

        cross_family_ids = {
            family_id
            for group in groups
            if group["kind"] == "cross_family"
            for family_id in group["family_ids"]
        }
        groups = [
            group
            for group in groups
            if group["kind"] == "cross_family"
            or not cross_family_ids.intersection(group["family_ids"])
        ]

        def savings(group: dict[str, Any]) -> int:
            keep = group["recommendation"]["keep_variant_id"]
            return sum(
                item["size_bytes"]
                for item in group["candidates"]
                if item["variant_id"] != keep
            )

        groups.sort(key=savings, reverse=True)
        return groups

    # -- quarantine ----------------------------------------------------

    def _read_cleanup_index(self) -> list[dict[str, Any]]:
        path = self.cleanup_index_path
        if not path.is_file():
            return []
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise RuntimeError(f"待清理索引损坏，拒绝继续操作：{path}") from exc
        if not isinstance(data, list) or not all(
            isinstance(item, dict) for item in data
        ):
            raise RuntimeError(f"待清理索引格式无效，拒绝继续操作：{path}")
        return data

    def _cleanup_entry_path(self, entry: dict[str, Any]) -> Path:
        value = entry.get("quarantined_path")
        if not isinstance(value, str) or not value:
            raise RuntimeError("待清理索引缺少隔离文件路径")
        stored = Path(value).expanduser()
        path = (stored if stored.is_absolute() else self.root / stored).resolve()
        try:
            path.relative_to(self.cleanup_dir.resolve())
        except ValueError:
            raise RuntimeError(f"待清理索引包含越界路径，拒绝操作：{path}") from None
        return path

    def _cleanup_entry_original_path(self, entry: dict[str, Any]) -> Path:
        value = entry.get("original_path")
        if not isinstance(value, str) or not value:
            raise RuntimeError("待清理索引缺少原文件路径")
        stored = Path(value).expanduser()
        path = (stored if stored.is_absolute() else self.root / stored).resolve()
        try:
            path.relative_to((self.root / "media").resolve())
        except ValueError:
            raise RuntimeError(f"待清理索引包含越界路径，拒绝操作：{path}") from None
        return path

    def _write_cleanup_index(self, entries: list[dict[str, Any]]) -> None:
        self.cleanup_dir.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=".cleanup-index-", suffix=".json", dir=self.cleanup_dir
        )
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            temp_path.write_text(
                json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            os.replace(temp_path, self.cleanup_index_path)
        finally:
            temp_path.unlink(missing_ok=True)

    def _recover_pending_cleanup_moves(self) -> None:
        """Finish or discard a cleanup intent left by an interrupted move."""

        if not self.cleanup_index_path.is_file():
            return
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            recovered: list[dict[str, Any]] = []
            changed = False
            for entry in entries:
                if entry.get("state") != "moving":
                    recovered.append(entry)
                    continue
                source = self._cleanup_entry_original_path(entry)
                target = self._cleanup_entry_path(entry)
                expected_hash = (entry.get("variant") or {}).get("sha256")
                if (
                    target.is_file()
                    and not source.exists()
                    and expected_hash
                    and sha256_file(target) == expected_hash
                ):
                    recovered.append({**entry, "state": "quarantined"})
                    changed = True
                elif source.is_file() and not target.exists():
                    # The process stopped before the atomic in-library rename.
                    changed = True
                else:
                    # Keep an unresolved intent visible rather than concealing a loss.
                    recovered.append(entry)
            if changed:
                self._write_cleanup_index(recovered)

    def _quarantine_variant_file(
        self,
        family: dict[str, Any],
        variant: dict[str, Any],
        *,
        reason: str,
    ) -> Path:
        source = self.variant_path(variant)
        if not source.is_file():
            raise FileNotFoundError(source)
        try:
            source.relative_to((self.root / "media").resolve())
        except ValueError:
            raise ValueError(f"拒绝隔离视频库 media 目录外的文件：{source}") from None
        token = uuid.uuid4().hex[:12]
        target = _unique_path(self.cleanup_dir / f"{token}_{source.name}")
        entry = {
            "token": token,
            "family_id": family["id"],
            "family_name": family["name"],
            "variant": copy.deepcopy(variant),
            "original_path": self.encode_path(source),
            "quarantined_path": self.encode_path(target),
            "quarantined_at": utc_now(),
            "reason": reason,
            "state": "moving",
        }
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            self.cleanup_dir.mkdir(parents=True, exist_ok=True)
            entries = self._read_cleanup_index()
            self._write_cleanup_index([*entries, entry])
            try:
                shutil.move(str(source), target)
                self._write_cleanup_index([*entries, {**entry, "state": "quarantined"}])
            except Exception:
                restored = source.is_file() and not target.exists()
                if target.is_file() and not source.exists():
                    try:
                        source.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(target), source)
                        restored = True
                    except OSError:
                        LOGGER.exception("Unable to roll back video quarantine move")
                if restored:
                    self._write_cleanup_index(entries)
                raise
        return target

    def _remove_cleanup_entries_for_paths(self, paths: Iterable[Path]) -> None:
        resolved = {path.resolve() for path in paths}
        if not resolved:
            return
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            self._write_cleanup_index(
                [
                    entry
                    for entry in entries
                    if self._cleanup_entry_path(entry) not in resolved
                ]
            )

    def pending_cleanup(self) -> list[dict[str, Any]]:
        entries = []
        for entry in self._read_cleanup_index():
            path = self._cleanup_entry_path(entry)
            entries.append(
                {
                    **entry,
                    "exists": path.is_file(),
                    "size_bytes": path.stat().st_size if path.is_file() else 0,
                }
            )
        return entries

    def quarantine_abnormal_variant(self, variant_id: str) -> dict[str, Any]:
        """Isolate one unreadable, unused non-source version."""
        family, variant = self.find_variant(variant_id)
        if not variant.get("probe_error") and self.status(variant) == "available":
            raise ValueError("只能隔离文件异常的视频版本")
        if variant_id in {
            family.get("source_variant_id"),
            family.get("active_variant_id"),
        }:
            raise ValueError("该版本仍是高清源或当前版本，请先切换到健康版本")
        references = sum(
            asset.get("original_variant_id") == variant_id
            for deck in self.decks()
            for asset in deck.get("assets", [])
        )
        if references:
            raise ValueError("该版本仍被 PPTX 引用，拒绝隔离")
        fallback = self.source_variant(family)
        if fallback["id"] == variant_id:
            raise ValueError("视频族没有可用的健康版本")
        original_data = copy.deepcopy(self.data)
        original_path = self.variant_path(variant)
        quarantined: Path | None = None
        try:
            removed = self._remove_variant_from_family(
                family, variant_id, fallback_variant_id=fallback["id"]
            )
            quarantined = self._quarantine_variant_file(
                family, removed, reason="人工确认：隔离文件异常版本"
            )
            self.save()
        except Exception:
            self.data = original_data
            restored: list[Path] = []
            if quarantined is not None and quarantined.is_file():
                original_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(quarantined), original_path)
                restored.append(quarantined)
            self._remove_cleanup_entries_for_paths(restored)
            raise
        self.record(
            "abnormal_variant_quarantined",
            family_id=family["id"],
            variant_id=variant_id,
        )
        return {"family_id": family["id"], "variant_id": variant_id}

    def restore_cleanup_entry(self, token: str) -> Path:
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            entry = next((item for item in entries if item.get("token") == token), None)
            if entry is None:
                raise KeyError(f"Unknown cleanup entry: {token}")
            quarantined = self._cleanup_entry_path(entry)
            target = self._cleanup_entry_original_path(entry)
            variant = copy.deepcopy(entry["variant"])
            original_data = copy.deepcopy(self.data)
            moved_to: Path | None = None
            try:
                family = self.family(entry["family_id"])
            except KeyError:
                family = None  # type: ignore[assignment]
            if family is None:
                raise RuntimeError(
                    "原视频族已不存在（可能已归并或移除），无法还原："
                    f"{entry['family_name']}"
                )
            existing = next(
                (item for item in family["variants"] if item["id"] == variant["id"]),
                None,
            )
            if not quarantined.is_file():
                if existing is not None:
                    existing_path = self.variant_path(existing)
                    if (
                        existing_path.is_file()
                        and sha256_file(existing_path) == variant["sha256"]
                    ):
                        self._write_cleanup_index(
                            [item for item in entries if item.get("token") != token]
                        )
                        return existing_path
                raise FileNotFoundError(quarantined)
            if sha256_file(quarantined) != variant["sha256"]:
                raise ValueError("待清理文件内容已变化，拒绝还原")
            if existing is not None:
                if self.variant_path(existing) == target and not target.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(quarantined), target)
                    self._write_cleanup_index(
                        [item for item in entries if item.get("token") != token]
                    )
                    return target
                raise RuntimeError("该版本已存在于视频库，拒绝重复还原")
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                target = _unique_path(target)
                shutil.move(str(quarantined), target)
                moved_to = target
                variant["path"] = self.encode_path(target)
                variant["mtime_ns"] = target.stat().st_mtime_ns
                variant["size_bytes"] = target.stat().st_size
                family["variants"].append(variant)
                self.save()
            except Exception:
                self.data = original_data
                if (
                    moved_to is not None
                    and moved_to.is_file()
                    and not quarantined.is_file()
                ):
                    try:
                        shutil.move(str(moved_to), quarantined)
                    except OSError as rollback_error:
                        LOGGER.error(
                            "Unable to roll back cleanup restore: %s", rollback_error
                        )
                raise
            self._write_cleanup_index(
                [item for item in entries if item.get("token") != token]
            )
        self.record(
            "cleanup_restored",
            token=token,
            family_id=family["id"],
            variant_id=variant["id"],
        )
        return moved_to

    def cleanup_pending_issues(self) -> list[str]:
        issues: list[str] = []
        entries = self._read_cleanup_index()
        if not entries:
            return issues
        # Build a lookup of variants still present in the library so we can
        # recognise entries whose restore committed to the manifest but whose
        # index removal was interrupted (crash window between save() and
        # _write_cleanup_index in restore_cleanup_entry). Such an entry is
        # already restored: its variant is back in a family, the media file
        # exists and matches the recorded hash. Treating it as an issue would
        # be a false alarm, so we silently skip it instead of deleting the
        # stale entry here (removal stays the responsibility of restore/empty).
        restored_variants: dict[str, dict[str, Any]] = {}
        for family in self.families():
            for variant in family["variants"]:
                restored_variants[variant["id"]] = variant
        referenced_variants: set[str] = set(restored_variants)
        referenced_variants.discard("")
        for family in self.families():
            referenced_variants.add(family.get("active_variant_id") or "")
            referenced_variants.add(family.get("source_variant_id") or "")
        for deck in self.decks():
            for asset in deck.get("assets", []):
                referenced_variants.add(asset.get("original_variant_id") or "")
        for entry in entries:
            variant_snapshot = entry.get("variant") or {}
            variant_id = variant_snapshot.get("id", "")
            existing = restored_variants.get(variant_id) if variant_id else None
            if existing is not None:
                try:
                    existing_path = self.variant_path(existing)
                    if (
                        existing_path.is_file()
                        and variant_snapshot.get("sha256")
                        and sha256_file(existing_path) == variant_snapshot["sha256"]
                    ):
                        # Already restored to the library; the index entry is
                        # a stale leftover, not a real issue.
                        continue
                except (FileNotFoundError, ValueError):
                    pass
            if variant_id and variant_id in referenced_variants:
                issues.append(
                    f"{entry.get('family_name', '?')} 的版本仍被清单引用：{variant_id[:8]}"
                )
            if entry.get("state") == "moving":
                issues.append(f"视频隔离操作未完成：{entry.get('token', '?')}")
            path = self._cleanup_entry_path(entry)
            if not path.is_file():
                issues.append(f"待清理文件丢失：{path}")
        return issues

    def empty_cleanup(self) -> int:
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            issues = self.cleanup_pending_issues()
            if issues:
                raise RuntimeError("待清理目录不满足清空条件：\n" + "\n".join(issues))
            entries = self._read_cleanup_index()
            removed = 0
            for entry in entries:
                path = self._cleanup_entry_path(entry)
                if path.is_file():
                    path.unlink()
                    removed += 1
            self._write_cleanup_index([])
            try:
                self.cleanup_index_path.unlink(missing_ok=True)
                self.cleanup_dir.rmdir()
            except OSError:
                pass
        self.record("cleanup_emptied", count=removed)
        return removed

    # -- plan application ------------------------------------------------

    def _remove_variant_from_family(
        self,
        family: dict[str, Any],
        variant_id: str,
        *,
        fallback_variant_id: str,
    ) -> dict[str, Any]:
        variant = next(
            (item for item in family["variants"] if item["id"] == variant_id), None
        )
        if variant is None:
            raise KeyError(f"Unknown video version: {variant_id}")
        if len(family["variants"]) <= 1:
            raise RuntimeError("视频族至少需要保留一个版本")
        known = family.setdefault("known_hashes", [])
        if variant["sha256"] not in known:
            known.append(variant["sha256"])
        family["variants"] = [
            item for item in family["variants"] if item["id"] != variant_id
        ]
        if family.get("active_variant_id") == variant_id:
            family["active_variant_id"] = fallback_variant_id
        if family.get("source_variant_id") == variant_id:
            family["source_variant_id"] = fallback_variant_id
            fallback = next(
                item for item in family["variants"] if item["id"] == fallback_variant_id
            )
            family["source_hashes"] = [fallback["sha256"]]
        for deck in self.decks():
            for asset in deck.get("assets", []):
                if (
                    asset.get("family_id") == family["id"]
                    and asset.get("original_variant_id") == variant_id
                ):
                    asset["original_variant_id"] = fallback_variant_id
        return variant

    def create_unified_version(
        self,
        family_id: str,
        *,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
        persist: bool = True,
    ) -> dict[str, Any]:
        """Encode a high-quality 1080p MP4 from the authoritative source."""
        _check_cancelled(cancel_callback)
        family = self.family(family_id)
        best = self.source_variant(family)
        best_path = self.require_variant_path(best)
        width = int(best.get("width") or 0)
        height = int(best.get("height") or 0)
        if width <= 0 or height <= 0:
            metadata = probe_video(best_path)
            width = int(metadata.get("width") or 0)
            height = int(metadata.get("height") or 0)
        if width <= 0 or height <= 0:
            raise RuntimeError(f"无法读取视频尺寸：{best_path}")
        if progress_callback:
            progress_callback(
                f"正在从 {best.get('label', '最佳版本')} 生成 1080p 统一版本"
            )
        work = Path(tempfile.mkdtemp(prefix="video-unified-", dir=self.root))
        original_data = copy.deepcopy(self.data)
        persisted: Path | None = None
        try:
            produced = work / "unified_1080p.mp4"
            _transcode_high_quality_mp4(
                best_path, produced, width, height, family_id=family["id"]
            )
            _check_cancelled(cancel_callback)
            digest = sha256_file(produced)
            existing = self.find_variant_by_hash(digest)
            if existing is not None and existing[0]["id"] == family_id:
                if persist:
                    self.set_source_variant(existing[1]["id"])
                else:
                    family["source_variant_id"] = existing[1]["id"]
                    family["source_hashes"] = [existing[1]["sha256"]]
                return existing[1]
            family_dir = self.variant_path(best).parent
            metadata = probe_video(produced)
            target = _unique_path(
                family_dir
                / _variant_filename(
                    f"{family['name']}_unified1080p",
                    metadata,
                    digest,
                    ".mp4",
                )
            )
            shutil.move(str(produced), target)
            persisted = target
            variant = {
                "id": str(uuid.uuid4()),
                "label": "1080p 统一版",
                "profile": "1080p_source",
                "path": self.encode_path(target),
                "sha256": digest,
                "size_bytes": target.stat().st_size,
                "mtime_ns": target.stat().st_mtime_ns,
                "created_at": utc_now(),
                "source_variant_id": best["id"],
                **metadata,
            }
            family["variants"].append(variant)
            family["source_variant_id"] = variant["id"]
            family["source_hashes"] = [variant["sha256"]]
            if variant["sha256"] not in family.setdefault("known_hashes", []):
                family["known_hashes"].append(variant["sha256"])
            fingerprint = _video_fingerprint(target, variant)
            if fingerprint is not None:
                family["content_fingerprint"] = fingerprint
            if persist:
                self.save()
                self.record(
                    "cleanup_unified_created",
                    family_id=family_id,
                    variant_id=variant["id"],
                    from_variant_id=best["id"],
                )
            return variant
        except Exception:
            self.data = original_data
            if persisted is not None:
                persisted.unlink(missing_ok=True)
            raise
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def _families_have_same_content(
        self,
        first: dict[str, Any],
        second: dict[str, Any],
        *,
        integrity_cache: dict[str, str] | None = None,
    ) -> bool:
        first_hashes = {item["sha256"] for item in first["variants"]}
        second_hashes = {item["sha256"] for item in second["variants"]}
        if first_hashes.intersection(second_hashes):
            return True
        first_fingerprint = self._family_fingerprint(first)
        second_fingerprint = self._family_fingerprint(second)
        fingerprint_match = bool(
            first_fingerprint
            and second_fingerprint
            and (
                _fingerprints_match(first_fingerprint, second_fingerprint)
                or _named_cleanup_match(
                    first["name"],
                    second["name"],
                    first_fingerprint,
                    second_fingerprint,
                )
            )
        )
        if fingerprint_match and (
            first_fingerprint.get("has_audio") or second_fingerprint.get("has_audio")
        ):
            try:
                correlation = _decoded_audio_correlation(
                    self.require_variant_path(self.source_variant(first)),
                    self.require_variant_path(self.source_variant(second)),
                )
            except (FileNotFoundError, ValueError):
                correlation = None
            fingerprint_match = correlation is not None and correlation >= 0.98
        return (
            fingerprint_match
            or self._named_damaged_copy_match(
                first, second, integrity_cache=integrity_cache
            )[0]
            or self._named_deep_cleanup_match(first, second)
        )

    def _validate_cleanup_decision(self, decision: dict[str, Any]) -> None:
        kind = decision.get("kind")
        remove_ids = list(dict.fromkeys(decision.get("remove_variant_ids", [])))
        force_remove_ids = list(
            dict.fromkeys(decision.get("force_remove_variant_ids", []))
        )
        keep_id = decision.get("keep_variant_id")
        if kind == "within_family":
            family = self.family(decision["family_id"])
            family_ids = {item["id"] for item in family["variants"]}
            if keep_id not in family_ids or not set(remove_ids).issubset(family_ids):
                raise ValueError("整理决策包含不属于该视频族的版本")
            if not set(force_remove_ids).issubset(remove_ids):
                raise ValueError("强制整理版本必须同时出现在移除列表中")
            if force_remove_ids and not self._variant_matches_family_source(
                family,
                next(item for item in family["variants"] if item["id"] == keep_id),
            ):
                raise ValueError("强制整理必须保留已核实为同一内容的版本")
            for variant_id in [keep_id, *remove_ids]:
                variant = next(
                    item for item in family["variants"] if item["id"] == variant_id
                )
                if (
                    variant_id not in force_remove_ids
                    and not self._variant_matches_family_source(family, variant)
                ):
                    raise ValueError("存在未核实为同一内容的视频版本，拒绝自动整理")
            return
        if kind != "cross_family":
            raise ValueError(f"未知的整理类型：{kind}")
        if force_remove_ids:
            raise ValueError("跨族归并不支持强制整理")

        target = self.family(decision["merge_into_family_id"])
        family_ids = list(dict.fromkeys(decision.get("merge_family_ids", [])))
        if target["id"] not in family_ids or len(family_ids) < 2:
            raise ValueError("跨族整理决策缺少有效的目标视频族")
        families = [self.family(family_id) for family_id in family_ids]
        target_variant_ids = {item["id"] for item in target["variants"]}
        if keep_id not in target_variant_ids:
            raise ValueError("保留版本不属于跨族整理的目标视频族")
        all_variant_ids = {
            item["id"] for family in families for item in family["variants"]
        }
        if not set(remove_ids).issubset(all_variant_ids):
            raise ValueError("跨族整理决策包含未知视频版本")
        for family in families:
            if family["id"] != target["id"] and not self._families_have_same_content(
                target, family
            ):
                raise ValueError("视频族内容身份未通过哈希或指纹核实，拒绝归并")

    def apply_cleanup_plan(
        self,
        decisions: list[dict[str, Any]],
        *,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> dict[str, Any]:
        """Apply per-group keep decisions. Each decision:
        {kind, family_id?, keep_variant_id, remove_variant_ids,
         merge_into_family_id?, merge_family_ids?, unify_first?}
        Files are only moved to the quarantine dir, never deleted. Each group
        is atomic: a failing group rolls back and does not stop the others."""
        results: list[dict[str, Any]] = []
        for index, decision in enumerate(decisions, start=1):
            _check_cancelled(cancel_callback)
            kind = decision.get("kind")
            original_data = copy.deepcopy(self.data)
            quarantined_files: list[tuple[Path, Path]] = []
            created_files: list[Path] = []
            try:
                self._validate_cleanup_decision(decision)
                effective_keep_id = decision["keep_variant_id"]
                # A decision that removes nothing, merges no families and does
                # not unify is a no-op. Skip the save()/record() so we don't
                # burn a manifest revision or emit a misleading audit entry.
                merge_family_ids = list(
                    dict.fromkeys(decision.get("merge_family_ids", []))
                )
                has_cross_merge = kind == "cross_family" and len(merge_family_ids) >= 2
                remove_ids = [
                    item
                    for item in decision.get("remove_variant_ids", [])
                    if item != effective_keep_id
                ]
                if (
                    not remove_ids
                    and not has_cross_merge
                    and not decision.get("unify_first")
                ):
                    results.append(
                        {
                            "ok": True,
                            "removed": 0,
                            "skipped": True,
                            "decision": decision,
                        }
                    )
                    continue
                if decision.get("unify_first"):
                    if kind == "cross_family":
                        unify_family_id = decision["merge_into_family_id"]
                    else:
                        unify_family_id = decision["family_id"]
                    if progress_callback:
                        progress_callback("正在生成统一 1080p 版本")
                    previous_ids = {
                        item["id"] for item in self.family(unify_family_id)["variants"]
                    }
                    unified = self.create_unified_version(
                        unify_family_id,
                        progress_callback=progress_callback,
                        cancel_callback=cancel_callback,
                        persist=False,
                    )
                    if unified["id"] not in previous_ids:
                        created_files.append(self.variant_path(unified))
                    if kind != "cross_family":
                        effective_keep_id = unified["id"]
                if kind == "cross_family":
                    target_id = decision["merge_into_family_id"]
                    for loser_id in decision.get("merge_family_ids", []):
                        _check_cancelled(cancel_callback)
                        if loser_id == target_id:
                            continue
                        if progress_callback:
                            progress_callback(
                                f"正在归并视频族 {loser_id[:8]} → {target_id[:8]}"
                            )
                        target_family = self.family(target_id)
                        loser_family = self.family(loser_id)
                        target_hashes = {
                            item["sha256"] for item in target_family["variants"]
                        }
                        for duplicate in loser_family["variants"]:
                            if duplicate["sha256"] not in target_hashes:
                                continue
                            original_path = self.variant_path(duplicate)
                            quarantined = self._quarantine_variant_file(
                                target_family,
                                duplicate,
                                reason=decision.get("reason")
                                or "整理视频库：归并完全重复文件",
                            )
                            quarantined_files.append((original_path, quarantined))
                        self.merge_families(
                            loser_id,
                            target_id,
                            persist=False,
                            confirmed_same_content=True,
                        )
                    effective_keep_id = self.source_variant(self.family(target_id))[
                        "id"
                    ]
                removed: list[dict[str, Any]] = []
                for variant_id in decision.get("remove_variant_ids", []):
                    _check_cancelled(cancel_callback)
                    if variant_id == effective_keep_id:
                        continue
                    try:
                        family, _variant = self.find_variant(variant_id)
                    except KeyError:
                        # The variant record may already be gone (e.g. dropped
                        # as an exact-hash duplicate during merge). Nothing to
                        # remove from the manifest; its file, if any, was
                        # handled by the merge. Skip rather than fail the group.
                        LOGGER.info(
                            "Cleanup removal skipped, variant no longer present: %s",
                            variant_id,
                        )
                        continue
                    removed_variant = self._remove_variant_from_family(
                        family, variant_id, fallback_variant_id=effective_keep_id
                    )
                    try:
                        target = self._quarantine_variant_file(
                            family,
                            removed_variant,
                            reason=decision.get("reason") or "整理视频库",
                        )
                        quarantined_files.append(
                            (self.variant_path(removed_variant), target)
                        )
                    except FileNotFoundError:
                        LOGGER.warning(
                            "Variant file already missing, skipped quarantine: %s",
                            removed_variant.get("path"),
                        )
                    removed.append(removed_variant)
                self.save()
                self.record(
                    "cleanup_plan_applied",
                    kind=kind,
                    removed=len(removed),
                    keep=effective_keep_id,
                )
                results.append(
                    {"ok": True, "removed": len(removed), "decision": decision}
                )
            except Exception as exc:
                self.data = original_data
                for path in created_files:
                    path.unlink(missing_ok=True)
                restored_quarantine_paths: list[Path] = []
                for original_path, quarantined in reversed(quarantined_files):
                    if quarantined.is_file() and not original_path.is_file():
                        try:
                            original_path.parent.mkdir(parents=True, exist_ok=True)
                            shutil.move(str(quarantined), original_path)
                            restored_quarantine_paths.append(quarantined)
                        except OSError as rollback_error:
                            LOGGER.error(
                                "Unable to roll back quarantine move: %s",
                                rollback_error,
                            )
                    elif original_path.is_file() and not quarantined.exists():
                        restored_quarantine_paths.append(quarantined)
                self._remove_cleanup_entries_for_paths(restored_quarantine_paths)
                results.append({"ok": False, "error": str(exc), "decision": decision})
        return {
            "applied": sum(1 for item in results if item["ok"]),
            "failed": sum(1 for item in results if not item["ok"]),
            "results": results,
        }

    def build_placeholder(
        self, source: Path, target: Path, variant: dict[str, Any]
    ) -> None:
        duration = float(variant.get("duration_sec") or 0)
        width = int(variant.get("width") or 0)
        height = int(variant.get("height") or 0)
        if duration <= 0 or width <= 0 or height <= 0:
            metadata = probe_video(source)
            duration = float(metadata.get("duration_sec") or 0)
            width = int(metadata.get("width") or 0)
            height = int(metadata.get("height") or 0)
        if duration <= 0 or width <= 0 or height <= 0:
            raise RuntimeError(
                f"Cannot create a timing-safe placeholder because video metadata is unavailable: {source}"
            )
        duration = max(0.2, duration)
        scale = min(1.0, 640 / width)
        output_width = max(2, int(width * scale) // 2 * 2)
        output_height = max(2, int(height * scale) // 2 * 2)
        poster = target.with_suffix(".jpg")
        try:
            run_binary(
                [
                    "ffmpeg",
                    "-y",
                    "-ss",
                    "0",
                    "-i",
                    str(source),
                    "-frames:v",
                    "1",
                    "-vf",
                    f"scale={output_width}:{output_height}:flags=lanczos",
                    str(poster),
                ],
                capture=True,
            )
            run_binary(
                [
                    "ffmpeg",
                    "-y",
                    "-loop",
                    "1",
                    "-framerate",
                    "1",
                    "-i",
                    str(poster),
                    "-t",
                    f"{duration:.3f}",
                    "-vf",
                    "fps=5,format=yuv420p",
                    "-c:v",
                    "libx264",
                    "-preset",
                    "veryfast",
                    "-tune",
                    "stillimage",
                    "-an",
                    "-movflags",
                    "+faststart",
                    "-metadata",
                    f"comment=pptx-tools:{variant.get('placeholder_token', '')}",
                    str(target),
                ],
                capture=True,
            )
        finally:
            poster.unlink(missing_ok=True)
        if not target.is_file() or target.stat().st_size == 0:
            raise RuntimeError(f"Failed to create placeholder video for {source}")

    def detach_deck(
        self,
        deck_id: str,
        *,
        source_pptx: Path | None = None,
        output_path: Path | None = None,
        placeholder_builder: PlaceholderBuilder | None = None,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> Path:
        deck = self.deck(deck_id)
        source = (
            source_pptx.expanduser().resolve()
            if source_pptx
            else self.deck_source_path(deck)
        )
        if not source.is_file():
            raise FileNotFoundError(source)
        if sha256_file(source) != deck["source_sha256"]:
            raise ValueError(
                "Source PPTX changed; add it as a new deck before detaching videos"
            )
        output = (
            output_path.expanduser().resolve()
            if output_path
            else (self.root / "decks" / "detached" / f"{source.stem}_lightweight.pptx")
        )
        output = _unique_path(output)
        work = Path(tempfile.mkdtemp(prefix="video-detach-", dir=self.root))
        replacements: dict[str, Path] = {}
        relationship_map: dict[str, str] = {}
        replacement_infos = {}
        remove_paths: set[str] = set()
        compact_assets: dict[str, CompactVideoAsset] = {}
        placeholder_hashes: dict[str, str] = {}
        builder = placeholder_builder or self.build_placeholder
        original_data = copy.deepcopy(self.data)
        succeeded = False
        try:
            with ZipFile(source, "r") as archive:
                total = len(deck["assets"])
                for index, item in enumerate(deck["assets"], start=1):
                    _check_cancelled(cancel_callback)
                    family = self.family(item["family_id"])
                    _, original = self.find_variant(item["original_variant_id"])
                    original_path = self.require_variant_path(original)
                    if progress_callback:
                        progress_callback(
                            f"正在生成占位视频 {index}/{total}: {family['name']}"
                        )
                    placeholder_file = work / f"placeholder-{index}.mp4"
                    placeholder_context = {
                        **original,
                        "placeholder_token": (
                            f"{self.data['project_id']}:{deck_id}:{item['part_path']}"
                        ),
                    }
                    builder(original_path, placeholder_file, placeholder_context)
                    placeholder_hashes[item["part_path"]] = sha256_file(
                        placeholder_file
                    )
                    placeholder_part = item["placeholder_part"]
                    replacements[placeholder_part] = placeholder_file
                    relationship_map[item["part_path"]] = placeholder_part
                    remove_paths.add(item["part_path"])
                    replacement_infos[placeholder_part] = archive.getinfo(
                        item["part_path"]
                    )
                    compact_assets[item["part_path"]] = CompactVideoAsset(
                        media_path=item["part_path"],
                        zip_size=archive.getinfo(item["part_path"]).file_size,
                        output_media_path=placeholder_part,
                    )
            patch_output_pptx(
                source,
                output,
                replacements,
                relationship_path_map=relationship_map,
                replacement_infos=replacement_infos,
                remove_paths=remove_paths,
                video_assets=compact_assets,
            )
            with ZipFile(output, "r") as archive:
                if archive.testzip() is not None:
                    raise BadZipFile("Generated lightweight PPTX failed ZIP validation")
                for item in deck["assets"]:
                    if (
                        item["placeholder_part"] not in archive.namelist()
                        or item["part_path"] in archive.namelist()
                    ):
                        raise RuntimeError(
                            "Generated lightweight PPTX has an invalid media mapping"
                        )
            output_scan = scan_embedded_videos(output)
            expected = {item["placeholder_part"] for item in deck["assets"]}
            if not expected.issubset(output_scan):
                raise RuntimeError(
                    "Generated lightweight PPTX lost one or more video relationships"
                )
            for item in deck["assets"]:
                item["placeholder_sha256"] = placeholder_hashes[item["part_path"]]
            deck["detached_outputs"].append(self._output_record(output, "lightweight"))
            self.save()
            self.record(
                "deck_detached", deck_id=deck_id, output=self.encode_path(output)
            )
            succeeded = True
            return output
        finally:
            shutil.rmtree(work, ignore_errors=True)
            if not succeeded:
                self.data = original_data
                output.unlink(missing_ok=True)

    @staticmethod
    def _current_media_part(
        item: dict[str, Any], scanned: dict[str, VideoAsset], archive: ZipFile
    ) -> str:
        if item["placeholder_part"] in scanned:
            return item["placeholder_part"]
        anchors = {
            (occurrence["slide_path"], occurrence["shape_id"])
            for occurrence in item["occurrences"]
        }
        matches = []
        for media_path, asset in scanned.items():
            current = {
                (occurrence.slide_path, occurrence.shape_id)
                for occurrence in asset.occurrences
            }
            if (
                anchors & current
                and item.get("placeholder_sha256")
                and _sha256_zip_member(archive, media_path)
                == item["placeholder_sha256"]
            ):
                matches.append(media_path)
        if len(matches) == 1:
            return matches[0]
        raise RuntimeError(
            f"Cannot uniquely match video placement for {item['part_path']}"
        )

    def restore_deck(
        self,
        deck_id: str,
        detached_pptx: Path,
        *,
        output_path: Path | None = None,
        progress_callback: ProgressCallback | None = None,
        cancel_callback: CancelCallback | None = None,
    ) -> Path:
        deck = self.deck(deck_id)
        detached = detached_pptx.expanduser().resolve()
        if not detached.is_file():
            raise FileNotFoundError(detached)
        output = (
            output_path.expanduser().resolve()
            if output_path
            else (self.root / "decks" / "restored" / f"{detached.stem}_restored.pptx")
        )
        output = _unique_path(output)
        scanned = scan_embedded_videos(detached)
        replacements: dict[str, Path] = {}
        relationship_map: dict[str, str] = {}
        replacement_infos = {}
        remove_paths: set[str] = set()
        compact_assets: dict[str, CompactVideoAsset] = {}
        expected_hashes: dict[str, str] = {}
        original_data = copy.deepcopy(self.data)
        reserved = set(scanned)
        with ZipFile(detached, "r") as archive:
            total = len(deck["assets"])
            for index, item in enumerate(deck["assets"], start=1):
                _check_cancelled(cancel_callback)
                family = self.family(item["family_id"])
                _, variant = self.find_variant(family["active_variant_id"])
                variant_path = self.require_variant_path(variant)
                if progress_callback:
                    progress_callback(f"正在还原视频 {index}/{total}: {family['name']}")
                current_part = self._current_media_part(item, scanned, archive)
                suffix = (
                    variant_path.suffix.lower()
                    or Path(item["part_path"]).suffix.lower()
                )
                original = Path(item["part_path"])
                if suffix == original.suffix.lower():
                    output_part = item["part_path"]
                else:
                    output_part = f"{original.parent.as_posix()}/{original.stem}_{variant['id'][:8]}{suffix}"
                while output_part in reserved and output_part != current_part:
                    output_part = f"{original.parent.as_posix()}/{original.stem}_{uuid.uuid4().hex[:8]}{suffix}"
                reserved.add(output_part)
                replacements[output_part] = variant_path
                relationship_map[current_part] = output_part
                remove_paths.add(current_part)
                replacement_infos[output_part] = archive.getinfo(current_part)
                compact_assets[current_part] = CompactVideoAsset(
                    media_path=current_part,
                    zip_size=archive.getinfo(current_part).file_size,
                    output_media_path=output_part,
                )
                expected_hashes[output_part] = variant["sha256"]
        try:
            patch_output_pptx(
                detached,
                output,
                replacements,
                relationship_path_map=relationship_map,
                replacement_infos=replacement_infos,
                remove_paths=remove_paths,
                video_assets=compact_assets,
            )
            with ZipFile(output, "r") as archive:
                if archive.testzip() is not None:
                    raise BadZipFile("Restored PPTX failed ZIP validation")
                for media_part, digest in expected_hashes.items():
                    if _sha256_zip_member(archive, media_part) != digest:
                        raise RuntimeError(
                            f"Restored video hash mismatch: {media_part}"
                        )
            output_scan = scan_embedded_videos(output)
            if not set(expected_hashes).issubset(output_scan):
                raise RuntimeError("Restored PPTX lost one or more video relationships")
            deck["restored_outputs"].append(self._output_record(output, "restored"))
            self.save()
            self.record(
                "deck_restored", deck_id=deck_id, output=self.encode_path(output)
            )
            return output
        except Exception:
            self.data = original_data
            output.unlink(missing_ok=True)
            raise

    def status(self, variant: dict[str, Any]) -> str:
        path = self.variant_path(variant)
        if not path.is_file():
            return "missing"
        if path.stat().st_size != variant["size_bytes"]:
            return "modified"
        if variant.get("mtime_ns") and path.stat().st_mtime_ns != variant["mtime_ns"]:
            return "modified"
        return "available"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Archive deduplicated PPTX video sources and upgrade PPTX videos"
    )
    sub = parser.add_subparsers(
        dest="action",
        required=True,
        metavar=(
            "{create,add,import-video,set-source,list,doctor,upgrade,"
            "detach,restore,activate,compress,relink}"
        ),
    )
    create = sub.add_parser("create")
    create.add_argument("project", type=Path)
    create.add_argument("--name")
    add = sub.add_parser("add")
    add.add_argument("project", type=Path)
    add.add_argument("pptx", nargs="+", type=Path)
    add.add_argument(
        "--source-quality",
        choices=("1080p", "mp4", "original"),
        default="1080p",
        help="Archive at high-quality 1080p, normalize to compatible MP4, or retain original bytes",
    )
    add.add_argument(
        "--category",
        default="",
        help="Relative folder under the library media directory",
    )
    add.add_argument(
        "--normalize-mp4",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    import_video = sub.add_parser("import-video")
    import_video.add_argument("project", type=Path)
    import_video.add_argument("video", nargs="+", type=Path)
    import_video.add_argument("--family-id")
    import_video.add_argument(
        "--source-quality", choices=("1080p", "mp4", "original"), default="1080p"
    )
    import_video.add_argument("--category", default="")
    set_source = sub.add_parser("set-source")
    set_source.add_argument("project", type=Path)
    set_source.add_argument("variant_id")
    list_parser = sub.add_parser("list")
    list_parser.add_argument("project", type=Path)
    doctor = sub.add_parser("doctor")
    doctor.add_argument("project", type=Path)
    doctor.add_argument(
        "--verify-hashes",
        action="store_true",
        help="Hash every registered video file in addition to fast metadata checks",
    )
    doctor.add_argument(
        "--prune-stale-outputs",
        action="store_true",
        help="Remove only history records whose generated PPTX files no longer exist",
    )
    doctor.add_argument(
        "--report", type=Path, help="Write the JSON report to this path"
    )
    upgrade = sub.add_parser("upgrade")
    upgrade.add_argument("project", type=Path)
    upgrade.add_argument("pptx", type=Path)
    upgrade.add_argument("--output", type=Path)
    upgrade.add_argument(
        "--incompatible-only",
        action="store_true",
        help="Replace only WMV/AVI and other non-MP4 embedded media",
    )
    detach = sub.add_parser("detach")
    detach.add_argument("project", type=Path)
    detach.add_argument("deck_id")
    detach.add_argument("--source", type=Path)
    detach.add_argument("--output", type=Path)
    restore = sub.add_parser("restore")
    restore.add_argument("project", type=Path)
    restore.add_argument("deck_id")
    restore.add_argument("detached", type=Path)
    restore.add_argument("--output", type=Path)
    activate = sub.add_parser("activate")
    activate.add_argument("project", type=Path)
    activate.add_argument("variant_id")
    compress = sub.add_parser("compress")
    compress.add_argument("project", type=Path)
    compress.add_argument("variant_id")
    compress.add_argument(
        "--profile", choices=("high", "balanced", "aggressive"), default="balanced"
    )
    relink = sub.add_parser("relink")
    relink.add_argument("project", type=Path)
    relink.add_argument("roots", nargs="*", type=Path)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.action == "create":
        project = VideoProject.create(args.project, args.name)
        print(project.manifest_path)
        return 0
    project = VideoProject.open(args.project)
    if args.action == "add":
        for path in args.pptx:
            result = project.archive_and_register_pptx(
                path,
                source_quality=args.source_quality,
                category=args.category,
                progress_callback=print,
            )
            print(
                f"{path}: added={result['added']} reused={result['reused']} "
                f"deck={result['deck']['id']}"
            )
        if args.normalize_mp4:
            print(
                "--normalize-mp4 is deprecated; original media is retained and "
                "PowerPoint-compatible MP4 is generated only when upgrading."
            )
    elif args.action == "import-video":
        results = [
            project.import_external_video(
                path,
                source_quality=args.source_quality,
                category=args.category,
                family_id=args.family_id,
            )
            for path in args.video
        ]
        print(json.dumps(results, ensure_ascii=False, indent=2))
    elif args.action == "set-source":
        project.set_source_variant(args.variant_id)
        family, variant = project.find_variant(args.variant_id)
        print(
            json.dumps(
                {
                    "family_id": family["id"],
                    "family_name": family["name"],
                    "source_variant_id": variant["id"],
                    "path": str(project.variant_path(variant)),
                },
                ensure_ascii=False,
            )
        )
    elif args.action == "list":
        print(
            json.dumps(
                {"name": project.data["name"], "families": project.families()},
                ensure_ascii=False,
                indent=2,
            )
        )
    elif args.action == "doctor":
        from pptx_tools.video_library_health import (
            audit_video_project,
            prune_missing_output_records,
        )

        pruned = (
            prune_missing_output_records(project) if args.prune_stale_outputs else 0
        )
        report = audit_video_project(project, verify_hashes=args.verify_hashes)
        report["pruned_stale_output_records"] = pruned
        rendered = json.dumps(report, ensure_ascii=False, indent=2)
        if args.report:
            destination = args.report.expanduser().resolve()
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_text(rendered + "\n", encoding="utf-8")
            print(destination)
        else:
            print(rendered)
        return 0 if report["ok"] else 1
    elif args.action == "upgrade":
        result = project.upgrade_pptx_from_library(
            args.pptx,
            output_path=args.output,
            incompatible_only=args.incompatible_only,
            progress_callback=print,
        )
        print(
            json.dumps(
                {**result, "output_pptx": str(result["output_pptx"] or "")},
                ensure_ascii=False,
            )
        )
    elif args.action == "detach":
        print(
            project.detach_deck(
                args.deck_id,
                source_pptx=args.source,
                output_path=args.output,
                progress_callback=print,
            )
        )
    elif args.action == "restore":
        print(
            project.restore_deck(
                args.deck_id,
                args.detached,
                output_path=args.output,
                progress_callback=print,
            )
        )
    elif args.action == "activate":
        project.activate_variant(args.variant_id)
    elif args.action == "compress":
        variant = project.compress_variant(args.variant_id, args.profile)
        print(project.variant_path(variant))
    elif args.action == "relink":
        print(
            json.dumps(
                project.relink_missing(args.roots, progress_callback=print),
                ensure_ascii=False,
                indent=2,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
