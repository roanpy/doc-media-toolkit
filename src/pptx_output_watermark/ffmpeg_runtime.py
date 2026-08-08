from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

from .process_utils import hidden_console_kwargs, run_process, subprocess_text_kwargs


def _binary_name(name: str) -> str:
    if sys.platform == "win32" and not name.lower().endswith(".exe"):
        return f"{name}.exe"
    return name


def resolve_binary(name: str) -> str | None:
    binary_name = _binary_name(name)
    override = ""
    for env_name in (
        f"PPTX_TOOLS_{name.upper()}",
        f"PPTX_OUTPUT_WATERMARK_{name.upper()}",
    ):
        override = os.environ.get(env_name, "").strip()
        if override:
            break
    candidates: list[Path] = []
    if override:
        candidates.append(Path(override).expanduser())

    bundle_roots = [
        Path(sys.executable).resolve().parent,
        Path(__file__).resolve().parents[2],
    ]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundle_roots.insert(0, Path(meipass))
    for root in bundle_roots:
        candidates.extend(
            [
                root / binary_name,
                root / "assets" / "bin" / binary_name,
                root / "assets" / "tools" / binary_name,
                root / "tools" / binary_name,
            ]
        )

    seen: set[str] = set()
    for candidate in candidates:
        resolved = str(candidate)
        if resolved in seen:
            continue
        seen.add(resolved)
        if candidate.exists():
            return str(candidate)

    return shutil.which(binary_name) or shutil.which(name)


def ensure_binary(name: str) -> str:
    resolved = resolve_binary(name)
    if resolved is None:
        raise FileNotFoundError(f"Required binary not found: {name}")
    return resolved


def run_binary(
    cmd: list[str],
    *,
    capture: bool = False,
) -> subprocess.CompletedProcess[str]:
    resolved_cmd = list(cmd)
    if resolved_cmd and resolved_cmd[0] in {"ffmpeg", "ffprobe"}:
        resolved_cmd[0] = ensure_binary(resolved_cmd[0])
    return run_process(
        resolved_cmd,
        check=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
        **subprocess_text_kwargs(),
        **hidden_console_kwargs(),
    )
