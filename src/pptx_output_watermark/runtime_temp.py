from __future__ import annotations

import json
import os
import shutil
import tempfile
import time
from pathlib import Path

OWNER_FILE_NAME = ".ppt_tools_runtime_owner.json"
DEFAULT_MAX_AGE_HOURS = 24
RUNTIME_PATTERNS = (
    "pptx_output_watermark_preview_source_*",
    "pptx_output_watermark_preview_overlay_*",
    "pptx_output_watermark_rendered_*",
    "pptx_output_watermark_images_*",
    "pptx_output_watermark_preprocessed_*",
    "pptx_output_watermark_video_*",
    "pptx_output_watermark_overlay_*",
    "pptx_output_watermark_pdf_*",
    "pptx_output_watermark_lo_profile_*",
)


def _runtime_root() -> Path:
    return Path(tempfile.gettempdir())


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def register_runtime_dir(path: Path, *, purpose: str) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    owner_file = path / OWNER_FILE_NAME
    payload = {
        "pid": os.getpid(),
        "created_at": time.time(),
        "purpose": purpose,
    }
    try:
        owner_file.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
    except Exception:
        pass
    return path


def create_runtime_temp_dir(prefix: str, *, purpose: str) -> Path:
    return register_runtime_dir(
        Path(tempfile.mkdtemp(prefix=prefix)),
        purpose=purpose,
    )


def cleanup_stale_runtime_entries(
    *,
    patterns: tuple[str, ...] = RUNTIME_PATTERNS,
    max_age_hours: int = DEFAULT_MAX_AGE_HOURS,
) -> None:
    cutoff = time.time() - max(1, max_age_hours) * 3600
    temp_root = _runtime_root()
    for pattern in patterns:
        for path in temp_root.glob(pattern):
            try:
                stat = path.stat()
            except OSError:
                continue
            if stat.st_mtime >= cutoff:
                continue
            if path.is_dir():
                owner_file = path / OWNER_FILE_NAME
                if owner_file.exists():
                    try:
                        payload = json.loads(owner_file.read_text(encoding="utf-8"))
                    except Exception:
                        payload = {}
                    pid = int(payload.get("pid") or 0)
                    if pid and _pid_alive(pid):
                        continue
                shutil.rmtree(path, ignore_errors=True)
            else:
                try:
                    path.unlink()
                except OSError:
                    pass
