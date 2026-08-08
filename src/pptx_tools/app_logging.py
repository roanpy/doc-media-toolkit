from __future__ import annotations

import logging
import json
import os
import sys
import tempfile
import threading
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from urllib.parse import urlparse


APP_DIR_NAME = "Doc Media Toolkit"
AI_AUDIT_MAX_BYTES = 5 * 1024 * 1024
_AI_AUDIT_LOCK = threading.Lock()
LOGGER_NAMES = (
    "pptx_tools",
    "pptx_output_watermark",
    "pptx_video_compactor",
    "pptx_video_compactor_gui",
    "pptx_quality_audit",
)


def app_dir_name() -> str:
    experimental = os.environ.get("PPTX_TOOLS_EXPERIMENTAL", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    } or (
        bool(getattr(sys, "frozen", False))
        and "experimental" in Path(sys.executable).stem.lower()
    )
    return f"{APP_DIR_NAME} Experimental" if experimental else APP_DIR_NAME


def log_directory() -> Path:
    if sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    elif os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        root = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state"))
    return root / app_dir_name() / "logs"


def configure_app_logging() -> Path:
    directory = log_directory()
    try:
        directory.mkdir(parents=True, exist_ok=True)
        log_path = directory / "app.log"
    except OSError:
        directory = Path(tempfile.gettempdir()) / app_dir_name() / "logs"
        try:
            directory.mkdir(parents=True, exist_ok=True)
            log_path = directory / "app.log"
        except OSError:
            log_path = Path(tempfile.gettempdir()) / "doc_media_toolkit_app.log"

    root = logging.getLogger(LOGGER_NAMES[0])
    handler = next(
        (
            item
            for item in root.handlers
            if isinstance(item, (RotatingFileHandler, logging.StreamHandler))
        ),
        None,
    )
    if handler is None:
        try:
            handler = RotatingFileHandler(
                log_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
            )
        except OSError:
            try:
                fallback_path = (
                    Path(tempfile.gettempdir()) / f"doc_media_toolkit_{os.getpid()}.log"
                )
                handler = RotatingFileHandler(
                    fallback_path,
                    maxBytes=1 * 1024 * 1024,
                    backupCount=1,
                    encoding="utf-8",
                )
                log_path = fallback_path
            except OSError:
                handler = logging.StreamHandler(sys.stderr)

        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
        )

    for name in LOGGER_NAMES:
        logger = logging.getLogger(name)
        if handler not in logger.handlers:
            logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return log_path


def write_ai_audit_event(
    *,
    media_kind: str,
    target_id: str,
    provider: str,
    model: str,
    vision_enabled: bool,
    applied_fields: list[str],
    merge_group_count: int,
) -> None:
    record = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "media_kind": media_kind,
        "target_id": target_id,
        "provider": urlparse(provider).netloc or "custom",
        "model": model,
        "vision_enabled": vision_enabled,
        "applied_fields": applied_fields,
        "merge_group_count": merge_group_count,
    }
    try:
        directory = log_directory()
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / "ai-audit.jsonl"
        with _AI_AUDIT_LOCK:
            if path.is_file() and path.stat().st_size >= AI_AUDIT_MAX_BYTES:
                previous = directory / "ai-audit.jsonl.1"
                previous.unlink(missing_ok=True)
                os.replace(path, previous)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    except OSError:
        logging.getLogger("pptx_tools").warning("Unable to write AI audit event")
