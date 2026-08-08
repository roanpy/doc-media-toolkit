from __future__ import annotations

import json
import os
import tempfile
import time
from pathlib import Path

DEFAULT_COOLDOWN_SECONDS = 30 * 60
_DISABLED_UNTIL: dict[str, float] = {}
_DISABLED_REASON: dict[str, str] = {}


def circuit_path(app_key: str) -> Path:
    normalized = "".join(ch if ch.isalnum() else "_" for ch in app_key.lower())
    return (
        Path(tempfile.gettempdir()) / f"pptx_output_watermark_{normalized}_circuit.json"
    )


def conversion_disabled(app_key: str) -> tuple[bool, str]:
    now = time.time()
    disabled_until = _DISABLED_UNTIL.get(app_key, 0.0)
    if now < disabled_until:
        return True, _DISABLED_REASON.get(app_key, "")
    try:
        payload = json.loads(circuit_path(app_key).read_text(encoding="utf-8"))
    except Exception:
        return False, ""
    disabled_until = float(payload.get("disabled_until") or 0.0)
    if now < disabled_until:
        return True, str(payload.get("reason") or "")
    return False, ""


def disable_conversion_temporarily(
    app_key: str,
    reason: str,
    *,
    cooldown_seconds: int = DEFAULT_COOLDOWN_SECONDS,
) -> None:
    disabled_until = time.time() + max(60, int(cooldown_seconds))
    _DISABLED_UNTIL[app_key] = disabled_until
    _DISABLED_REASON[app_key] = reason
    payload = {
        "disabled_until": disabled_until,
        "reason": reason,
        "pid": os.getpid(),
        "created_at": time.time(),
    }
    try:
        circuit_path(app_key).write_text(
            json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass
