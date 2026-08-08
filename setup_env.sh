#!/usr/bin/env bash
set -euo pipefail

PYTHON_BIN="${PYTHON_BIN:-python3.12}"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "Python 3.12 is required. Install it or set PYTHON_BIN to its executable." >&2
  exit 1
fi

if ! "$PYTHON_BIN" -c 'import sys; raise SystemExit(sys.version_info[:2] != (3, 12))'; then
  echo "Release environments must use Python 3.12; found $("$PYTHON_BIN" --version 2>&1)." >&2
  exit 1
fi

"$PYTHON_BIN" -m venv --clear .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -e ".[build,dev]"

.venv/bin/python - <<'PY'
import shutil
print("ffmpeg:", shutil.which("ffmpeg") or "missing")
print("ffprobe:", shutil.which("ffprobe") or "missing")
PY
