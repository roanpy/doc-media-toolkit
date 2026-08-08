from __future__ import annotations

import os
import sys
from pathlib import Path


def _candidate_plugin_roots() -> list[Path]:
    executable_dir = Path(sys.executable).resolve().parent
    bundle_root = Path(getattr(sys, "_MEIPASS", executable_dir)).resolve()
    return [
        bundle_root / "PySide6" / "Qt" / "plugins",
        executable_dir / "_internal" / "PySide6" / "Qt" / "plugins",
        executable_dir / "PySide6" / "Qt" / "plugins",
    ]


if sys.platform == "win32":
    for plugin_root in _candidate_plugin_roots():
        platform_root = plugin_root / "platforms"
        if (platform_root / "qwindows.dll").exists():
            os.environ.setdefault("QT_PLUGIN_PATH", str(plugin_root))
            os.environ.setdefault("QT_QPA_PLATFORM_PLUGIN_PATH", str(platform_root))
            break
