"""LibreOffice process isolation helpers."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

from .process_utils import run_process, subprocess_text_kwargs
from .runtime_temp import create_runtime_temp_dir

SOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
LIBREOFFICE_CONVERSION_COOLDOWN_SECONDS = 30 * 60

_CONVERSION_DISABLED_UNTIL = 0.0
_CONVERSION_DISABLED_REASON = ""


def _safe_which(command: str) -> str | None:
    try:
        return shutil.which(command)
    except AttributeError:
        return None


def _bundled_soffice_candidates() -> list[str]:
    """Return complete LibreOffice runtimes bundled beside a frozen app."""

    roots: list[Path] = []
    frozen_root = getattr(sys, "_MEIPASS", "")
    if frozen_root:
        roots.append(Path(frozen_root))
    executable = Path(sys.executable).resolve()
    roots.extend([executable.parent, executable.parent / "_internal"])
    if sys.platform == "darwin":
        for parent in executable.parents:
            if parent.name == "Contents":
                roots.extend([parent / "Resources", parent / "Frameworks"])
                break

    relative_paths = (
        (Path("libreoffice") / "LibreOffice.app" / "Contents" / "MacOS" / "soffice",)
        if sys.platform == "darwin"
        else (
            Path("libreoffice") / "LibreOffice" / "program" / "soffice.com",
            Path("libreoffice") / "LibreOffice" / "program" / "soffice.exe",
        )
    )
    return [str(root / relative) for root in roots for relative in relative_paths]


def _soffice_candidates() -> list[str]:
    candidates: list[str] = []
    for env_name in ("PPTX_TOOLS_SOFFICE", "PPTX_OUTPUT_WATERMARK_SOFFICE"):
        override = os.environ.get(env_name, "").strip()
        if override:
            candidates.append(override)
            break

    candidates.extend(_bundled_soffice_candidates())

    if sys.platform == "darwin":
        candidates.append(SOFFICE_PATH)
    elif sys.platform == "win32":
        for env_name in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            root = os.environ.get(env_name, "").strip()
            if not root:
                continue
            candidates.extend(
                [
                    str(Path(root) / "LibreOffice" / "program" / "soffice.com"),
                    str(Path(root) / "LibreOffice" / "program" / "soffice.exe"),
                    str(
                        Path(root)
                        / "Programs"
                        / "LibreOffice"
                        / "program"
                        / "soffice.com"
                    ),
                    str(
                        Path(root)
                        / "Programs"
                        / "LibreOffice"
                        / "program"
                        / "soffice.exe"
                    ),
                ]
            )

    candidates.extend(
        filter(
            None,
            [
                _safe_which("soffice"),
                _safe_which("soffice.com"),
                _safe_which("soffice.exe"),
                _safe_which("libreoffice"),
            ],
        )
    )
    if sys.platform != "darwin":
        candidates.append(SOFFICE_PATH)
    return candidates


def resolve_soffice_path() -> str:
    candidates = _soffice_candidates()
    seen: set[str] = set()
    for candidate in candidates:
        resolved = str(Path(candidate).expanduser())
        if resolved in seen:
            continue
        seen.add(resolved)
        if os.path.exists(resolved):
            return resolved
    return candidates[0] if candidates else SOFFICE_PATH


def conversion_circuit_path() -> str:
    return os.path.join(
        tempfile.gettempdir(), "pptx_output_watermark_libreoffice_circuit.json"
    )


def conversion_disabled() -> tuple[bool, str]:
    if time.time() < _CONVERSION_DISABLED_UNTIL:
        return True, _CONVERSION_DISABLED_REASON
    try:
        with open(conversion_circuit_path(), "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        disabled_until = float(payload.get("disabled_until") or 0)
        if time.time() < disabled_until:
            return True, str(payload.get("reason") or "LibreOffice conversion disabled")
    except Exception:
        pass
    return False, ""


def disable_conversion_temporarily(reason: str) -> None:
    global _CONVERSION_DISABLED_REASON
    global _CONVERSION_DISABLED_UNTIL
    _CONVERSION_DISABLED_REASON = reason
    _CONVERSION_DISABLED_UNTIL = time.time() + LIBREOFFICE_CONVERSION_COOLDOWN_SECONDS
    try:
        with open(conversion_circuit_path(), "w", encoding="utf-8") as handle:
            json.dump(
                {
                    "disabled_until": _CONVERSION_DISABLED_UNTIL,
                    "reason": reason,
                    "pid": os.getpid(),
                    "created_at": time.time(),
                },
                handle,
                ensure_ascii=False,
            )
    except Exception:
        pass


def subprocess_env(profile_dir: str) -> dict[str, str]:
    if sys.platform == "win32":
        # Instead of trying to clean a dirty environment (which is prone to
        # edge cases and case-sensitivity issues), we construct a pristine
        # environment from scratch. This guarantees LibreOffice won't inherit
        # PyInstaller's sys._MEIPASS path, PySide6's QT_PLUGIN_PATH, or
        # anything else that could cause DLL hijacking (0xC0000142).
        env = {}
        # Core Windows variables needed for basic execution
        for k in (
            "SystemRoot",
            "windir",
            "TEMP",
            "TMP",
            "USERPROFILE",
            "APPDATA",
            "LOCALAPPDATA",
        ):
            if k in os.environ:
                env[k] = os.environ[k]

        # Construct a minimal, clean PATH
        sys_root = env.get("SystemRoot", r"C:\Windows")
        env["PATH"] = os.pathsep.join(
            [
                sys_root,
                os.path.join(sys_root, "System32"),
                os.path.join(sys_root, "System32", "Wbem"),
                os.path.join(sys_root, "System32", "WindowsPowerShell", "v1.0"),
            ]
        )
    else:
        env = os.environ.copy()
        env.setdefault("SAL_USE_VCLPLUGIN", "svp")

    env.setdefault("JAVA_TOOL_OPTIONS", "-Djava.awt.headless=true")
    env["UserInstallation"] = Path(profile_dir).as_uri()
    return env


def run_convert_command(
    soffice_path: str,
    args: list[str],
    *,
    profile_dir: str,
    timeout_seconds: int,
) -> subprocess.CompletedProcess:
    direct_cmd = [
        soffice_path,
        f"-env:UserInstallation={Path(profile_dir).as_uri()}",
        *args,
    ]
    if sys.platform == "darwin" and os.path.exists(soffice_path):
        app_bundle = next(
            (
                parent
                for parent in Path(soffice_path).resolve().parents
                if parent.suffix == ".app"
            ),
            Path(soffice_path),
        )
        launch_cmd = [
            "open",
            "-W",
            "-n",
            "-g",
            "-a",
            str(app_bundle),
            "--args",
            f"-env:UserInstallation=file://{profile_dir}",
            *args,
        ]
        return run_process(
            launch_cmd,
            capture_output=True,
            timeout=max(10, int(timeout_seconds)),
            env=subprocess_env(profile_dir),
            **subprocess_text_kwargs(),
        )

    bat_dir = (
        Path(profile_dir) if profile_dir else Path(tempfile.mkdtemp(prefix="lo_bat_"))
    )
    bat_path = bat_dir / "run_lo.bat"

    # Force minimal pristine environment inside the BAT script
    sys_root = os.environ.get("SystemRoot", r"C:\Windows")
    bat_lines = [
        "@echo off",
        f"set PATH={sys_root}\\System32;{sys_root}\\System32\\Wbem;{sys_root}\\System32\\WindowsPowerShell\\v1.0",
        "set QT_PLUGIN_PATH=",
        "set QT_QPA_PLATFORM_PLUGIN_PATH=",
    ]
    if profile_dir:
        bat_lines.append(f"set UserInstallation=file:///{Path(profile_dir).as_posix()}")

    # Quote arguments properly
    cmd_str = f'"{direct_cmd[0]}" ' + " ".join(
        f'"{a}"' if " " in str(a) or "&" in str(a) else str(a) for a in direct_cmd[1:]
    )
    bat_lines.append(cmd_str)

    bat_encoding = "mbcs" if sys.platform == "win32" else "utf-8"
    bat_path.write_text("\n".join(bat_lines), encoding=bat_encoding)

    try:
        return run_process(
            [str(bat_path)],
            capture_output=True,
            timeout=max(10, int(timeout_seconds)),
            # Use default environment, the BAT file will scrub it
            env=None,
            **subprocess_text_kwargs(),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    finally:
        if not profile_dir:
            shutil.rmtree(bat_dir, ignore_errors=True)


def convert_file_to_dir(
    input_path: str | Path,
    output_dir: str | Path,
    convert_to: str,
    *,
    timeout_seconds: int,
    soffice_path: str | None = None,
    lock_timeout_seconds: int | None = None,
) -> subprocess.CompletedProcess | None:
    resolved_soffice = soffice_path or resolve_soffice_path()
    if not os.path.exists(resolved_soffice):
        return None
    args = [
        "--headless",
        "--invisible",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        convert_to,
        "--outdir",
        str(output_dir),
        str(input_path),
    ]
    profile_dir = str(
        create_runtime_temp_dir(
            "pptx_output_watermark_lo_profile_",
            purpose="libreoffice_conversion_profile",
        )
    )
    try:
        with conversion_lock(
            timeout_seconds=lock_timeout_seconds
            if lock_timeout_seconds is not None
            else max(30, int(timeout_seconds) + 30)
        ):
            return run_convert_command(
                resolved_soffice,
                args,
                profile_dir=profile_dir,
                timeout_seconds=timeout_seconds,
            )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


@contextmanager
def conversion_lock(timeout_seconds: int = 600):
    lock_path = os.path.join(
        tempfile.gettempdir(), "pptx_output_watermark_libreoffice.lock"
    )
    handle = open(lock_path, "a+", encoding="utf-8")
    try:
        try:
            import fcntl
        except ImportError:
            yield
            return
        deadline = time.time() + max(1, int(timeout_seconds))
        while True:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                handle.seek(0)
                handle.truncate()
                handle.write(f"pid={os.getpid()} started_at={time.time()}\n")
                handle.flush()
                break
            except BlockingIOError:
                if time.time() >= deadline:
                    raise TimeoutError(
                        "Timed out waiting for LibreOffice conversion lock"
                    )
                time.sleep(0.5)
        try:
            yield
        finally:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
    finally:
        handle.close()
