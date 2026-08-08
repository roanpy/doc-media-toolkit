from __future__ import annotations

import locale
import os
import signal
import subprocess
import sys
import threading
import time


_ACTIVE_PROCESSES: dict[subprocess.Popen, int] = {}
_ACTIVE_PROCESSES_LOCK = threading.Lock()


def start_process(args, **kwargs) -> subprocess.Popen:
    if os.name != "nt":
        kwargs.setdefault("start_new_session", True)
    process = subprocess.Popen(args, **kwargs)
    with _ACTIVE_PROCESSES_LOCK:
        _ACTIVE_PROCESSES[process] = threading.get_ident()
    return process


def finish_process(process: subprocess.Popen) -> None:
    with _ACTIVE_PROCESSES_LOCK:
        _ACTIVE_PROCESSES.pop(process, None)


def kill_process(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        else:
            os.killpg(process.pid, signal.SIGKILL)
    except (OSError, ProcessLookupError):
        pass


def terminate_process(process: subprocess.Popen, grace_seconds: float = 1.0) -> None:
    if process.poll() is not None:
        return
    if os.name == "nt":
        kill_process(process)
        process.wait()
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except (OSError, ProcessLookupError):
        pass
    try:
        process.wait(timeout=grace_seconds)
    except subprocess.TimeoutExpired:
        kill_process(process)
        process.wait()


def run_process(
    args,
    *,
    input=None,
    capture_output: bool = False,
    timeout: float | None = None,
    check: bool = False,
    **kwargs,
) -> subprocess.CompletedProcess:
    if capture_output:
        kwargs["stdout"] = subprocess.PIPE
        kwargs["stderr"] = subprocess.PIPE
    process = start_process(args, **kwargs)
    try:
        try:
            stdout, stderr = process.communicate(input, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            kill_process(process)
            stdout, stderr = process.communicate()
            raise subprocess.TimeoutExpired(
                exc.cmd, exc.timeout, output=stdout, stderr=stderr
            ) from None
    finally:
        finish_process(process)
    completed = subprocess.CompletedProcess(args, process.returncode, stdout, stderr)
    if check:
        completed.check_returncode()
    return completed


def terminate_active_processes(
    grace_seconds: float = 1.0,
    *,
    owner_thread_id: int | None = None,
) -> None:
    with _ACTIVE_PROCESSES_LOCK:
        processes = [
            process
            for process, thread_id in _ACTIVE_PROCESSES.items()
            if owner_thread_id is None or thread_id == owner_thread_id
        ]
    for process in processes:
        if process.poll() is not None:
            continue
        if os.name == "nt":
            kill_process(process)
            continue
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except (OSError, ProcessLookupError):
            pass
    deadline = time.monotonic() + grace_seconds
    for process in processes:
        if process.poll() is not None:
            continue
        try:
            process.wait(timeout=max(0.0, deadline - time.monotonic()))
        except subprocess.TimeoutExpired:
            kill_process(process)


def hidden_console_kwargs() -> dict[str, object]:
    if sys.platform != "win32":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startupinfo,
    }


def subprocess_text_encoding() -> str:
    encoding = locale.getpreferredencoding(False) or "utf-8"
    return encoding if encoding.lower() != "ascii" else "utf-8"


def subprocess_text_kwargs() -> dict[str, object]:
    return {
        "text": True,
        "encoding": subprocess_text_encoding(),
        "errors": "replace",
    }
