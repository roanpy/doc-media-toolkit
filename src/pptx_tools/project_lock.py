from __future__ import annotations

import json
import os
import shutil
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


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


@contextmanager
def project_write_lock(
    root: Path,
    lock_name: str,
    timeout_seconds: float = 5.0,
) -> Iterator[None]:
    lock = root / lock_name
    token = uuid.uuid4().hex
    deadline = time.monotonic() + timeout_seconds
    while True:
        try:
            lock.mkdir()
            try:
                (lock / "owner.json").write_text(
                    json.dumps({"pid": os.getpid(), "token": token}), encoding="utf-8"
                )
            except Exception:
                shutil.rmtree(lock, ignore_errors=True)
                raise
            break
        except FileExistsError:
            try:
                owner = json.loads((lock / "owner.json").read_text(encoding="utf-8"))
                if not _pid_alive(int(owner.get("pid", 0))):
                    shutil.rmtree(lock, ignore_errors=True)
                    continue
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                try:
                    if time.time() - lock.stat().st_mtime > 30:
                        shutil.rmtree(lock, ignore_errors=True)
                        continue
                except OSError:
                    pass
            if time.monotonic() >= deadline:
                raise TimeoutError("Media library is being modified by another process")
            time.sleep(0.05)
    try:
        yield
    finally:
        try:
            owner = json.loads((lock / "owner.json").read_text(encoding="utf-8"))
            if owner.get("token") == token:
                shutil.rmtree(lock, ignore_errors=True)
        except (OSError, json.JSONDecodeError):
            pass
