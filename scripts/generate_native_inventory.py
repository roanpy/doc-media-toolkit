#!/usr/bin/env python3
"""Create a deterministic inventory of native files in a release directory.

The inventory is evidence, not a license verdict.  It records relative paths,
hashes, file descriptions, architectures, and dependency-tool output without
leaking the build machine's absolute paths.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

NATIVE_SUFFIXES = {".bin", ".dylib", ".dll", ".exe", ".pyd", ".so"}
ARCHITECTURE_PATTERNS = (
    ("arm64", re.compile(r"\barm64\b", re.IGNORECASE)),
    ("aarch64", re.compile(r"\baarch64\b", re.IGNORECASE)),
    ("x86_64", re.compile(r"\b(?:x86[-_ ]64|amd64)\b", re.IGNORECASE)),
    ("i386", re.compile(r"\b(?:i386|i686|80386)\b", re.IGNORECASE)),
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _run_text(command: list[str], *, timeout: int = 30) -> tuple[str, str | None]:
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return "", str(exc)
    output = (result.stdout or result.stderr).strip()
    if result.returncode != 0:
        return output, f"exit {result.returncode}"
    return output, None


def _file_description(path: Path) -> tuple[str, str | None]:
    executable = shutil.which("file")
    if executable is None:
        return "", "file command is unavailable"
    return _run_text([executable, "-b", str(path)])


def _architectures(description: str) -> list[str]:
    return [
        name for name, pattern in ARCHITECTURE_PATTERNS if pattern.search(description)
    ]


def _otool_dependencies(path: Path) -> tuple[list[str], str | None]:
    if sys.platform != "darwin" or shutil.which("otool") is None:
        return [], None
    description, _error = _file_description(path)
    if "Mach-O" not in description:
        return [], None
    output, error = _run_text(["otool", "-L", str(path)])
    if error:
        return [], error
    dependencies = []
    for line in output.splitlines()[1:]:
        value = line.strip().split(" (")[0]
        if value:
            dependencies.append(value)
    return dependencies, None


def _windows_dependencies(path: Path) -> tuple[list[str], str | None, str | None]:
    if os.name != "nt":
        return [], None, None
    command: list[str] | None = None
    tool_name: str | None = None
    dumpbin = shutil.which("dumpbin")
    if dumpbin:
        command, tool_name = [dumpbin, "/DEPENDENTS", str(path)], "dumpbin"
    else:
        objdump = shutil.which("llvm-objdump") or shutil.which("objdump")
        if objdump:
            command, tool_name = [objdump, "-p", str(path)], Path(objdump).name
    if command is None:
        return [], None, None
    output, error = _run_text(command)
    if error:
        return [], error, tool_name
    dependencies = []
    for line in output.splitlines():
        match = re.search(r"(?:DLL Name:\s*)?([A-Za-z0-9_.-]+\.dll)\s*$", line)
        if match:
            dependencies.append(match.group(1))
    return sorted(set(dependencies)), None, tool_name


def _is_native(path: Path) -> bool:
    if path.suffix.lower() in NATIVE_SUFFIXES:
        return True
    return path.is_file() and os.access(path, os.X_OK) and not path.suffix


def build_inventory(
    root: Path,
    *,
    platform_name: str | None = None,
    architecture: str | None = None,
    artifact: Path | None = None,
) -> dict[str, Any]:
    root = root.expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"inventory root is not a directory: {root}")
    artifact_path = artifact.expanduser().resolve() if artifact is not None else None
    if artifact_path is not None and not artifact_path.is_file():
        raise ValueError(f"artifact is not a file: {artifact_path}")

    entries: list[dict[str, Any]] = []
    external_symlinks: list[str] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or not _is_native(path):
            continue
        if path.is_symlink():
            try:
                path.resolve().relative_to(root)
            except ValueError:
                external_symlinks.append(path.relative_to(root).as_posix())
                continue
        relative = path.relative_to(root).as_posix()
        description, description_error = _file_description(path)
        dependencies, dependency_error = _otool_dependencies(path)
        windows_deps, windows_error, dependency_tool = _windows_dependencies(path)
        if windows_deps:
            dependencies = windows_deps
        entry: dict[str, Any] = {
            "path": relative,
            "size_bytes": path.stat().st_size,
            "sha256": _sha256(path),
            "file_description": description,
            "architectures": _architectures(description),
            "dependencies": dependencies,
        }
        errors = [
            error
            for error in (description_error, dependency_error, windows_error)
            if error
        ]
        if dependency_tool:
            entry["dependency_tool"] = dependency_tool
        if errors:
            entry["inspection_warnings"] = errors
        entries.append(entry)

    inventory = {
        "schema": "doc-media-toolkit.native-inventory.v1",
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "platform": platform_name or sys.platform,
        "architecture": architecture or platform.machine(),
        "artifact_root": root.name,
        "entries": entries,
        "external_symlinks": sorted(external_symlinks),
    }
    if artifact_path is not None:
        inventory["artifact"] = artifact_path.name
        inventory["artifact_sha256"] = _sha256(artifact_path)
    return inventory


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inventory native files in an unpacked app or portable release directory."
    )
    parser.add_argument(
        "root", type=Path, help="Unpacked app or portable release directory"
    )
    parser.add_argument("--output", type=Path, required=True, help="JSON report path")
    parser.add_argument("--platform", default=None, help="Release platform label")
    parser.add_argument(
        "--architecture", default=None, help="Release architecture label"
    )
    parser.add_argument(
        "--artifact",
        type=Path,
        default=None,
        help="Packaged artifact file whose SHA-256 binds this inventory",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        inventory = build_inventory(
            args.root,
            platform_name=args.platform,
            architecture=args.architecture,
            artifact=args.artifact,
        )
    except (OSError, ValueError) as exc:
        print(f"native inventory failed: {exc}", file=sys.stderr)
        return 1
    if not inventory["entries"]:
        print("native inventory failed: no native files found", file=sys.stderr)
        return 1
    if inventory["external_symlinks"]:
        print("native inventory failed: external symlinks found", file=sys.stderr)
        return 1
    output = args.output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(inventory, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"native inventory: {len(inventory['entries'])} file(s) -> {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
