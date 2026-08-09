#!/usr/bin/env python3
"""Run an installed malware scanner against one release artifact.

The command fails closed when no supported scanner is available.  It never
labels an artifact clean based on codesigning, Gatekeeper, or a missing scan.
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

SCHEMA = "doc-media-toolkit.malware-scan.v1"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _tree_sha256(root: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(item for item in root.rglob("*") if item.is_file()):
        relative = path.relative_to(root).as_posix().encode("utf-8")
        digest.update(len(relative).to_bytes(8, "big"))
        digest.update(relative)
        digest.update(bytes.fromhex(_sha256(path)))
    return digest.hexdigest()


def _find_defender() -> str | None:
    direct = shutil.which("MpCmdRun.exe")
    if direct:
        return direct
    program_data = os.environ.get("ProgramData")
    if not program_data:
        return None
    root = Path(program_data) / "Microsoft" / "Windows Defender" / "Platform"
    if not root.is_dir():
        return None
    candidates = sorted(root.glob("*/MpCmdRun.exe"), reverse=True)
    return str(candidates[0]) if candidates else None


def find_scanner(explicit: str | None = None) -> tuple[str | None, str | None]:
    if explicit:
        candidate = shutil.which(explicit) or explicit
        return (candidate if Path(candidate).is_file() else None), "explicit"
    clamscan = shutil.which("clamscan")
    if clamscan:
        return clamscan, "clamscan"
    defender = _find_defender()
    if defender:
        return defender, "windows-defender"
    return None, None


def _command(scanner: str, kind: str, artifact: Path) -> list[str]:
    if kind == "clamscan":
        if artifact.is_dir():
            return [scanner, "--no-summary", "--recursive", str(artifact)]
        return [scanner, "--no-summary", str(artifact)]
    if kind == "windows-defender":
        return [
            scanner,
            "-Scan",
            "-ScanType",
            "3",
            "-File",
            str(artifact),
            "-DisableRemediation",
        ]
    return [scanner, str(artifact)]


def _run(command: list[str]) -> tuple[int | None, str, str]:
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=60 * 60,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return None, "", str(exc)
    return result.returncode, result.stdout[-4000:], result.stderr[-4000:]


def _redact_output(text: str, artifact: Path, *aliases: Path) -> str:
    candidates: dict[str, str] = {}
    for path in (artifact, *aliases):
        candidates.update(
            {
                str(path): "<artifact>",
                str(path.absolute()): "<artifact>",
                str(path.parent): "<artifact-dir>",
                str(path.parent.absolute()): "<artifact-dir>",
                str(path.parent.resolve()): "<artifact-dir>",
            }
        )
    # Windows runners may stringify the same temporary directory through a
    # short (RUNNER~1) or long user path depending on whether it was resolved.
    # Redact both separator styles and path casing so scanner output cannot
    # leak the host temporary directory.
    normalized: dict[str, str] = {}
    for value, replacement in candidates.items():
        for variant in {value, value.replace("\\", "/"), value.replace("/", "\\")}:
            normalized[variant] = replacement
    candidates = normalized
    for value, replacement in list(candidates.items()):
        if value.startswith("/private/"):
            candidates[value.removeprefix("/private")] = replacement
    for value, replacement in sorted(
        candidates.items(), key=lambda item: -len(item[0])
    ):
        text = text.replace(value, replacement)
        if os.name == "nt":
            text = re.sub(re.escape(value), replacement, text, flags=re.IGNORECASE)
    return text


def scan_artifact(
    artifact: Path,
    *,
    scanner: str | None = None,
) -> dict[str, Any]:
    artifact_argument = artifact.expanduser()
    artifact_absolute = artifact_argument.absolute()
    artifact = artifact_absolute.resolve()
    if not artifact.is_file() and not artifact.is_dir():
        raise ValueError(f"artifact does not exist: {artifact}")
    scanner_path, scanner_kind = find_scanner(scanner)
    digest = _sha256(artifact) if artifact.is_file() else None
    tree_digest = _tree_sha256(artifact) if artifact.is_dir() else None
    report: dict[str, Any] = {
        "schema": SCHEMA,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "platform": sys.platform,
        "machine": platform.machine(),
        "artifact": artifact.name,
        "artifact_sha256": digest,
        "artifact_tree_sha256": tree_digest,
        "scanner": Path(scanner_path).name if scanner_path else None,
        "status": "unavailable",
        "exit_code": None,
        "stdout_tail": "",
        "stderr_tail": "",
    }
    if scanner_path is None or scanner_kind is None:
        report["reason"] = (
            "No supported malware scanner found; clean status is not asserted."
        )
        return report
    command = _command(scanner_path, scanner_kind, artifact)
    exit_code, stdout, stderr = _run(command)
    report["exit_code"] = exit_code
    report["stdout_tail"] = _redact_output(
        stdout, artifact, artifact_argument, artifact_absolute
    )
    report["stderr_tail"] = _redact_output(
        stderr, artifact, artifact_argument, artifact_absolute
    )
    if exit_code == 0:
        report["status"] = "clean"
    elif exit_code == 1 and scanner_kind == "clamscan":
        report["status"] = "infected"
    else:
        report["status"] = "error"
    return report


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run ClamAV or Windows Defender and write fail-closed JSON evidence."
    )
    parser.add_argument(
        "artifact", type=Path, help="DMG, ZIP, EXE, or unpacked app directory"
    )
    parser.add_argument("--output", type=Path, required=True, help="JSON report path")
    parser.add_argument("--scanner", default=None, help="Explicit scanner executable")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        report = scan_artifact(args.artifact, scanner=args.scanner)
    except (OSError, ValueError) as exc:
        print(f"malware scan failed: {exc}", file=sys.stderr)
        return 1
    output = args.output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"malware scan: {report['status']} -> {output}")
    return 0 if report["status"] == "clean" else 1


if __name__ == "__main__":
    raise SystemExit(main())
