#!/usr/bin/env python3
from __future__ import annotations

import re
import subprocess
from pathlib import Path

LOCAL_OR_SECRET_PATTERNS = [
    re.compile(pattern, re.IGNORECASE)
    for pattern in (
        r"/Users/[^\s'\"`]+",
        r"\\Users\\[^\s'\"`]+",
        r"[A-Z]:\\Users\\[^\s'\"`]+",
        r"/private/var/[^\s'\"`]+",
        r"\bDownloads\b",
        r"BEGIN (RSA |EC |OPENSSH |)PRIVATE KEY",
        r"GH_TOKEN\s*=",
        r"GITHUB_TOKEN\s*=",
        r"APPLE_ID_PASSWORD\s*=",
        r"WINDOWS_CERT_PASSWORD\s*=",
        r"AWS_SECRET_ACCESS_KEY\s*=",
    )
]
SELF_FILE = Path(__file__).resolve()

TEXT_SUFFIXES = {
    ".cfg",
    ".css",
    ".html",
    ".ini",
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".sh",
    ".svg",
    ".toml",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}


def public_files(root: Path) -> list[Path]:
    try:
        result = subprocess.run(
            ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
            cwd=root,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
        )
        return [root / line for line in result.stdout.splitlines() if line.strip()]
    except (subprocess.CalledProcessError, FileNotFoundError):
        ignored_dirs = {".git", ".venv", "build", "dist", "__pycache__"}
        return [
            path
            for path in root.rglob("*")
            if path.is_file()
            and not any(part in ignored_dirs for part in path.relative_to(root).parts)
        ]


def is_text_file(path: Path) -> bool:
    return path.suffix.lower() in TEXT_SUFFIXES or path.name in {"LICENSE", "README"}


def check_file_content(root: Path, path: Path) -> list[str]:
    if not path.is_file() or path.resolve() == SELF_FILE:
        return []
    if not is_text_file(path):
        return []
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return []

    findings: list[str] = []
    rel = path.relative_to(root)
    for pattern in LOCAL_OR_SECRET_PATTERNS:
        for match in pattern.finditer(text):
            findings.append(
                f"{rel}: matched sensitive/local pattern `{match.group(0)}`"
            )
    return findings


def check_hard_links(root: Path, path: Path) -> list[str]:
    try:
        stat = path.stat()
    except OSError:
        return []
    if not path.is_file() or stat.st_nlink <= 1:
        return []
    return [f"{path.relative_to(root)}: hard-link count is {stat.st_nlink}"]


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    findings: list[str] = []
    for path in public_files(root):
        findings.extend(check_file_content(root, path))
        findings.extend(check_hard_links(root, path))

    if findings:
        print("Public safety check failed:")
        for item in findings:
            print(f"- {item}")
        return 1

    print("Public safety check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
