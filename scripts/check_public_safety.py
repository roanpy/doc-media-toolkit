#!/usr/bin/env python3
from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path
from typing import Sequence

LOCAL_OR_SECRET_PATTERNS = tuple(
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
        r"\bAKIA[0-9A-Z]{16}\b",
        r"\bAIza[0-9A-Za-z_-]{35}\b",
        r"\bhf_[A-Za-z0-9]{20,}\b",
        r"\bgh[pousr]_[A-Za-z0-9]{20,}\b",
        r"\bsk-(?:proj-)?[A-Za-z0-9_-]{20,}\b",
    )
)
PRIVATE_DENYLIST_NAME = ".public-safety-denylist.local"
SELF_FILE = Path(__file__).resolve()
MAX_TEXT_BYTES = 5 * 1024 * 1024

SENSITIVE_NAMES = {
    ".git-credentials",
    ".netrc",
    ".npmrc",
    ".pypirc",
    ".zshrc",
    "application_default_credentials.json",
    "credentials.json",
    "id_ed25519",
    "id_ecdsa",
    "id_rsa",
    "service-account.json",
    "service_account.json",
}
SENSITIVE_PARTS = {
    ".aws",
    ".azure",
    ".cache",
    ".claude",
    ".codex",
    ".gcloud",
    ".gemini",
    ".gnupg",
    ".hermes",
    ".huggingface",
    ".lmstudio",
    ".ollama",
    ".pi",
    ".ssh",
    ".zcode",
}
SENSITIVE_SUFFIXES = {
    ".ckpt",
    ".db",
    ".gguf",
    ".key",
    ".keystore",
    ".log",
    ".jks",
    ".p12",
    ".pem",
    ".pfx",
    ".pt",
    ".pth",
    ".safetensors",
    ".sqlite",
    ".sqlite3",
    ".token",
}

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
            [
                "git",
                "ls-files",
                "-z",
                "--cached",
                "--others",
                "--exclude-standard",
            ],
            cwd=root,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
        )
        return [root / os.fsdecode(item) for item in result.stdout.split(b"\0") if item]
    except (subprocess.CalledProcessError, FileNotFoundError):
        ignored_dirs = {".git", ".venv", "build", "dist", "__pycache__"}
        return [
            path
            for path in root.rglob("*")
            if (path.is_file() or path.is_symlink())
            and not any(part in ignored_dirs for part in path.relative_to(root).parts)
        ]


def is_text_file(path: Path) -> bool:
    if path.suffix.lower() in TEXT_SUFFIXES or path.name in {"LICENSE", "README"}:
        return True
    try:
        with path.open("rb") as handle:
            prefix = handle.read(8192)
        prefix.decode("utf-8")
    except (OSError, UnicodeDecodeError):
        return False
    return b"\0" not in prefix


def check_sensitive_path(root: Path, path: Path) -> list[str]:
    relative = path.relative_to(root)
    name = path.name.lower()
    sensitive = (
        name == ".env"
        or (name.startswith(".env.") and name != ".env.example")
        or name in SENSITIVE_NAMES
        or path.suffix.lower() in SENSITIVE_SUFFIXES
        or any(part.lower() in SENSITIVE_PARTS for part in relative.parts[:-1])
    )
    return (
        [f"{relative.as_posix()}: sensitive file name or directory"]
        if sensitive
        else []
    )


def check_symbolic_link(root: Path, path: Path) -> list[str]:
    return (
        [f"{path.relative_to(root).as_posix()}: symbolic links are not allowed"]
        if path.is_symlink()
        else []
    )


def private_denylist_patterns(root: Path) -> tuple[re.Pattern[str], ...]:
    path = root / PRIVATE_DENYLIST_NAME
    if not path.is_file():
        return ()
    phrases = (
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    )
    return tuple(re.compile(re.escape(phrase), re.IGNORECASE) for phrase in phrases)


def check_file_content(
    root: Path,
    path: Path,
    patterns: Sequence[re.Pattern[str]] = LOCAL_OR_SECRET_PATTERNS,
) -> list[str]:
    if path.is_symlink() or not path.is_file() or path.resolve() == SELF_FILE:
        return []
    if not is_text_file(path):
        return []
    try:
        if path.stat().st_size > MAX_TEXT_BYTES:
            return [
                f"{path.relative_to(root).as_posix()}: text file exceeds safety scan limit"
            ]
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return [
            f"{path.relative_to(root).as_posix()}: declared text file is not valid UTF-8"
        ]
    except OSError:
        return [f"{path.relative_to(root).as_posix()}: file could not be read safely"]

    findings: list[str] = []
    rel = path.relative_to(root)
    for pattern in patterns:
        if pattern.search(text):
            findings.append(f"{rel}: matched sensitive/local pattern")
    return findings


def check_hard_links(root: Path, path: Path) -> list[str]:
    if path.is_symlink():
        return []
    try:
        stat = path.stat()
    except OSError:
        return []
    if not path.is_file() or stat.st_nlink <= 1:
        return []
    return [f"{path.relative_to(root).as_posix()}: hard-link count is {stat.st_nlink}"]


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    patterns = (*LOCAL_OR_SECRET_PATTERNS, *private_denylist_patterns(root))
    findings: list[str] = []
    for path in public_files(root):
        findings.extend(check_sensitive_path(root, path))
        findings.extend(check_symbolic_link(root, path))
        findings.extend(check_file_content(root, path, patterns))
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
