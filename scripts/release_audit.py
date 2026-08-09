#!/usr/bin/env python3
"""Local pre-release dependency and artifact audit.

A standalone checklist runner for the moments before tagging/packaging. It does
NOT replace the project's test suite; it complements it by verifying the supply
chain and artifact provenance that tests do not cover:

1. Lockfile integrity: ``uv lock --check`` and a locked sync dry-run.
2. Dependency vulnerability scan via ``pip-audit`` (external tool; if absent it
   is reported as a skipped step, never added as a runtime dependency).
3. Optional CycloneDX SBOM via ``uv export`` (only when explicitly requested).
4. Git branch/commit provenance and a clean working tree.
5. Packaged binary version + SHA-256 hashes for everything under ``dist/``.
6. Optional fail-closed public-binary evidence validation.

This is a target-host entry point. It can run locally or inside the manually
dispatched candidate-build workflow; it never publishes a release.

Usage::

    python scripts/release_audit.py --check
    python scripts/release_audit.py --check --with-sbom dist/sbom.json
    python scripts/release_audit.py --check --dist-dir dist --output release-audit.md
    python scripts/release_audit.py --check --public-binary \
      --with-sbom release-assets/sbom.cdx.json \
      --evidence release-assets/public-binary-evidence.json \
      --dist-dir release-assets
    python scripts/release_audit.py --help
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import platform
import re
import shutil
import subprocess
import sys
import tarfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from subprocess import TimeoutExpired as SubprocessTimeoutExpired
from typing import Any

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_SRC = _PROJECT_ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

# Imported after sys.path mutation so the project version is resolvable.
from pptx_tools import __version__ as PROJECT_VERSION  # noqa: E402


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _run(
    cmd: list[str],
    *,
    cwd: Path = _PROJECT_ROOT,
    timeout: int = 300,
) -> dict[str, Any]:
    """Run a command, capturing output without raising on non-zero exit."""
    try:
        completed = subprocess.run(  # noqa: S603 - command is constructed internally
            cmd,
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        return {
            "cmd": cmd,
            "returncode": completed.returncode,
            "stdout": completed.stdout,
            "stderr": completed.stderr,
        }
    except FileNotFoundError:
        return {
            "cmd": cmd,
            "returncode": 127,
            "stdout": "",
            "stderr": f"{cmd[0]} not found",
        }
    except SubprocessTimeoutExpired:
        return {
            "cmd": cmd,
            "returncode": 124,
            "stdout": "",
            "stderr": f"timed out after {timeout}s",
        }


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _tool_available(name: str) -> bool:
    return shutil.which(name) is not None


def check_uv_lock() -> dict[str, Any]:
    """Verify the lockfile is in sync with pyproject.toml."""
    uv = shutil.which("uv")
    if uv is None:
        return {
            "name": "uv lock --check",
            "status": "fail",
            "detail": "uv not installed; install uv or run on a host that has it.",
        }
    lock_check = _run([uv, "lock", "--check"])
    sync_check = _run([uv, "sync", "--locked", "--dry-run"])
    ok = lock_check["returncode"] == 0 and sync_check["returncode"] == 0
    return {
        "name": "uv lock --check / locked sync",
        "status": "pass" if ok else "fail",
        "detail": (
            f"uv lock --check exit={lock_check['returncode']}; "
            f"uv sync --locked --dry-run exit={sync_check['returncode']}"
        ),
        "uv_lock_stderr": lock_check["stderr"].strip()[:2000] or None,
        "uv_sync_stderr": sync_check["stderr"].strip()[:2000] or None,
    }


def check_pip_audit() -> dict[str, Any]:
    """Run pip-audit if present; never install it as a dependency."""
    executable = shutil.which("pip-audit")
    if executable is not None:
        command = [executable, "--local", "--desc"]
    elif importlib.util.find_spec("pip_audit") is not None:
        command = [sys.executable, "-m", "pip_audit", "--local", "--desc"]
    else:
        return {
            "name": "pip-audit",
            "status": "skipped",
            "detail": (
                "pip-audit not found in PATH. It is an external, opt-in tool: "
                "install it in an isolated env (e.g. `uvx pip-audit`) to run this step."
            ),
        }
    result = _run(command, timeout=600)
    # pip-audit returns non-zero when vulnerabilities are found.
    vulns = result["stdout"]
    status = "pass" if result["returncode"] == 0 else "vulns_found"
    return {
        "name": "pip-audit",
        "status": status,
        "detail": f"exit={result['returncode']}",
        "output": vulns.strip()[:4000] or None,
    }


def check_git_state() -> dict[str, Any]:
    """Require a reproducible commit for a release audit."""
    git = shutil.which("git")
    if git is None:
        return {
            "name": "git provenance",
            "status": "fail",
            "detail": "git not installed; release provenance cannot be recorded.",
        }
    commit = _run([git, "rev-parse", "HEAD"])
    branch = _run([git, "branch", "--show-current"])
    status = _run([git, "status", "--porcelain=v1", "--untracked-files=all"])
    if commit["returncode"] != 0 or branch["returncode"] != 0:
        return {
            "name": "git provenance",
            "status": "fail",
            "detail": "unable to resolve the current Git commit/branch.",
        }
    commit_id = commit["stdout"].strip()
    branch_name = branch["stdout"].strip() or "(detached HEAD)"
    dirty_files = [line for line in status["stdout"].splitlines() if line][:100]
    dirty = bool(dirty_files) or status["returncode"] != 0
    return {
        "name": "git provenance",
        "status": "fail" if dirty else "pass",
        "detail": (
            f"branch={branch_name}; commit={commit_id[:12]}; "
            f"working_tree={'dirty' if dirty else 'clean'}"
        ),
        "branch": branch_name,
        "commit": commit_id,
        "dirty": dirty,
        "dirty_files": dirty_files,
    }


def check_sbom(output_path: Path | None) -> dict[str, Any]:
    """Export a CycloneDX SBOM from the locked dependency graph when requested."""
    uv = shutil.which("uv")
    if uv is None:
        return {
            "name": "CycloneDX SBOM",
            "status": "skipped",
            "detail": (
                "uv not found in PATH. SBOM generation is optional; install uv "
                "or run this check on a host that has it."
            ),
        }
    if output_path is None:
        return {
            "name": "CycloneDX SBOM",
            "status": "skipped",
            "detail": "not requested (pass --with-sbom PATH to generate it).",
        }
    target = output_path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    result = _run(
        [
            uv,
            "export",
            "--locked",
            "--format",
            "cyclonedx1.5",
            "--output-file",
            str(target),
        ]
    )
    ok = result["returncode"] == 0 and target.is_file()
    return {
        "name": "CycloneDX SBOM",
        "status": "pass" if ok else "fail",
        "detail": f"exit={result['returncode']}; target={target}",
        "sbom_path": str(target) if target.is_file() else None,
        "stderr": result["stderr"].strip()[:2000] or None,
    }


def check_dist_artifacts(dist_dir: Path) -> dict[str, Any]:
    """Hash every file under dist/ and record the project version."""
    files: list[dict[str, Any]] = []
    if dist_dir.is_dir():
        for path in sorted(dist_dir.rglob("*")):
            if path.is_file():
                stat = path.stat()
                files.append(
                    {
                        "path": str(path.relative_to(dist_dir)),
                        "sha256": _sha256(path),
                        "size_bytes": stat.st_size,
                    }
                )
    return {
        "name": "packaged artifacts",
        "status": "pass" if files else "fail",
        "detail": f"{len(files)} file(s) hashed under {dist_dir}",
        "project_version": PROJECT_VERSION,
        "dist_dir": str(dist_dir),
        "files": files,
    }


def check_python_runtime() -> dict[str, Any]:
    """Record the Python version so cross-platform mismatches are visible."""
    return {
        "name": "python runtime",
        "status": "pass",
        "detail": sys.version.split()[0],
        "platform": platform.platform(),
        "machine": platform.machine(),
    }


def _evidence_file(
    root: Path,
    item: Any,
    *,
    label: str,
) -> tuple[Path | None, str | None]:
    if not isinstance(item, dict):
        return None, f"{label} must be an object with path and sha256."
    relative = item.get("path")
    expected = item.get("sha256")
    if not isinstance(relative, str) or not relative.strip():
        return None, f"{label}.path is required."
    if not isinstance(expected, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", expected):
        return None, f"{label}.sha256 must be a 64-character hexadecimal digest."
    candidate = (root / relative).resolve()
    try:
        candidate.relative_to(root)
    except ValueError:
        return None, f"{label}.path escapes the release directory."
    if not candidate.is_file():
        return None, f"{label} is missing: {relative}"
    actual = _sha256(candidate)
    if actual.lower() != expected.lower():
        return None, f"{label} SHA-256 mismatch: {relative}"
    return candidate, None


def _read_json_sidecar(
    root: Path,
    item: Any,
    *,
    label: str,
) -> tuple[Path | None, dict[str, Any] | None, str | None]:
    path, error = _evidence_file(root, item, label=label)
    if error or path is None:
        return path, None, error
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        return path, None, f"{label} is not valid JSON: {exc}"
    if not isinstance(value, dict):
        return path, None, f"{label} must contain a JSON object."
    return path, value, None


def _validate_nonempty_sidecar(
    root: Path,
    item: Any,
    *,
    label: str,
) -> list[str]:
    path, error = _evidence_file(root, item, label=label)
    if error:
        return [error]
    if path is None or path.stat().st_size == 0:
        return [f"{label} must not be empty"]
    return []


def _validate_sbom(
    root: Path,
    item: Any,
    *,
    label: str,
    artifact_sha256: str | None,
) -> list[str]:
    _path, document, error = _read_json_sidecar(root, item, label=label)
    if error:
        return [error]
    assert document is not None
    findings: list[str] = []
    if document.get("bomFormat") != "CycloneDX":
        findings.append(f"{label}.bomFormat must be CycloneDX")
    spec_version = document.get("specVersion")
    if not isinstance(spec_version, str) or not spec_version.startswith("1."):
        findings.append(f"{label}.specVersion must be a CycloneDX 1.x version")
    components = document.get("components")
    if not isinstance(components, list) or not components:
        findings.append(f"{label}.components must contain at least one component")
    if artifact_sha256 is not None and item.get("artifact_sha256") != artifact_sha256:
        findings.append(f"{label}.artifact_sha256 does not match the packaged artifact")
    return findings


def _validate_native_inventory(
    root: Path,
    item: Any,
    *,
    label: str,
    artifact_sha256: str | None,
) -> list[str]:
    _path, document, error = _read_json_sidecar(root, item, label=label)
    if error:
        return [error]
    assert document is not None
    findings: list[str] = []
    if document.get("schema") != "doc-media-toolkit.native-inventory.v1":
        findings.append(f"{label}.schema is unsupported")
    entries = document.get("entries")
    if not isinstance(entries, list) or not entries:
        findings.append(f"{label}.entries must contain at least one native file")
        return findings
    if artifact_sha256 is not None and item.get("artifact_sha256") != artifact_sha256:
        findings.append(f"{label}.artifact_sha256 does not match the packaged artifact")
    for index, entry in enumerate(entries):
        entry_label = f"{label}.entries[{index}]"
        if not isinstance(entry, dict):
            findings.append(f"{entry_label} must be an object")
            continue
        relative = entry.get("path")
        if not isinstance(relative, str) or not relative.strip():
            findings.append(f"{entry_label}.path is required")
        else:
            path = PurePosixPath(relative)
            if path.is_absolute() or ".." in path.parts:
                findings.append(f"{entry_label}.path must stay within the artifact")
        digest = entry.get("sha256")
        if not isinstance(digest, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", digest):
            findings.append(
                f"{entry_label}.sha256 must be a 64-character hexadecimal digest"
            )
    return findings


def _validate_malware_report(
    root: Path,
    item: Any,
    *,
    label: str,
    artifact_sha256: str | None,
) -> list[str]:
    _path, document, error = _read_json_sidecar(root, item, label=label)
    if error:
        return [error]
    assert document is not None
    findings: list[str] = []
    if document.get("schema") != "doc-media-toolkit.malware-scan.v1":
        findings.append(f"{label}.schema is unsupported")
    if document.get("status") != "clean":
        findings.append(f"{label}.status must be clean")
    if document.get("exit_code") != 0:
        findings.append(f"{label}.exit_code must be 0")
    scanner = document.get("scanner")
    if not isinstance(scanner, str) or not scanner.strip():
        findings.append(f"{label}.scanner is required")
    if (
        artifact_sha256 is not None
        and document.get("artifact_sha256") != artifact_sha256
    ):
        findings.append(f"{label}.artifact_sha256 does not match the packaged artifact")
    return findings


def _validate_corresponding_source(path: Path, *, label: str) -> list[str]:
    required = {
        "build_ffmpeg_runtime.sh",
        "changes.diff",
        "SHA256SUMS",
        "BUILD-INFO.txt",
        "sources/ffmpeg-8.1.2.tar.xz",
        "sources/zlib-1.3.2.tar.xz",
    }
    required_nonempty = required - {"changes.diff"}
    try:
        with tarfile.open(path, mode="r:*") as archive:
            unsafe: list[str] = []
            non_regular: list[str] = []
            regular_sizes: dict[str, list[int]] = {}
            for member in archive.getmembers():
                name = member.name.rstrip("/")
                if not name:
                    continue
                member_path = PurePosixPath(name)
                if member_path.is_absolute() or ".." in member_path.parts:
                    unsafe.append(name)
                if not member.isdir() and not member.isreg():
                    non_regular.append(name)
                if member.isreg():
                    relative_name = name.split("/", 1)[1] if "/" in name else name
                    regular_sizes.setdefault(relative_name, []).append(member.size)
    except (OSError, tarfile.TarError) as exc:
        return [f"{label} is not a readable tar archive: {exc}"]
    findings = [f"{label} contains unsafe archive member paths"] if unsafe else []
    if non_regular:
        findings.append(f"{label} contains non-regular archive members")
    missing = sorted(required - regular_sizes.keys())
    if missing:
        findings.append(f"{label} is missing required files: {', '.join(missing)}")
    empty = sorted(
        name
        for name, sizes in regular_sizes.items()
        if name in required_nonempty and not any(size > 0 for size in sizes)
    )
    if empty:
        findings.append(f"{label} contains empty required files: {', '.join(empty)}")
    if not any(
        name.startswith("sources/x264-") and any(size > 0 for size in sizes)
        for name, sizes in regular_sizes.items()
    ):
        findings.append(f"{label} is missing the pinned x264 source archive")
    return findings


def check_public_binary_evidence(
    evidence_path: Path | None,
    dist_dir: Path,
) -> dict[str, Any]:
    """Fail closed unless every public binary has verifiable sidecar evidence."""
    name = "public binary evidence"
    if evidence_path is None:
        return {
            "name": name,
            "status": "fail",
            "detail": "--evidence is required with --public-binary.",
        }
    try:
        evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        return {
            "name": name,
            "status": "fail",
            "detail": f"unable to read evidence: {exc}",
        }

    findings: list[str] = []
    root = dist_dir.expanduser().resolve()
    if evidence.get("schema") != "doc-media-toolkit.public-binary-evidence.v1":
        findings.append("unsupported or missing evidence schema")
    if evidence.get("version") != PROJECT_VERSION:
        findings.append("evidence version does not match the application version")

    artifacts = evidence.get("artifacts")
    if not isinstance(artifacts, list) or not artifacts:
        findings.append("at least one artifact is required")
        artifacts = []

    for index, artifact in enumerate(artifacts):
        prefix = f"artifacts[{index}]"
        if not isinstance(artifact, dict):
            findings.append(f"{prefix} must be an object")
            continue
        artifact_path, error = _evidence_file(root, artifact, label=prefix)
        if error:
            findings.append(error)
        artifact_sha256 = _sha256(artifact_path) if artifact_path is not None else None
        platform_name = artifact.get("platform")
        package_type = artifact.get("package_type")
        if platform_name not in {"macos", "windows"}:
            findings.append(f"{prefix}.platform must be macos or windows")
        allowed_package_types = {
            "macos": {"dmg", "app-zip"},
            "windows": {"portable-zip", "installer"},
        }
        if package_type not in allowed_package_types.get(platform_name, set()):
            findings.append(f"{prefix}.package_type is not allowed for {platform_name}")
        architecture = artifact.get("architecture")
        if not isinstance(architecture, str) or not architecture.strip():
            findings.append(f"{prefix}.architecture is required")
        if platform_name == "windows" and package_type == "onefile":
            findings.append(
                f"{prefix} uses blocked Windows onefile distribution; use a replaceable onedir package"
            )

        signature = artifact.get("signature")
        if not isinstance(signature, dict) or signature.get("status") != "valid":
            findings.append(f"{prefix} lacks a valid platform signature")
        else:
            expected_signature = (
                "developer-id" if platform_name == "macos" else "authenticode"
            )
            if signature.get("type") != expected_signature:
                findings.append(f"{prefix}.signature.type must be {expected_signature}")
            findings.extend(
                _validate_nonempty_sidecar(
                    root,
                    signature.get("report"),
                    label=f"{prefix}.signature.report",
                )
            )

        if platform_name == "macos":
            notarization = artifact.get("notarization")
            if (
                not isinstance(notarization, dict)
                or notarization.get("status") != "valid"
            ):
                findings.append(f"{prefix} lacks valid Apple notarization evidence")
            else:
                findings.extend(
                    _validate_nonempty_sidecar(
                        root,
                        notarization.get("report"),
                        label=f"{prefix}.notarization.report",
                    )
                )

        malware = artifact.get("malware_scan")
        if not isinstance(malware, dict) or malware.get("status") != "clean":
            findings.append(f"{prefix} lacks a clean malware-scan result")
        else:
            findings.extend(
                _validate_malware_report(
                    root,
                    malware.get("report"),
                    label=f"{prefix}.malware_scan.report",
                    artifact_sha256=artifact_sha256,
                )
            )

        findings.extend(
            _validate_sbom(
                root,
                artifact.get("sbom"),
                label=f"{prefix}.sbom",
                artifact_sha256=artifact_sha256,
            )
        )
        findings.extend(
            _validate_native_inventory(
                root,
                artifact.get("native_inventory"),
                label=f"{prefix}.native_inventory",
                artifact_sha256=artifact_sha256,
            )
        )

        ffmpeg = artifact.get("ffmpeg")
        if not isinstance(ffmpeg, dict) or not isinstance(ffmpeg.get("bundled"), bool):
            findings.append(f"{prefix}.ffmpeg.bundled must be true or false")
        elif ffmpeg["bundled"]:
            if ffmpeg.get("version") != "8.1.2":
                findings.append(f"{prefix}.ffmpeg.version must be 8.1.2")
            configuration = ffmpeg.get("configuration")
            if not isinstance(configuration, str) or not {
                "--enable-gpl",
                "--enable-version3",
                "--enable-libx264",
            }.issubset(configuration.split()):
                findings.append(f"{prefix}.ffmpeg.configuration is incomplete")
            if ffmpeg.get("license") != "GPL-3.0-or-later":
                findings.append(f"{prefix}.ffmpeg.license must be GPL-3.0-or-later")
            source, error = _evidence_file(
                root,
                ffmpeg.get("corresponding_source"),
                label=f"{prefix}.ffmpeg.corresponding_source",
            )
            if error:
                findings.append(error)
            elif source is not None:
                findings.extend(
                    _validate_corresponding_source(
                        source, label=f"{prefix}.ffmpeg.corresponding_source"
                    )
                )
        elif (
            not isinstance(ffmpeg.get("runtime_requirement"), str)
            or not ffmpeg["runtime_requirement"].strip()
        ):
            findings.append(
                f"{prefix}.ffmpeg.runtime_requirement is required when FFmpeg is external"
            )

    return {
        "name": name,
        "status": "fail" if findings else "pass",
        "detail": (
            f"{len(findings)} blocking finding(s)"
            if findings
            else f"{len(artifacts)} artifact(s) have complete evidence"
        ),
        "findings": findings,
    }


def run_all_checks(args_ns: argparse.Namespace) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    checks.append(check_python_runtime())
    checks.append(check_git_state())
    checks.append(check_uv_lock())
    if args_ns.public_binary and args_ns.skip_pip_audit:
        checks.append(
            {
                "name": "pip-audit",
                "status": "fail",
                "detail": "--skip-pip-audit is not allowed with --public-binary.",
            }
        )
    elif not args_ns.skip_pip_audit:
        pip_audit = check_pip_audit()
        if args_ns.public_binary and pip_audit["status"] == "skipped":
            pip_audit["status"] = "fail"
            pip_audit["detail"] = "pip-audit is required for a public binary release."
        checks.append(pip_audit)
    if args_ns.with_sbom is not None and not args_ns.skip_sbom:
        checks.append(check_sbom(args_ns.with_sbom))
    elif args_ns.public_binary:
        checks.append(
            {
                "name": "CycloneDX SBOM",
                "status": "fail",
                "detail": "--with-sbom is required with --public-binary.",
            }
        )
    checks.append(check_dist_artifacts(args_ns.dist_dir))
    if args_ns.public_binary:
        checks.append(check_public_binary_evidence(args_ns.evidence, args_ns.dist_dir))

    statuses = {c["status"] for c in checks}
    overall = "pass" if not (statuses & {"fail", "vulns_found"}) else "fail"
    return {
        "schema": "pptx-tools.release-audit",
        "generated_at": _utc_now_iso(),
        "overall_status": overall,
        "project_version": PROJECT_VERSION,
        "git": next(
            (check for check in checks if check["name"] == "git provenance"),
            {},
        ),
        "checks": checks,
    }


def write_markdown(report: dict[str, Any], path: Path) -> None:
    lines = [
        "# 发布前审计报告",
        "",
        f"- 生成时间：`{report['generated_at']}`",
        f"- 项目版本：`{report['project_version']}`",
        f"- 总体状态：**{report['overall_status']}**",
        "",
        "## 检查项",
        "",
        "| 检查 | 状态 | 说明 |",
        "| --- | --- | --- |",
    ]
    for check in report["checks"]:
        lines.append(f"| {check['name']} | `{check['status']}` | {check['detail']} |")
    lines.extend(
        [
            "",
            "## Git 溯源",
            "",
            f"- 分支：`{report.get('git', {}).get('branch', '-')}`",
            f"- 提交：`{report.get('git', {}).get('commit', '-')}`",
            f"- 工作树：`{'dirty' if report.get('git', {}).get('dirty') else 'clean'}`",
            "",
            "",
            "## 平台与签名边界",
            "",
            "- 跨平台产物必须在对应平台构建；PyInstaller 不支持交叉编译。",
            "- macOS 产物默认 ad-hoc 签名；Developer ID 公证需单独配置 `--notary-profile`，"
            "本审计不执行公证。",
            "- Windows 产物签名需在 Windows 主机用证书完成；本审计只记录哈希，不签名。",
            "- 普通候选中 pip-audit 为可选外部工具；公开二进制模式必须执行。CycloneDX 由现有 uv 导出，不纳入运行时依赖。",
        ]
    )
    artifact_check = next(
        (c for c in report["checks"] if c["name"] == "packaged artifacts"), None
    )
    if artifact_check and artifact_check.get("files"):
        lines.extend(
            [
                "",
                "## 产物哈希",
                "",
                "| 文件 | 大小(bytes) | SHA-256 |",
                "| --- | ---: | --- |",
            ]
        )
        for item in artifact_check["files"]:
            lines.append(
                f"| {item['path']} | {item['size_bytes']} | `{item['sha256']}` |"
            )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Local pre-release dependency and artifact audit. Runs uv lock check, "
            "Git provenance, optional pip-audit, optional CycloneDX SBOM, and dist/ "
            "hash recording. "
            "Runs locally or in the manual candidate workflow; never publishes."
        )
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Run the audit checks (default action).",
    )
    parser.add_argument(
        "--dist-dir",
        type=Path,
        default=_PROJECT_ROOT / "dist",
        help="Directory to scan for packaged artifacts (default: ./dist).",
    )
    parser.add_argument(
        "--with-sbom",
        type=Path,
        default=None,
        help="Generate a CycloneDX SBOM at this path with uv export.",
    )
    parser.add_argument(
        "--skip-pip-audit",
        action="store_true",
        help="Skip the pip-audit step even if the tool is installed.",
    )
    parser.add_argument(
        "--skip-sbom",
        action="store_true",
        help="Skip CycloneDX SBOM generation entirely.",
    )
    parser.add_argument(
        "--public-binary",
        action="store_true",
        help=(
            "Apply fail-closed public binary gates. Requires a generated SBOM, "
            "pip-audit, and --evidence with signatures, notarization where applicable, "
            "malware scan, native inventory, and FFmpeg source-delivery evidence."
        ),
    )
    parser.add_argument(
        "--evidence",
        type=Path,
        default=None,
        help="Public-binary evidence JSON inside the release directory.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=_PROJECT_ROOT / "release-audit.md",
        help="Markdown report output path.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args_ns = parse_args(argv)
    # --check is the default and only mode; kept for CLI ergonomics.
    report = run_all_checks(args_ns)
    args_ns.output.parent.mkdir(parents=True, exist_ok=True)
    write_markdown(report, args_ns.output)
    json_path = args_ns.output.with_suffix(".json")
    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"release audit: overall={report['overall_status']}")
    for check in report["checks"]:
        print(f"  - {check['name']}: {check['status']} ({check['detail']})")
    print(f"Wrote {args_ns.output}")
    print(f"Wrote {json_path}")
    return 0 if report["overall_status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
