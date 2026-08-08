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

No GitHub Actions: this is a local entry point meant to be run by hand on the
build host, on each target platform, before producing a release artifact.

Usage::

    python scripts/release_audit.py --check
    python scripts/release_audit.py --check --with-sbom dist/sbom.json
    python scripts/release_audit.py --check --dist-dir dist --output release-audit.md
    python scripts/release_audit.py --help
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import platform
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
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


def run_all_checks(args_ns: argparse.Namespace) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    checks.append(check_python_runtime())
    checks.append(check_git_state())
    checks.append(check_uv_lock())
    if not args_ns.skip_pip_audit:
        checks.append(check_pip_audit())
    if args_ns.with_sbom is not None and not args_ns.skip_sbom:
        checks.append(check_sbom(args_ns.with_sbom))
    checks.append(check_dist_artifacts(args_ns.dist_dir))

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
            "- pip-audit 为可选外部工具；CycloneDX 由现有 uv 导出，不纳入运行时依赖。",
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
            "No GitHub Actions; run by hand on each build platform."
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
