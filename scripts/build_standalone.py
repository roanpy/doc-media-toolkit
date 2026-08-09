#!/usr/bin/env python3
from __future__ import annotations

import argparse
import ast
import os
import plistlib
import re
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Iterable
from importlib.metadata import PackageNotFoundError, distribution
from pathlib import Path

from packaging.markers import default_environment
from packaging.requirements import Requirement
from packaging.utils import canonicalize_name


BUNDLE_IDENTIFIER = "com.roanpy.doc-media-toolkit"
EXPERIMENTAL_BUNDLE_IDENTIFIER = "com.roanpy.doc-media-toolkit.experimental"
RELEASE_PYTHON = (3, 12)
MINIMUM_MACOS_VERSION = "13.0"
CLI_HIDDEN_IMPORTS = (
    "pptx_output_watermark.cli",
    "pptx_video_compactor",
    "pptx_tools.video_manager",
    "pptx_tools.image_manager",
    # Lazy-loaded by pptx_video_compactor._compact_document_backend; PyInstaller
    # static analysis cannot see importlib.import_module targets.
    "docx_image_compactor",
    "pdf_image_compactor",
    "xlsx_image_compactor",
    "pikepdf",
)
QT_LICENSE_DISTRIBUTIONS = {
    "pyside6",
    "pyside6-addons",
    "pyside6-essentials",
    "shiboken6",
}


def require_release_python(version_info: tuple[int, int] | None = None) -> None:
    current = version_info or sys.version_info[:2]
    if current != RELEASE_PYTHON:
        expected = ".".join(map(str, RELEASE_PYTHON))
        actual = ".".join(map(str, current))
        raise SystemExit(
            f"Standalone releases require Python {expected}; found Python {actual}. "
            "Run ./setup_env.sh and build with .venv/bin/python."
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build standalone Doc Media Toolkit bundles."
    )
    parser.add_argument("--name", help="Output application/executable name.")
    parser.add_argument(
        "--experimental",
        action="store_true",
        help="Build an isolated Experimental app under dist/experimental/<branch>/<commit>.",
    )
    parser.add_argument(
        "--gui", action="store_true", help="Build the tabbed desktop GUI."
    )
    parser.add_argument("--cli", action="store_true", help="Build the unified CLI.")
    parser.add_argument(
        "--onefile", action="store_true", help="Build a single executable."
    )
    parser.add_argument(
        "--windows-onefile",
        action="store_true",
        help="Shortcut for Windows GUI one-file build. Must run on Windows.",
    )
    parser.add_argument(
        "--target-platform",
        choices=["auto", "macos", "windows", "linux"],
        default="auto",
        help="Expected build host platform. PyInstaller does not cross-compile.",
    )
    parser.add_argument("--clean", action="store_true", help="Clean PyInstaller cache.")
    parser.add_argument("--icon", type=Path, help="Application icon path.")
    parser.add_argument(
        "--dmg", action="store_true", help="Build a macOS DMG for the GUI app."
    )
    parser.add_argument(
        "--dmg-output", type=Path, help="Override macOS DMG output path."
    )
    parser.add_argument(
        "--bundle-ffmpeg",
        action="store_true",
        help="Bundle ffmpeg and ffprobe when available.",
    )
    parser.add_argument(
        "--require-ffmpeg-bundle",
        action="store_true",
        help="Fail if --bundle-ffmpeg cannot locate ffmpeg or ffprobe.",
    )
    parser.add_argument(
        "--bundle-libreoffice",
        action="store_true",
        help="Bundle a complete LibreOffice runtime for offline PDF export.",
    )
    parser.add_argument(
        "--require-libreoffice-bundle",
        action="store_true",
        help="Fail if a complete LibreOffice runtime cannot be located.",
    )
    parser.add_argument(
        "--libreoffice-root",
        type=Path,
        help="LibreOffice.app on macOS or the LibreOffice install directory on Windows.",
    )
    parser.add_argument(
        "--max-bundled-binary-mb",
        type=int,
        default=260,
        help="Fail when bundled ffmpeg/ffprobe exceeds this size. Use 0 to disable.",
    )
    parser.add_argument(
        "--codesign-identity",
        default=os.environ.get("PPTX_TOOLS_CODESIGN_IDENTITY", "-"),
        help="macOS signing identity. Defaults to ad-hoc signing.",
    )
    parser.add_argument(
        "--notary-profile",
        default=os.environ.get("PPTX_TOOLS_NOTARY_PROFILE", ""),
        help="Optional keychain profile created by xcrun notarytool store-credentials.",
    )
    return parser.parse_args()


def host_platform() -> str:
    return {"darwin": "macos", "win32": "windows", "linux": "linux"}.get(
        sys.platform, sys.platform
    )


def normalize_args(args: argparse.Namespace, project_root: Path) -> argparse.Namespace:
    if args.windows_onefile:
        args.gui = True
        args.onefile = True
        args.target_platform = "windows"
        if args.icon is None:
            args.icon = project_root / "assets" / "app_icon.ico"

    if not args.gui and not args.cli:
        args.gui = True
    if args.gui and args.cli:
        raise SystemExit("Choose only one of --gui or --cli.")

    actual_platform = host_platform()
    if args.target_platform != "auto" and args.target_platform != actual_platform:
        raise SystemExit(
            "PyInstaller cannot cross-compile. "
            f"Requested {args.target_platform}, but this host is {actual_platform}."
        )
    if args.dmg:
        if actual_platform != "macos":
            raise SystemExit("--dmg is only supported on macOS.")
        if args.cli or args.onefile:
            raise SystemExit("--dmg requires a macOS GUI onedir build.")
    if (
        getattr(args, "bundle_libreoffice", False)
        or getattr(args, "require_libreoffice_bundle", False)
    ) and args.onefile:
        raise SystemExit(
            "Bundled LibreOffice requires an onedir build; onefile would unpack "
            "hundreds of megabytes on every launch."
        )

    if args.name is None:
        args.name = "Doc Media Toolkit"
    if args.experimental and "experimental" not in args.name.lower():
        args.name = f"{args.name} Experimental"

    if args.gui and args.icon is None:
        default_icon = {
            "darwin": project_root / "assets" / "app_icon.icns",
            "win32": project_root / "assets" / "app_icon.ico",
        }.get(sys.platform)
        if default_icon and default_icon.exists():
            args.icon = default_icon
    return args


def binary_name(name: str) -> str:
    return f"{name}.exe" if os.name == "nt" else name


def experimental_dist_root(project_root: Path) -> Path:
    def git_value(*arguments: str) -> str:
        return subprocess.run(
            ["git", *arguments],
            cwd=project_root,
            check=True,
            capture_output=True,
            text=True,
        ).stdout.strip()

    branch = git_value("branch", "--show-current") or "detached"
    commit = git_value("rev-parse", "--short", "HEAD")
    safe_branch = re.sub(r"[^A-Za-z0-9._-]+", "_", branch).strip("._-")
    return project_root / "dist" / "experimental" / safe_branch / commit


def resolve_env_binary(name: str) -> Path | None:
    env_names = [
        f"PPTX_TOOLS_{name.upper()}",
        f"PPTX_OUTPUT_WATERMARK_{name.upper()}",
        f"PPTX_VIDEO_COMPACTOR_{name.upper()}",
    ]
    for env_name in env_names:
        override = os.environ.get(env_name, "").strip()
        if not override:
            continue
        candidate = Path(override).expanduser().resolve()
        if candidate.exists():
            return candidate
        raise SystemExit(f"{env_name} points to a missing file: {candidate}")
    return None


def find_chocolatey_real_binary(name: str) -> Path | None:
    if os.name != "nt":
        return None
    choco_root = Path(os.environ.get("ChocolateyInstall", r"C:\ProgramData\chocolatey"))
    tools_root = choco_root / "lib" / name / "tools"
    exe_name = binary_name(name)
    if not tools_root.exists():
        return None
    for candidate in tools_root.rglob(exe_name):
        if candidate.parent.name.lower() == "bin":
            return candidate.resolve()
    for candidate in tools_root.rglob(exe_name):
        return candidate.resolve()
    return None


def resolve_binary(name: str) -> Path | None:
    env_binary = resolve_env_binary(name)
    if env_binary is not None:
        return env_binary

    found = shutil.which(binary_name(name)) or shutil.which(name)
    if found:
        candidate = Path(found).resolve()
        if os.name == "nt" and "chocolatey" in str(candidate).lower():
            real_binary = find_chocolatey_real_binary(name)
            if real_binary is not None:
                return real_binary
        return candidate
    return find_chocolatey_real_binary(name)


def add_optional_ffmpeg_binaries(
    pyinstaller_args: list[str],
    sep: str,
    *,
    required: bool,
    max_binary_mb: int,
) -> None:
    binaries: dict[str, Path] = {}
    for name in ("ffmpeg", "ffprobe"):
        binary_path = resolve_binary(name)
        if binary_path is None:
            continue
        if max_binary_mb > 0:
            size_mb = binary_path.stat().st_size / 1024 / 1024
            if size_mb > max_binary_mb:
                raise SystemExit(
                    f"Refusing to bundle oversized {name}: {size_mb:.1f} MB at {binary_path}. "
                    "Use FFmpeg essentials, set PPTX_TOOLS_FFMPEG/PPTX_TOOLS_FFPROBE, "
                    "or pass --max-bundled-binary-mb 0 to override."
                )
        binaries[name] = binary_path

    missing = [name for name in ("ffmpeg", "ffprobe") if name not in binaries]
    if missing and required:
        raise SystemExit(
            "Unable to locate required bundled video tool(s): "
            f"{', '.join(missing)}. Install FFmpeg or set PPTX_TOOLS_FFMPEG / PPTX_TOOLS_FFPROBE."
        )
    if missing:
        return

    license_files = find_ffmpeg_license_files(binaries.values())
    if not license_files:
        raise SystemExit(
            "FFmpeg was found, but its LICENSE/COPYING files were not. "
            "Set PPTX_TOOLS_FFMPEG_LICENSE_DIR to the matching distribution license directory."
        )
    for binary_path in binaries.values():
        pyinstaller_args.extend(["--add-binary", f"{binary_path}{sep}."])
    for license_path in license_files:
        pyinstaller_args.extend(["--add-data", f"{license_path}{sep}licenses/ffmpeg"])


def find_ffmpeg_license_files(binary_paths: Iterable[Path]) -> list[Path]:
    explicit = os.environ.get("PPTX_TOOLS_FFMPEG_LICENSE_DIR", "").strip()
    found: dict[Path, None] = {}
    roots_by_binary = (
        [[Path(explicit).expanduser().resolve()]]
        if explicit
        else [list(Path(path).resolve().parents)[:4] for path in binary_paths]
    )
    for roots in roots_by_binary:
        matches: list[Path] = []
        for root in roots:
            if not root.is_dir():
                continue
            for pattern in ("*LICENSE*", "*COPYING*", "*NOTICE*"):
                matches.extend(
                    candidate.resolve()
                    for candidate in root.glob(pattern)
                    if candidate.is_file()
                )
            if matches:
                break
        if not matches:
            return []
        found.update(dict.fromkeys(matches))
    return sorted(found)


def _distribution_notice_files(name: str) -> tuple[list[Path], bool]:
    try:
        package = distribution(name)
    except PackageNotFoundError as exc:
        raise SystemExit(
            f"Required package metadata is missing for license collection: {name}"
        ) from exc

    files: list[Path] = []
    has_license_text = False
    for item in package.files or ():
        parts = tuple(str(part) for part in item.parts)
        lower_parts = tuple(part.lower() for part in parts)
        filename = parts[-1].lower()
        in_dist_info = any(part.endswith(".dist-info") for part in lower_parts)
        is_metadata = in_dist_info and filename == "metadata"
        is_notice = in_dist_info and (
            "licenses" in lower_parts
            or filename.startswith(("license", "copying", "notice"))
        )
        if not (is_metadata or is_notice):
            continue
        resolved = Path(package.locate_file(item)).resolve()
        if resolved.is_file():
            files.append(resolved)
            has_license_text = has_license_text or is_notice
    return sorted(dict.fromkeys(files)), has_license_text


def runtime_distribution_names(project_name: str = "pptx-tools") -> list[str]:
    environment = default_environment()
    environment["extra"] = ""
    pending = [project_name]
    seen = {canonicalize_name(project_name)}
    runtime: dict[str, str] = {}
    while pending:
        package = distribution(pending.pop())
        for raw_requirement in package.requires or ():
            requirement = Requirement(raw_requirement)
            if requirement.marker and not requirement.marker.evaluate(environment):
                continue
            canonical = canonicalize_name(requirement.name)
            if canonical in seen:
                continue
            seen.add(canonical)
            runtime[canonical] = requirement.name
            pending.append(requirement.name)
    return [runtime[name] for name in sorted(runtime)]


def _python_license_file() -> Path:
    base_prefix = Path(sys.base_prefix).resolve()
    version_dir = f"python{sys.version_info[0]}.{sys.version_info[1]}"
    roots = [Path(sys.executable).resolve().parent, base_prefix]
    for root in roots:
        for candidate_root in (root, *list(root.parents)[:4]):
            for filename in ("LICENSE.txt", "LICENSE"):
                candidate = candidate_root / filename
                if candidate.is_file():
                    return candidate
        for candidate in (
            base_prefix / "share" / "doc" / version_dir / "LICENSE.txt",
            base_prefix / "share" / "doc" / version_dir / "LICENSE",
            base_prefix / "lib" / version_dir / "LICENSE.txt",
            base_prefix / "lib" / version_dir / "LICENSE",
        ):
            if candidate.is_file():
                return candidate
    raise SystemExit("Python runtime license file was not found; refusing to package.")


def add_bundle_license_files(
    pyinstaller_args: list[str], sep: str, project_root: Path
) -> None:
    static_files = (
        (project_root / "LICENSE", "licenses/project"),
        (project_root / "THIRD_PARTY_NOTICES.md", "licenses/project"),
        (project_root / "assets/fonts/OFL.txt", "licenses/assets/fonts"),
        (
            project_root / "assets/icons/PHOSPHOR-LICENSE.txt",
            "licenses/assets/icons",
        ),
        (project_root / "licenses/GPL-3.0-only.txt", "licenses/qt"),
        (project_root / "licenses/LGPL-3.0-only.txt", "licenses/qt"),
        (_python_license_file(), "licenses/python-runtime"),
    )
    for source, destination in static_files:
        if not source.is_file():
            raise SystemExit(f"Required bundled license file is missing: {source}")
        pyinstaller_args.extend(["--add-data", f"{source}{sep}{destination}"])

    names = [*runtime_distribution_names(), "PyInstaller"]
    for name in names:
        notice_files, has_license_text = _distribution_notice_files(name)
        canonical_name = canonicalize_name(name)
        if not has_license_text and canonical_name not in QT_LICENSE_DISTRIBUTIONS:
            raise SystemExit(
                f"No license text found in installed distribution metadata: {name}"
            )
        destination = f"licenses/python/{canonical_name}"
        for source in notice_files:
            pyinstaller_args.extend(["--add-data", f"{source}{sep}{destination}"])


def resolve_libreoffice_root(explicit_root: Path | None = None) -> Path | None:
    candidates: list[Path] = []
    if explicit_root is not None:
        candidates.append(explicit_root.expanduser())
    env_root = os.environ.get("PPTX_TOOLS_LIBREOFFICE_ROOT", "").strip()
    if env_root:
        candidates.append(Path(env_root).expanduser())
    if sys.platform == "darwin":
        candidates.append(Path("/Applications/LibreOffice.app"))
    elif sys.platform == "win32":
        for env_name in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(env_name, "").strip()
            if base:
                candidates.extend(
                    [
                        Path(base) / "LibreOffice",
                        Path(base) / "Programs" / "LibreOffice",
                    ]
                )

    for candidate in candidates:
        resolved = candidate.resolve()
        executable = (
            resolved / "Contents" / "MacOS" / "soffice"
            if sys.platform == "darwin"
            else resolved / "program" / "soffice.exe"
        )
        license_candidates = (
            (
                resolved / "Contents" / "Resources" / "LICENSE",
                resolved / "Contents" / "Resources" / "LICENSE.html",
            )
            if sys.platform == "darwin"
            else (
                resolved / "LICENSE",
                resolved / "LICENSE.html",
                resolved / "program" / "LICENSE",
                resolved / "program" / "LICENSE.html",
                resolved / "program" / "license.txt",
            )
        )
        if executable.exists() and any(path.exists() for path in license_candidates):
            return resolved
    return None


def add_optional_libreoffice_runtime(
    pyinstaller_args: list[str],
    sep: str,
    *,
    explicit_root: Path | None,
    required: bool,
) -> Path | None:
    root = resolve_libreoffice_root(explicit_root)
    if root is None:
        if required:
            raise SystemExit(
                "Unable to locate a complete LibreOffice runtime. Install LibreOffice "
                "or pass --libreoffice-root / set PPTX_TOOLS_LIBREOFFICE_ROOT."
            )
        return None
    destination = (
        "libreoffice/LibreOffice.app"
        if sys.platform == "darwin"
        else "libreoffice/LibreOffice"
    )
    pyinstaller_args.extend(["--add-data", f"{root}{sep}{destination}"])
    return root


def add_qt_plugin(
    pyinstaller_args: list[str], source: Path, dest: str, sep: str
) -> None:
    if source.exists():
        pyinstaller_args.extend(["--add-binary", f"{source}{sep}{dest}"])


def add_required_qt_plugins(pyinstaller_args: list[str], sep: str) -> None:
    if os.name != "nt":
        return

    from PySide6.QtCore import QLibraryInfo

    plugins_dir = Path(QLibraryInfo.path(QLibraryInfo.LibraryPath.PluginsPath))
    add_qt_plugin(
        pyinstaller_args,
        plugins_dir / "platforms" / "qwindows.dll",
        "PySide6/Qt/plugins/platforms",
        sep,
    )
    add_qt_plugin(
        pyinstaller_args,
        plugins_dir / "styles" / "qwindowsvistastyle.dll",
        "PySide6/Qt/plugins/styles",
        sep,
    )
    for plugin_name in ("qgif.dll", "qico.dll", "qjpeg.dll", "qsvg.dll", "qtiff.dll"):
        add_qt_plugin(
            pyinstaller_args,
            plugins_dir / "imageformats" / plugin_name,
            "PySide6/Qt/plugins/imageformats",
            sep,
        )


def remove_conflicting_dist_output(dist_root: Path, args: argparse.Namespace) -> None:
    output_path = dist_root / binary_name(args.name)
    if args.onefile:
        if output_path.is_dir():
            shutil.rmtree(output_path)
        return
    if output_path.is_file():
        output_path.unlink()


def project_version(project_root: Path) -> str:
    version_file = project_root / "src" / "pptx_tools" / "__init__.py"
    module = ast.parse(
        version_file.read_text(encoding="utf-8"), filename=str(version_file)
    )
    for statement in module.body:
        if not isinstance(statement, ast.Assign):
            continue
        if not any(
            isinstance(target, ast.Name) and target.id == "__version__"
            for target in statement.targets
        ):
            continue
        if isinstance(statement.value, ast.Constant) and isinstance(
            statement.value.value, str
        ):
            return statement.value.value
    raise SystemExit(f"Unable to read __version__ from {version_file}")


def finalize_macos_bundle_metadata(
    app_bundle: Path,
    *,
    version: str,
    codesign_identity: str = "-",
    minimum_system_version: str = MINIMUM_MACOS_VERSION,
    bundle_identifier: str = BUNDLE_IDENTIFIER,
) -> None:
    plist_path = app_bundle / "Contents" / "Info.plist"
    if not plist_path.exists():
        raise SystemExit(f"macOS bundle metadata is missing: {plist_path}")
    with plist_path.open("rb") as plist_file:
        metadata = plistlib.load(plist_file)
    metadata.update(
        {
            "CFBundleIdentifier": bundle_identifier,
            "CFBundleShortVersionString": version,
            "CFBundleVersion": version,
            "LSMinimumSystemVersion": minimum_system_version,
        }
    )
    with plist_path.open("wb") as plist_file:
        plistlib.dump(metadata, plist_file)
    sign_command = [
        "codesign",
        "--force",
        "--deep",
        "--sign",
        codesign_identity,
    ]
    if codesign_identity != "-":
        sign_command.extend(["--options", "runtime", "--timestamp"])
    subprocess.run([*sign_command, str(app_bundle)], check=True)
    subprocess.run(
        ["codesign", "--verify", "--deep", "--strict", str(app_bundle)],
        check=True,
    )


def macos_arch_label() -> str:
    machine = os.uname().machine.lower() if hasattr(os, "uname") else "unknown"
    return {"arm64": "arm64", "x86_64": "x64"}.get(machine, machine)


def default_dmg_output_path(project_root: Path, app_name: str) -> Path:
    return project_root / "dist" / f"{app_name}-macOS-{macos_arch_label()}.dmg"


def create_macos_dmg(app_bundle: Path, dmg_path: Path, volume_name: str) -> None:
    if not app_bundle.exists():
        raise SystemExit(f"App bundle not found for DMG creation: {app_bundle}")

    dmg_path.parent.mkdir(parents=True, exist_ok=True)
    dmg_path.unlink(missing_ok=True)

    with tempfile.TemporaryDirectory(prefix="ppt_tools_dmg_") as temp_dir:
        staging_dir = Path(temp_dir) / "staging"
        staging_dir.mkdir(parents=True, exist_ok=True)
        shutil.copytree(app_bundle, staging_dir / app_bundle.name, symlinks=True)
        os.symlink("/Applications", staging_dir / "Applications")
        subprocess.run(
            [
                "hdiutil",
                "create",
                "-volname",
                volume_name,
                "-srcfolder",
                str(staging_dir),
                "-ov",
                "-format",
                "UDZO",
                str(dmg_path),
            ],
            check=True,
        )


def notarize_macos_dmg(dmg_path: Path, keychain_profile: str) -> None:
    if not keychain_profile.strip():
        return
    subprocess.run(
        [
            "xcrun",
            "notarytool",
            "submit",
            str(dmg_path),
            "--keychain-profile",
            keychain_profile,
            "--wait",
        ],
        check=True,
    )
    subprocess.run(["xcrun", "stapler", "staple", str(dmg_path)], check=True)


def main() -> int:
    require_release_python()
    args = normalize_args(parse_args(), Path(__file__).resolve().parents[1])
    project_root = Path(__file__).resolve().parents[1]
    dist_root = (
        experimental_dist_root(project_root)
        if args.experimental
        else project_root / "dist"
    )
    entrypoint = project_root / (
        "pptx_tools_gui.py" if args.gui else "pptx_tools_cli.py"
    )
    sep = ";" if os.name == "nt" else ":"

    pyinstaller_args = [
        "--noconfirm",
        "--windowed" if args.gui else "--console",
        "--name",
        args.name,
        "--distpath",
        str(dist_root),
        "--workpath",
        str(project_root / "build" / "pyinstaller"),
        "--specpath",
        str(project_root / "build" / "spec"),
        "--paths",
        str(project_root / "src"),
        "--add-data",
        f"{project_root / 'assets'}{sep}assets",
        "--add-data",
        f"{project_root / 'config'}{sep}config",
    ]
    add_bundle_license_files(pyinstaller_args, sep, project_root)
    if args.bundle_ffmpeg or args.require_ffmpeg_bundle:
        add_optional_ffmpeg_binaries(
            pyinstaller_args,
            sep,
            required=args.require_ffmpeg_bundle,
            max_binary_mb=args.max_bundled_binary_mb,
        )
    if args.bundle_libreoffice or args.require_libreoffice_bundle:
        add_optional_libreoffice_runtime(
            pyinstaller_args,
            sep,
            explicit_root=args.libreoffice_root,
            required=args.require_libreoffice_bundle,
        )
    if args.gui:
        pyinstaller_args.extend(
            [
                "--hidden-import",
                "PySide6.QtCore",
                "--hidden-import",
                "PySide6.QtGui",
                "--hidden-import",
                "PySide6.QtWidgets",
                "--hidden-import",
                "PySide6.QtMultimedia",
                "--hidden-import",
                "PySide6.QtMultimediaWidgets",
                "--hidden-import",
                "pptx_output_watermark.gui",
                "--hidden-import",
                "pptx_video_compactor_gui",
                "--hidden-import",
                "pptx_tools.image_manager_gui",
                "--hidden-import",
                "pptx_tools.image_manager",
                "--hidden-import",
                "pptx_tools.ai_client",
                # Lazy-loaded document backends (see CLI_HIDDEN_IMPORTS).
                "--hidden-import",
                "docx_image_compactor",
                "--hidden-import",
                "pdf_image_compactor",
                "--hidden-import",
                "xlsx_image_compactor",
                "--hidden-import",
                "pikepdf",
                "--runtime-hook",
                str(project_root / "scripts" / "pyinstaller_runtime_hook.py"),
            ]
        )
        if sys.platform == "win32":
            pyinstaller_args.extend(
                [
                    "--hidden-import",
                    "comtypes",
                    "--hidden-import",
                    "comtypes.stream",
                ]
            )
        add_required_qt_plugins(pyinstaller_args, sep)
    else:
        for module in CLI_HIDDEN_IMPORTS:
            pyinstaller_args.extend(["--hidden-import", module])
    if args.clean:
        pyinstaller_args.append("--clean")
    if args.icon:
        pyinstaller_args.extend(["--icon", str(args.icon.expanduser().resolve())])
    pyinstaller_args.append("--onefile" if args.onefile else "--onedir")
    pyinstaller_args.append(str(entrypoint))
    remove_conflicting_dist_output(dist_root, args)

    try:
        import PyInstaller.__main__ as pyinstaller
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "PyInstaller is not installed. Install build deps first: "
            "python -m pip install -e '.[build]'"
        ) from exc
    pyinstaller.run(pyinstaller_args)

    if sys.platform == "darwin" and args.gui:
        finalize_macos_bundle_metadata(
            dist_root / f"{args.name}.app",
            version=project_version(project_root),
            codesign_identity=args.codesign_identity,
            bundle_identifier=(
                EXPERIMENTAL_BUNDLE_IDENTIFIER
                if args.experimental
                else BUNDLE_IDENTIFIER
            ),
        )

    if args.dmg:
        app_bundle = dist_root / f"{args.name}.app"
        dmg_path = (
            args.dmg_output.expanduser().resolve()
            if args.dmg_output is not None
            else (
                dist_root / f"{args.name}-macOS-{macos_arch_label()}.dmg"
                if args.experimental
                else default_dmg_output_path(project_root, args.name)
            )
        )
        create_macos_dmg(app_bundle, dmg_path, args.name)
        notarize_macos_dmg(dmg_path, args.notary_profile)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
