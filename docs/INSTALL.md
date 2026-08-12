# Installation and platform support

[简体中文](INSTALL.zh-CN.md)

## Current distribution status

The source release is public. Prebuilt macOS and Windows packages remain release
candidates until their platform signatures, notarization where applicable,
artifact-level license inventory, malware scan, SBOM, and bundled-FFmpeg source
delivery evidence pass the public-binary gate. Do not treat an unsigned candidate
as a trusted installer.

The next candidates bundle FFmpeg 8.1.2 built from pinned FFmpeg/x264 source and
publish the matching corresponding-source archive on the same Release. The older
0.2.0 Homebrew/Gyan candidates remain withdrawn and must not be republished.

## Supported and tested environments

| Platform | Current evidence | Public package status |
| --- | --- | --- |
| macOS 13 or later, Apple Silicon | Native build, package audit, DMG verification, and offscreen smoke test | Candidate only; ad-hoc signed and not notarized |
| macOS 13 or later, Intel | Native build, package audit, DMG verification, and offscreen smoke test | Candidate only; ad-hoc signed and not notarized |
| 64-bit Windows | x64 build and offscreen smoke test on the Windows Server 2025 hosted build environment | Candidate only; supported Windows client baseline and Authenticode signature are not yet finalized |
| Linux | Source execution may work for non-platform-specific paths | No supported desktop package |

## Install from source

Python 3.12, Git, and
[uv](https://docs.astral.sh/uv/getting-started/installation/) must already be
installed. Feature-specific external runtimes are listed below.

macOS or Linux:

```bash
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
uv sync --locked --all-extras
.venv/bin/pptx-tools-gui
```

Windows PowerShell:

```powershell
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
uv sync --locked --all-extras --python 3.12
.venv\Scripts\Activate.ps1
pptx-tools-gui
```

## Choose a future signed package

- `macOS-arm64.dmg`: Apple Silicon Macs.
- `macOS-x64.dmg`: Intel Macs.
- `windows-x64-portable.zip`: 64-bit Windows; extract the complete directory
  before running the executable. The directory form keeps Qt libraries and license
  files replaceable and inspectable.
- `SHA256SUMS.txt`: hashes for transport-integrity verification. A matching hash is
  necessary but does not replace code signing or notarization.

After a signed macOS package is published, drag the app from the DMG to
Applications. A valid release should pass Gatekeeper normally. Do not disable
Gatekeeper globally. After a signed Windows package is published, verify the
publisher in the file properties before running it; do not bypass SmartScreen for
an unknown or mismatched publisher.

## Feature-specific runtimes

| Workflow | Runtime |
| --- | --- |
| Video compression, quality audit, video watermark, and video-library transcoding | Signed release packages bundle the source-pinned FFmpeg 8.1.2 runtime; source installs resolve configured or system binaries |
| PPTX to PDF/image export on Windows | PowerPoint or WPS first, then LibreOffice fallback |
| DOCX to PDF on Windows | Word or WPS first, then LibreOffice fallback |
| PPTX to PDF/image export on macOS | LibreOffice first, then Keynote fallback with Automation permission |
| DOCX to PDF on macOS | LibreOffice first, then Pages fallback with Automation permission |
| PDF embedded-image classification | Poppler `pdfimages` when that compression path is used |

Office, WPS, LibreOffice, Keynote, Pages, and system Poppler are external programs
and are not part of the standard package.

## Upgrade and data safety

- Normal compression and watermark operations save a new output by default.
- Before upgrading the application, finish active jobs and copy each complete video
  or image library directory to a backup location. The manifest and its `.bak` file
  are not sufficient without the accompanying managed media.
- Version 0.2.x uses schema version 1 for both video and image libraries. A future
  incompatible schema must ship an explicit migration before the version is raised.
- Keep the old application available until representative documents and each library
  pass a read-only health check in the new version.

## Verify an artifact

Use the SHA-256 tool provided by the operating system and compare the result with
`SHA256SUMS.txt`. Also verify the platform signature and the release page's exact
tag and commit. Report mismatches through the repository security policy instead of
running the file.
