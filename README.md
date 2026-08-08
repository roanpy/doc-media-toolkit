<div align="center">
  <img src="assets/app_icon_v2.png" width="88" alt="Doc Media Toolkit icon">
  <h1>Doc Media Toolkit</h1>
  <p><strong>Reduce oversized PPTX files safely and turn document videos and images into reusable, managed assets.</strong></p>

  English · [简体中文](README.zh-CN.md)

  [![Status: Stable](https://img.shields.io/badge/status-stable-brightgreen.svg)](#project-status)
  [![Version 0.2.0](https://img.shields.io/badge/version-0.2.0-2563eb.svg)](src/pptx_tools/__init__.py)
  [![Python 3.10–3.13](https://img.shields.io/badge/Python-3.10--3.13-3776AB.svg?logo=python&logoColor=white)](pyproject.toml)
  [![macOS | Windows](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey.svg)](#quick-start)
  [![MIT License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
</div>

![Doc Media Toolkit English compression workspace using anonymous synthetic test files](docs/images/doc-media-toolkit-compression-en.png)

<p align="center"><sub>Actual UI rendered from the current source tree with isolated settings and anonymous synthetic test files. No real project, path, or business content is shown.</sub></p>

## The problem it solves

Embedded video and high-resolution images can make PPTX files unnecessarily large. Manual compression makes it difficult to balance a target size, playback compatibility, and visual quality. The same media is often duplicated across documents, while the authoritative high-quality source becomes hard to track.

Doc Media Toolkit is built for that workflow:

- Compress PPTX video and images by preset or target size, using display area, reuse, and media characteristics to allocate the budget.
- Audit visual quality after compression and re-encode below-threshold assets at a higher tier instead of optimizing size blindly.
- Archive video and image assets from documents, deduplicate exact content, preserve source/version relationships, and safely restore higher-quality media.
- Optionally use a compatible AI model to organize, rename, classify, and tag video and image assets. Vision-capable models can analyze confirmed previews, while every result remains a reviewable suggestion.
- Create image-based PPTX files with text or image watermarks, including watermarking and reinserting embedded videos.
- Accept DOCX, PDF, XLSX, standalone images, and standalone videos as supporting formats in the same document-media workflow.

## Key capabilities

| Capability | What it provides |
| --- | --- |
| Size-aware compression | PPTX-first target sizing, independent video/image presets, CPU/GPU strategies, SSIM quality audits, and tiered optimization |
| Video and image libraries | Document media archiving, SHA-256 deduplication, conservative version matching, source tracking, and safe media restoration |
| AI-assisted organization | Optional naming, category, tag, and merge-candidate suggestions; previews are sent only when vision is supported and explicitly enabled |
| Watermark export | Batch PDF/PPTX/image/video output, text or image watermarks, image-based PPTX, and embedded-video watermark reinsertion |
| Local-first safety | Save-as by default; similarity, AI, merge, and cleanup decisions require review, with recoverable quarantine before deletion |

This is not a general-purpose video or image editor. Its media features exist to reduce document size, preserve playable delivery, and manage document assets.

## Scope and coverage

| Area | Status | Boundary |
| --- | --- | --- |
| PPTX media compression | Core | Target-size planning, quality protection, and CPU/GPU video paths are the primary workflow. |
| Watermark export | Core | Image-based PPTX plus document, image, and video watermark output; source files stay unchanged by default. |
| Video asset library | Core | Archive, deduplicate, track, and restore document video sources with conservative matching. |
| DOCX / PDF / XLSX media handling | Supporting | Uses format-specific backends and external office/PDF runtimes where required; layout and text are preserved, but this is not full document reflow. |
| Image asset library | Ancillary | Available for common image workflows, with a smaller validated surface than PPTX/video paths. |
| AI organization | Optional | Generates reviewable names, tags, and grouping suggestions only; it does not silently rename, merge, or delete assets. |
| General editing, cloud sync, and OCR | Out of scope | Use a dedicated editor, storage service, or OCR tool for these jobs. |

## Interface preview

### Watermark export

![English watermark export workspace](docs/images/doc-media-toolkit-watermark-en.png)

### AI and library settings

![English AI and library settings using an invalid example endpoint with no API key or local path](docs/images/doc-media-toolkit-settings-en.png)

Every image above is captured from the current source using anonymous synthetic files and invalid example configuration. No real project, API key, local path, or business content is shown.

## Quick start

Python 3.12 is recommended. The complete desktop install includes the required Python packages; FFmpeg, Poppler, and office runtimes depend on the selected workflow and platform.

macOS / Linux:

```bash
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
bash setup_env.sh
.venv/bin/pptx-tools-gui
```

Windows PowerShell:

```powershell
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev,build]"
pptx-tools-gui
```

The compatibility CLI name remains `pptx-tools`:

```bash
pptx-tools --help
pptx-tools compact --help
pptx-tools watermark --help
pptx-tools videos --help
pptx-tools images --help
```

## Safety and privacy

- Normal processing writes a new file instead of overwriting the source.
- Library removal, merging, and cleanup move managed copies to recoverable quarantine first.
- Filename, folder, duration, or resolution alone never proves that two media files are identical.
- AI organization suggestions are optional. API keys stay in process memory; complete documents, complete videos, and local paths are not sent by default.
- A public DMG/EXE is not an “MIT-only” artifact. Qt, FFmpeg, PDFium, Python, and platform-native libraries still require artifact-level license review and signing.

See the [dependency rationale](docs/DEPENDENCIES.md), [licensing guide](docs/LICENSING.md), and [security policy](SECURITY.md).

## Project status

The project is published as a **stable release**. PPTX compression, watermarking, the video asset library, and the primary document-compatibility paths have been confirmed. Image asset management is an ancillary capability and has not yet received the same depth of validation as the core workflows. The application shell, watermark workspace, compression workspace, and help center support Simplified Chinese and English; the video and image library workspaces remain Chinese-first. First launch defaults to English. Set `PPTX_TOOLS_LANG=zh` before launch to use Chinese.

The public repository is `doc-media-toolkit`; the Python distribution and CLI retain `pptx-tools` for compatibility. A signed installer has not yet been published as a GitHub Release.

## Roadmap and participation

Near-term priorities are to complete English coverage for the video and image libraries; establish license, signing, and reproducibility gates for macOS and Windows installers; continue calibrating target-size and quality protection with anonymous fixtures; and stabilize DOCX/PDF/XLSX support without changing document layout or formatting.

Reproducible [bug reports](https://github.com/roanpy/doc-media-toolkit/issues/new?template=bug_report.yml) and focused [feature requests](https://github.com/roanpy/doc-media-toolkit/issues/new?template=feature_request.yml) are welcome. Remove confidential content, personal information, and local paths from documents, screenshots, and logs. Report security issues privately under the [security policy](SECURITY.md).

## Documentation

- [Complete Chinese user guide](docs/USER_GUIDE.zh-CN.md)
- [Architecture and module boundaries](docs/ARCHITECTURE.md)
- [Smart target-size compression specification](docs/SMART_TARGET_COMPRESSION.md)
- [Quality and release gates](docs/QUALITY_GATES.md)
- [Release and packaging guide](docs/RELEASE.md)
- [Contributing](CONTRIBUTING.md) · [Code of Conduct](CODE_OF_CONDUCT.md) · [Third-party notices](THIRD_PARTY_NOTICES.md)

## Development and validation

```bash
QT_QPA_PLATFORM=offscreen PYTHONPATH=src:tests .venv/bin/python scripts/run_tests_isolated.py
.venv/bin/ruff check src tests scripts
.venv/bin/python scripts/check_public_safety.py
```

GitHub workflows are manual-only. Local verification is the default quality gate.

## AI-assisted development disclosure

The project is owner-led and developed primarily with assistance from **OpenAI Codex** for design, implementation, testing, review, and documentation. **Google Gemini** and **Anthropic Claude Code** provide additional generation and cross-checking. The project owner reviews and tests all AI-assisted output and remains responsible for maintenance; AI tools are not copyright holders or licensing entities.

## License

The project's own source is released under the [MIT License](LICENSE). Third-party libraries, fonts, icons, and bundled runtimes retain their respective licenses; see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
