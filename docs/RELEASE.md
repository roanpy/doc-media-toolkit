# Release Notes

## Artifacts

Recommended package names:

- `Doc Media Toolkit-macOS-arm64.dmg`
- `Doc Media Toolkit-macOS-x64.dmg`
- `Doc Media Toolkit-macOS-arm64.zip`
- `Doc Media Toolkit-macOS-x64.zip`
- `Doc Media Toolkit-windows-x64.zip`
- `Doc Media Toolkit-windows-x64-onefile.zip`
- `Doc Media Toolkit-windows-x64-onefile.exe`
- `Doc Media Toolkit-cli-macOS-arm64.tar.gz`
- `Doc Media Toolkit-cli-macOS-x64.tar.gz`
- `Doc Media Toolkit-cli-windows-x64.zip`
- `Doc Media Toolkit-cli-windows-x64.exe`

## Build Commands

macOS GUI:

```bash
python scripts/build_standalone.py --gui --clean --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --dmg --name "Doc Media Toolkit"
```

Optional custom DMG path:

```bash
python scripts/build_standalone.py --gui --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --dmg --dmg-output "/absolute/path/Doc Media Toolkit-macOS-arm64.dmg" --name "Doc Media Toolkit"
```

Optional complete offline build:

```bash
python scripts/build_standalone.py --gui --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --bundle-libreoffice --require-libreoffice-bundle --dmg --name "Doc Media Toolkit"
```

Use `--libreoffice-root` or `PPTX_TOOLS_LIBREOFFICE_ROOT` to select the complete
runtime. LibreOffice bundling is intentionally limited to onedir builds.

Windows GUI one-file:

```powershell
python scripts\build_standalone.py --windows-onefile --bundle-ffmpeg --require-ffmpeg-bundle --name "Doc Media Toolkit"
```

The Windows GUI build defaults to `assets/app_icon.ico`, and the packaged app name defaults to `Doc Media Toolkit` when `--name` is omitted.

CLI one-file:

```bash
python scripts/build_standalone.py --cli --onefile --clean --bundle-ffmpeg --require-ffmpeg-bundle --name "Doc Media Toolkit"
```

Isolated experimental macOS GUI (use for compression branches; never publish as
the formal app):

```bash
PPTX_TOOLS_EXPERIMENTAL=1 .venv/bin/python scripts/build_standalone.py --experimental --gui --clean --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --name "Doc Media Toolkit"
```

`--experimental` appends `Experimental`, uses bundle ID
`com.roanpy.doc-media-toolkit.experimental`, and writes to
`dist/experimental/<sanitized-branch>/<commit>/`. At runtime it isolates
QSettings, logs and temporary names; default user outputs end in
`_experimental`. Experimental artifacts must not replace, sign as, upload over,
or auto-update the formal application.

Notes:

- `--windows-onefile` is only for Windows GUI.
- Windows CLI should still use `--cli --onefile`, but must be built on a Windows host.
- `--dmg` is only for macOS GUI onedir builds.

## Distribution Notes

- Desktop packages should bundle `ffmpeg` and `ffprobe`.
- Validate bundled tools, lazy document backends, `pikepdf`, app signature,
  offscreen launch, and DMG structure directly. File size is not a completeness
  test because compression and runtime versions change it.
- Media quality audit also depends on bundled `ffmpeg`; it now evaluates every processed asset and no longer skips items only because the size delta is small.
- WMV/legacy video assets can be scored only when the bundled or system `ffmpeg` can decode them; otherwise the audit reports a clear decode/SSIM failure with the FFmpeg stderr tail instead of a generic missing-score error.
- Incremental optimize depends on the same `ffmpeg` resolution path. Post-compression auto-audit, manual audit, and post-optimize re-audit should all resolve the same bundled or external binary.
- Media compression persists separate base SSIM thresholds: video defaults to `0.95` (High), `0.93` (Balanced), and `0.90` (Low volume), images to `0.99`; a user-edited video threshold is preserved across preset changes and sessions. Safe target mode defaults to CPU two-pass. Target GPU is off by default; ordinary presets separately persist auto/CPU/prefer-GPU strategy. Forced output is only offered after safe output exceeds target and still enforces absolute redlines.
- GUI windows initialize the shared UI font before `QMainWindow` construction, so packaged builds and offscreen automation should not emit the old `Sans Serif` alias warning.
- Standard packages do not bundle Office/WPS/LibreOffice/Keynote/Pages. An
  explicit `--bundle-libreoffice` onedir build includes the complete LibreOffice
  runtime and its license notices.
- Every standalone build bundles the project MIT license, third-party inventory,
  Python runtime license, Qt LGPL/GPL texts, asset licenses, and installed Python
  package metadata/license files under `licenses/`; missing required notices fail
  the build.
- macOS DMG builds contain only `Doc Media Toolkit.app` and an `Applications` shortcut, not the local `/Applications` contents.
- `PPTX -> PDF`: Windows uses PowerPoint/WPS COM first and falls back to LibreOffice; macOS uses LibreOffice first and falls back to Keynote.
- `DOCX -> PDF`: Windows uses Word/WPS COM first and falls back to LibreOffice; macOS uses LibreOffice first and falls back to Pages.
- macOS fallback checks Automation permission separately from install state. If Keynote or Pages is installed but not authorized, the GUI should offer a System Settings action instead of a download link.
- `PDF` input bypasses external PDF export engines and goes straight into editable-PDF watermarking or image-PDF rebuilding.
- `pypdfium2` replaces external `pdftoppm`.
- Windows release builds should use FFmpeg essentials and avoid accidentally bundling oversized full static binaries. A one-file executable larger than 200MB is acceptable when PySide6 and FFmpeg are bundled.
- Public one-file distribution remains blocked until the Qt LGPL replacement/relinking path and the exact native-library inventory are verified. Prefer onedir for the first public binary release.
- macOS local packages are not notarized unless Developer ID signing and notarization are configured separately.
- macOS builds copy `pptx_tools.__version__` into `CFBundleShortVersionString` and `CFBundleVersion`, use `com.roanpy.doc-media-toolkit` as the bundle identifier, then re-sign and strictly verify the app before DMG creation.
- Media audit and incremental optimize support batch runs. Incremental optimize rewrites only failed assets into the previous compressed PPTX and should not emit a third deliverable.
- The compression GUI can automatically audit newly completed outputs and optimize/re-audit failed PPTX assets until they pass or the High compressed result is retained.
- Incremental optimize policy is fixed: `Balanced/Low -> High`; if High still misses the SSIM threshold, the compressed result is retained and the log stops further automatic retries. Decode/structure audit failures still restore the original media.
- Watermark export runtime temp roots (`pptx_output_watermark_*`) should be cleaned on success/close and reaped on next launch after abnormal termination.
- Switching tools preserves each tool window's settings and logs. Shutdown cancels tracked child process groups and waits for worker, preview, and audit threads without `QThread.terminate()`.
- The video-library workspace stores only video media, while deck records retain PPTX hashes, path aliases, output records, and slide/shape anchors. PPTX files and compact media are not copied into the library.
- Source import deduplicates exact SHA-256 matches globally and conservatively groups unique re-encodes only when duration, aspect, video-frame, and audio-spectrum fingerprints agree. The default library policy keeps sources at or below 1080p byte-exact and converts larger sources to high-quality 1080p; users can select full originals. Original and compact PPTX hashes remain aliases without retaining low-quality physical copies.
- External videos can be batch-imported into the library. Unique content matches become versions of an existing family, unmatched videos create a family, and ambiguous matches are reported without mutation. Manual version imports use the same identity guard.
- Source import accepts a validated relative category below `media/`. Renaming a family also renames its authoritative source file; moving or renaming files does not change hash/family identity.
- Compression can archive PPTX video sources before encoding. Library upgrade prefers exact known hashes; unregistered re-encodes require one unique conservative match across duration, aspect ratio, five perceptual frame hashes, and the audio spectrum. Filename, duration, or resolution alone is never identity, and clips with different audio or edited duration are rejected.
- Upgrade keeps the existing media part whenever possible, preserving slide XML, relationships, shapes, posters, geometry, and playback timing. Non-compatible sources are converted transiently to H.264/AAC within a no-upscale 1080p envelope.
- Video library preferences use platform-native Qt settings. Rotating application logs remain under the platform application-data directory.
- Libraries keep a last-valid `video-project.json.bak`, reject stale concurrent saves, and relink renamed or moved video files by SHA-256.
- Video and image cleanup use a recoverable intent log before moving files. New quarantine paths are project-relative, restore verifies SHA-256, and any path outside the managed media/cleanup roots fails closed.
- Windows release builds pin the FFmpeg essentials archive and verify its SHA-256. Release jobs smoke-launch the GUI, verify bundled tools and macOS DMG/signature structure, and publish per-platform `SHA256SUMS` files.
- Windows release builds download Poppler `26.02.0-0` from the pinned `oschwartz10612/poppler-windows` release and verify SHA-256 (`993e4a94376ed712fafc7058d724ea0b943d118bbd2305cd9ed55174eb85cda5`) before PDF tests. Poppler remains a build/test runtime and is not bundled into the standard app.
- Watermark, compression, video-library, and image-library UI logs share the same rotating application log.
- Both `ci.yml` and `release.yml` are manual-only. Pushing commits or tags does not consume Actions minutes; a maintainer must explicitly dispatch a workflow. Both workflows enforce `ruff check src tests scripts` and `ruff format --check src tests scripts` before tests/builds; Markdown examples are reviewed separately and are not treated as Python source.

## Local Pre-Release Audit & Compression Benchmark

These are local, manual entry points that do NOT depend on GitHub Actions. Run them
on each build host before producing a release artifact.

### Dependency & artifact audit

```bash
python scripts/release_audit.py --check
python scripts/release_audit.py --check --with-sbom dist/sbom.json
```

`release_audit.py` covers:

- Git branch/commit provenance and a clean working-tree requirement before release.
- `uv lock --check` and `uv sync --locked --dry-run` (lockfile integrity).
- `pip-audit` as an **external, opt-in tool**: if absent it is reported as a
  skipped step and is never added as a runtime dependency.
- Optional CycloneDX SBOM via `uv export --format cyclonedx1.5` when explicitly requested
  with `--with-sbom`.
- Packaged binary version (from `pptx_tools.__version__`) plus SHA-256 hashes for
  every file under `dist/`.

It writes a Markdown checklist (`release-audit.md`) and a JSON sidecar. A dirty
working tree fails the audit so the package can be traced to a reproducible commit.
It does not build, sign, notarize, or publish.

### Manifest-driven compression benchmark

```bash
python scripts/run_compression_benchmark.py --manifest /path/to/manifest.json
python scripts/run_compression_benchmark.py --self-check
```

`run_compression_benchmark.py` reuses `compact_input_path` to run the smart-target
compression core over sanitized samples referenced by an external manifest. It
records target capacity error, actual capacity, per-asset quality/structure results,
correction rounds, elapsed time, and CPU/GPU/fallback signals. Sample files and the
manifest are NOT committed to Git; reports keep file names and content hashes but do not
write absolute input paths. Contract:
`docs/COMPRESSION_BENCHMARK.md`.

### Signing & notarization boundaries

- Cross-platform artifacts must be built on the target platform; PyInstaller does
  not cross-compile. macOS local validation is not evidence that Windows builds run.
- The manual GitHub release workflow installs and verifies Poppler on every build
  host because PDF image-area classification depends on `pdfimages`; Poppler is
  used by the build/test environment and is not bundled into the standard app.
- macOS packages are ad-hoc signed by default. Developer ID signing and notarization
  require a separate keychain profile (`--notary-profile`); the local audit does not
  perform notarization.
- Windows artifacts must be signed with a certificate on a Windows host; the local
  audit only records hashes and does not sign.
- Source licensing and binary licensing are separate gates. Follow
  `docs/LICENSING.md`; an FFmpeg build containing `--enable-gpl` requires the
  corresponding GPL source-delivery obligations for that exact binary.
- Benchmark corpora (samples and manifest) must live outside the working tree and
  are git-ignored. Cross-platform benchmark results must be produced on the matching
  platform because GPU probing, FFmpeg paths, and encoder availability are
  platform-specific.
