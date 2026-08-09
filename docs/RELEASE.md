# Release and packaging guide

Public source releases and public binary releases use separate gates. The manual
workflow builds private candidates only; it has read-only repository permission and
contains no GitHub Release publication job.

## Artifacts

Recommended candidate package names:

- `Doc Media Toolkit-macOS-arm64.dmg`
- `Doc Media Toolkit-macOS-x64.dmg`
- `Doc Media Toolkit-macOS-arm64.zip`
- `Doc Media Toolkit-macOS-x64.zip`
- `Doc Media Toolkit-windows-x64-portable.zip`
- `SBOM-macOS-arm64.cdx.json`
- `SBOM-macOS-x64.cdx.json`
- `SBOM-windows-x64.cdx.json`
- `RELEASE-AUDIT-<platform>.md` and `.json`
- `FFMPEG-<platform>.txt`
- `Doc-Media-Toolkit-FFmpeg-8.1.2-<platform>-corresponding-source.tar.gz`

## Build Commands

macOS GUI:

```bash
scripts/build_ffmpeg_runtime.sh release-assets/ffmpeg-runtime
export PPTX_TOOLS_FFMPEG="$PWD/release-assets/ffmpeg-runtime/bin/ffmpeg"
export PPTX_TOOLS_FFPROBE="$PWD/release-assets/ffmpeg-runtime/bin/ffprobe"
export PPTX_TOOLS_FFMPEG_LICENSE_DIR="$PWD/release-assets/ffmpeg-runtime/licenses"
python scripts/build_standalone.py --gui --clean --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --dmg --name "Doc Media Toolkit"
```

Optional custom DMG path:

```bash
python scripts/build_standalone.py --gui --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --dmg --dmg-output "/absolute/path/Doc Media Toolkit-macOS-arm64.dmg" --name "Doc Media Toolkit"
```

Optional private/internal complete offline build:

```bash
python scripts/build_standalone.py --gui --target-platform macos --bundle-ffmpeg --require-ffmpeg-bundle --bundle-libreoffice --require-libreoffice-bundle --dmg --name "Doc Media Toolkit"
```

Use `--libreoffice-root` or `PPTX_TOOLS_LIBREOFFICE_ROOT` to select the complete
runtime. LibreOffice bundling is intentionally limited to onedir builds. Do not
publish this variant until its complete notices and matching source-code access
have passed a separate artifact-level license audit.

Windows GUI onedir candidate:

```powershell
python scripts\build_standalone.py --gui --target-platform windows --bundle-ffmpeg --require-ffmpeg-bundle --name "Doc Media Toolkit"
Compress-Archive -Path "dist\Doc Media Toolkit" -DestinationPath "Doc Media Toolkit-windows-x64-portable.zip"
```

The Windows GUI build defaults to `assets/app_icon.ico`, and the packaged app name defaults to `Doc Media Toolkit` when `--name` is omitted. `--windows-onefile` remains available for private testing, but the public-binary gate rejects it because the current one-file layout does not provide the verified Qt replacement/relinking path required by this project.

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

- `--windows-onefile` is only for private Windows GUI testing.
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
  runtime and its license notices for private/internal use; public redistribution
  additionally requires independent verification of its matching source-access
  obligations.
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
- Public candidates must use `scripts/build_ffmpeg_runtime.sh`, not a Homebrew or Gyan prebuilt binary. The script pins FFmpeg 8.1.2, x264, and zlib 1.3.2 sources by SHA-256, enables libx264 plus VideoToolbox on macOS or Media Foundation on Windows, records the exact configuration, and emits the matching corresponding-source archive. Public Windows distribution uses an onedir portable ZIP so Qt libraries remain replaceable and inspectable.
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
- Candidate jobs build FFmpeg 8.1.2, x264, and zlib 1.3.2 from pinned, hashed source; smoke-launch the GUI; verify bundled tools and macOS DMG/signature structure; and emit the corresponding-source archive, per-platform checksums, SBOMs, native-file inventories, dependency audits, and FFmpeg build information. They never publish a GitHub Release.
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
python scripts/release_audit.py --check --public-binary \
  --dist-dir release-assets \
  --with-sbom release-assets/sbom.cdx.json \
  --evidence release-assets/public-binary-evidence.json
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
- Fail-closed public-binary evidence validation with `--public-binary`. This mode
  requires pip-audit, a generated SBOM, artifact hashes, valid signature evidence,
  notarization evidence for macOS, a clean malware-scan report, native inventory,
  and complete bundled-FFmpeg source-delivery evidence. Windows one-file packages
  are rejected.

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

### Public-binary evidence schema

`--evidence` accepts JSON with schema
`doc-media-toolkit.public-binary-evidence.v1`. Every item under `artifacts` records
the artifact `path`, `sha256`, `platform`, `architecture`, and `package_type`, plus
a trusted signature type (`developer-id` or `authenticode`) and verified sidecar
files for `signature`, `notarization` on macOS, `malware_scan`,
`sbom`, and `native_inventory`. A bundled FFmpeg entry also records `version`, full
`configuration`, `license`, and a hashed `corresponding_source` archive. Paths are
resolved inside `--dist-dir`; path escape, missing files, or hash mismatches fail.

The evidence file is an index of independently produced reports, not a substitute
for signature, notarization, malware, or legal review.

### Signing, FFmpeg, and notarization boundaries

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
- The preferred product path is a repeatable, source-pinned minimal GPL FFmpeg build because
  the compression quality contract depends on libx264. The release source bundle
  must include the exact FFmpeg and linked GPL component sources, build scripts,
  patches, and recorded configuration for each platform. Merely linking to an
  upstream commit or including the GPL text does not satisfy this repository's gate.
- Benchmark corpora (samples and manifest) must live outside the working tree and
  are git-ignored. Cross-platform benchmark results must be produced on the matching
  platform because GPU probing, FFmpeg paths, and encoder availability are
  platform-specific.
