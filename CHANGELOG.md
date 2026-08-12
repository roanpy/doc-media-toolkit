# Changelog

All notable changes to Doc Media Toolkit are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

## [0.2.3] - 2026-08-12

### Security

- Normalize OOXML relationship separators before path validation, execute
  LibreOffice with an argv list instead of a batch shell, and refuse macro-enabled
  Office documents when a Windows COM engine cannot disable automation macros.
- Bound PDF bitmap allocation and page dimensions before rendering or watermarking,
  while retaining pypdf's decompression-bomb limit for array-based streams.

### Fixed

- Pin CI bootstrap tools and correct the local assistant setup and packaging commands.

## [0.2.2] - 2026-08-11

### Security

- Parse user-supplied Office XML parts with `defusedxml` to reject unsafe entity
  declarations before relationship and media analysis.
- Public-binary evidence now validates SBOM/native-inventory/malware report
  schemas, binds the SBOM and native inventory to the packaged artifact hash, and
  checks the contents of the bundled FFmpeg corresponding-source archive; empty
  or placeholder sidecars fail closed.
- Project-specific confidential phrases can be kept in a Git-ignored local
  denylist, and public-safety failures no longer echo matched private text.

### Added

- Added local, dependency-free release helpers for deterministic native-file
  inventories and fail-closed ClamAV/Windows Defender malware reports. They are
  target-host evidence tools and do not run or publish GitHub Actions.
- Added monthly Dependabot updates for the locked Python graph and SHA-pinned
  GitHub Actions, plus explicit repository rules for AI-assisted changes.

### Fixed

- Renamed the Python distribution to `doc-media-toolkit` to avoid collision with
  an unrelated PyPI project; the `pptx_tools` import package and existing CLI
  commands remain compatible.
- Manual CI and candidate builds now install exact, hashed dependency versions
  exported from `uv.lock` instead of resolving unconstrained newer packages.
- Python source distributions now include the generated DOCX/XLSX fixtures
  required by their packaged tests.

## [0.2.1] - 2026-08-09

### Added

- **Source-pinned bundled FFmpeg 8.1.2**: Release candidates now build a minimal GPL runtime from SHA-256-pinned FFmpeg, x264, and zlib source, retain libx264 plus macOS VideoToolbox or Windows Media Foundation encoding, and emit the matching source archives, build script, configuration, toolchain record, licenses, and hashes beside each platform package.
- **Public repository and license-complete packaging**: Public links now target `roanpy/doc-media-toolkit`; standalone builds collect the project, Python, Qt, asset, direct/transitive Python dependency, PyInstaller, FFmpeg, and optional LibreOffice notices under `licenses/`, with a bilingual licensing guide separating MIT source publication from target-specific binary obligations.
- **Open-source readiness**: The source is now MIT-licensed with bilingual README entry points, contribution/security/conduct policies, issue and pull-request templates, third-party notices, accurate language-support boundaries, and an in-app open-source/privacy help topic. Project metadata and repository guidance now consistently use Doc Media Toolkit as the public product name while retaining the `pptx-tools` package and CLI for compatibility.
- **Formal document image compression**: DOCX/DOCM, PDF, and XLSX/XLSM now route through the shared compression GUI and CLI. Document jobs require an explicit target size and an enabled image profile, retain format-specific structure/layout gates, and include their lazy-loaded backends in standalone builds.
- **Unified desktop workspace**: The four PySide6 workspaces now share one navigation shell, canonical typography, status treatment, and dialog hierarchy. Watermark export prioritizes the document preview, compression the file/result queue, the video library the family/version list, and the image library the image list/preview. Identity-review dialogs retain evidence, while cleanup, restore, health, help, and confirmation surfaces use consistent safe, primary, and destructive actions.
- **Architecture contract**: `docs/ARCHITECTURE.md` now records entry points, dependency direction, data stores, write/recovery protocols, trust boundaries, external runtimes, verification gates, and extension rules.
- **Compression result workspace**: Dynamic compression now keeps a persistent per-file result table visible for input type, original/output size, saving, state, and output path, while verbose logs remain available on demand.
- **Backfill quality tiers**: Library backfill now offers 最佳 (default, unchanged behavior), 高质量 (1080p, CRF 20, 12 Mbps ceiling), and 均衡 (720p, CRF 23, 5 Mbps ceiling) tiers with ceiling semantics — sources already within tier specs embed byte-identical, others transcode to the tier with capped VBR, and missing bitrate probes take the safe transcode direction. The confirm dialog carries a tier selector remembered via `QSettings`, output filenames gain a tier suffix, and transcoded masters register as family hash aliases without adding variants. Backfill transcodes now pass timestamps through (`-fps_mode passthrough`, requires FFmpeg ≥ 5.1) and refuse to encode explicitly empty audio tracks, matching the compressor hardening.

### Fixed

- **Fail-closed binary release gate**: The manual candidate workflow has read-only repository permission and no publishing job; public-binary audits now reject skipped dependency scans, missing SBOM/evidence, unsigned or unnotarized artifacts, Windows one-file packages, missing native/malware reports, and bundled GPL FFmpeg without hashed corresponding source.
- **Public release safety**: Current documentation no longer contains maintainer-specific absolute paths, the public-safety gate covers the prepared tree, and release workflows no longer run automatically on tags.
- **Measured compression controls**: The assessment row now decides between one-line and split layouts from the actual available width and control size hints instead of a fixed window breakpoint. The optional forced-output action is included only when visible, preventing clipped labels and buttons at 880–1180 px while preserving the one-line layout when it fits.
- **Video cleanup confirmation**: A cleanup group with only duration/audio/content mismatches
  no longer looks actionable and then silently does nothing. The dialog explains the lock;
  the three cleanup modes are mutually exclusive, while the separate within-family force
  checkbox can quarantine locked variants only after confirming a verified keep version and
  a second warning.
- **Malformed audio preservation**: Videos with an explicitly zero-duration, one-frame
  audio stream are no longer sent through an encode path that silently drops the track.
  The original media bytes and relationship are preserved, counted as non-compressible in
  target planning, and reported as `unusable_audio_stream_preserved`.
- **VFR frame preservation**: Videos that are not intentionally downsampled now use FFmpeg
  timestamp passthrough, preventing nominal stream FPS from silently dropping VFR frames;
  explicit frame-rate reductions retain their dedicated target-frame validation.
- **Formal target-compression audit**: Target-mode SSIM misses now retain the measured compressed candidate (forced redline and decode/structure/metadata failures still restore the source); downscale plans can no longer be bypassed by a near-original copy, unchanged capacity retries are skipped, frame-count validation tolerates one container metadata frame, GUI progress stays monotonic across retries, and manual video thresholds persist across preset changes and sessions.
- **Compact desktop typography and native previews**: Main pages, help, dialogs, tables, and controls use the canonical 18/16/13/12/11 px role scale. Watermark preview can show two consecutive pages vertically without exposing the platform's light scroll-area background, scrollbars share one dark style, status content no longer gets pushed to the bottom by an empty spacer, and video playback runs inside the detail drawer through QtMultimedia instead of opening the system player.
- **Design-spec consistency**: The video detail poster grows toward a responsive 16:9 frame without overlapping metadata, narrow threshold fields remain readable at 1280×800, the video-library toolbar stays on one line when space permits and wraps safely on smaller windows, and the external UI handoff covers all four workspaces with the correct minimum baseline.
- **Recoverable media cleanup**: Video and image quarantine now records a `moving` intent before the file move, verifies hashes during interrupted-move recovery, stores new project-relative paths, restricts original/quarantine paths to managed roots, and removes only the current operation's entries during rollback.
- **AI client trust boundary**: OpenAI-compatible Base URLs now require safe HTTP(S) syntax without embedded credentials, queries, or fragments; response bodies are bounded and retry waits remain cancellable.
- **Repository quality gates**: Development setup installs lint/format tooling, manual CI and release workflows enforce full-repository formatting, and PPTX media extraction streams large ZIP members instead of loading each file wholly into memory.
- **Responsive review previews**: Video match and restore evidence now fills each comparison panel at the available width without stretching or cropping widescreen frames. Watermark page navigation shares the preview toolbar and hides unavailable single-page controls, while the video library keeps only frequent review actions visible and moves secondary maintenance commands into a state-aware menu.
- **Design-aligned video workflow**: The video library now uses removable PPTX chips, an expandable list for additional selections, a collapsed workflow-settings menu, compact library switching, and a responsive click-to-enlarge detail poster without changing any archive, matching, restore, or association rules.

## [0.2.0] - 2026-07-23

### Added

- **PPTX restore review plan**: High-quality PPTX restore now pauses before writing and shows every embedded video in one checklist, including exact/content matches, already-high-quality media, unresolved items, occurrence counts, and the planned action. Selecting an item loads cached 10%/50%/90% cover frames for the current media and chosen family source, with full-video playback and searchable family selection. Each item can remain unchanged, use a one-off override, or explicitly learn a same-video hash only after the output package passes validation.
- **Assisted video matching**: Videos can be dropped onto the library list. Ambiguous external imports and unmatched PPTX media now open a side-by-side review with cached cover frames, playback, ranked duration/frame/audio evidence, explicit family selection, optional hash remembering, and a safe skip path. Unlinked families reuse the same review before manual merge, while unreadable unused non-source versions can be moved to quarantine without changing PPTX references.
- **Video library review and sorting**: The library view reports unlinked, multi-version, and abnormal families, supports focused review filters with an explicit zero-result state, and sorts from any table header. Sort column and direction persist across launches, while same-name families keep the higher-quality source first.
- **External video import**: Batch-import standalone videos into the library using exact hashes followed by conservative duration/aspect/frame/audio fingerprints. Unique matches become family variants, unmatched files create families, and ambiguous matches are skipped.
- **Safe source-PPTX replacement**: Compression can atomically replace PPTX inputs only after archiving their videos, with image compression and automatic audit disabled. Renaming or moving the PPTX does not affect later video-library matching.
- **Library cleanup (整理视频库)**: Scan the video library for within-family redundant variants and cross-family duplicate families (SHA-256 overlap + duration/frame/audio fingerprint clustering). Each group is evaluated side by side — resolution, bitrate, codec, size, audio track, SSIM against the authoritative source, and match confidence — with a recommended keep choice (the source, or smaller-but-close requiring consistent duration/audio, SSIM above a configurable threshold, and adequate resolution). Close-quality groups can be unified into a single 1080p high-quality version re-encoded from the registered source only. Removals move files into a quarantine `_cleanup/` directory (with an `index.json` snapshot) instead of deleting them; entries can be restored, and the quarantine can only be emptied after hash-alias and family migration checks pass. Candidates with different audio tracks, trimmed durations, or insufficient confidence are flagged and blocked from automatic cleanup.
- **PPTX Video Asset Library (third tab)**: Archive one authoritative source while retaining PPTX hashes, path aliases, and slide/shape anchors without copying PPTX files. The default policy only downscales sources above 1080p, while full-original storage remains selectable. Compression registers compact MP4 hashes as aliases; unregistered re-encodes require a unique conservative duration/aspect/five-frame/audio match before upgrade.
- **CLI `videos` subcommand**: `videos add` archives deduplicated source media and registers PPTX shape associations, `videos upgrade` replaces exact known hashes with PowerPoint-compatible high-quality sources, `videos doctor` audits library integrity, and `videos list` reports video families.
- **Video library health checks**: The GUI and CLI inspect media entities, unique hash ownership, deck sources, slide/shape references, historical outputs, untracked files, and quarantine metadata. Full mode recomputes SHA-256 for every entity and distinguishes content changes from mtime-only drift.
- **Auto audit & optimize (Compression tab)**: After compression, optionally audit quality automatically; PPTX files below the SSIM threshold are re-compressed at a higher tier and re-audited. Standalone images/videos are audited only.
- **Rotating application logs**: All tools write to a shared `RotatingFileHandler` log under the per-OS application data directory (`Doc Media Toolkit/logs/`). The Video Asset Management tab can open the log directory directly.
- **Preference persistence**: Compression remembers whether to archive source videos and the video library remembers the last opened location via `QSettings`.
- **Manual CI**: `ci.yml` runs Ruff and the unittest suite on demand across Ubuntu (Python 3.10/3.13), macOS (3.12), and Windows (3.12). Ordinary pushes do not start remote builds.
- **Release hardening**: Windows FFmpeg download is pinned by version and verified by SHA-256; packaged builds smoke-test the GUI offscreen and verify bundled ffmpeg/ffprobe; release assets ship with SHA256SUMS.
- **End-to-end tests**: Real ffmpeg-generated videos run through the full extract → detach → restore chain, asserting slide XML (animation timeline) preservation, placeholder duration, and byte-exact restores.

### Fixed

- **Atomic library registration**: Every user-facing PPTX archive path now stores media and registers the PPTX/shape map as one recoverable workflow. Registration failure rolls back newly added manifest entities and files instead of leaving an unexplained unlinked family.
- **Manifest fail-closed validation**: The complete next manifest is validated before replacing the last known-good copy. Unreadable current manifests cannot be overwritten directly, and a hash that belongs to multiple families blocks automatic matching until the conflict is reviewed.
- **Product naming and guidance**: The Chinese product name is now `文档媒体工具箱`, and the third workspace is consistently described as `PPTX 视频资产库` across the shell, standalone windows, help, README, handoff, and packaging metadata.
- **Reviewable video-family maintenance**: Selected families now expose a direct `核实版本` action instead of relying on the global cleanup command. Ambiguous-match previews use labeled, compact 10%/50%/90% comparison panels and hide stale container suffixes from historical family names. Manual family merges disclose affected variants, hash aliases, PPTX files, and media references; the core rejects unconfirmed low-confidence merges while preserving and validating every deck association.
- **Compact workflow controls**: The video-library PPTX picker now exposes its existing batch selection, long library paths and PPTX selections no longer crowd actions, maintenance buttons use a compact toolbar treatment, and the compression header no longer clips automatic-audit controls at narrower widths.
- **macOS bundle metadata**: Local and release `.app` bundles now use the package version and a valid reverse-domain bundle identifier instead of PyInstaller's `0.0.0` and display-name defaults, then re-sign and verify the updated bundle before DMG creation.
- **Large PPTX high-quality restore**: Multiple shapes that resolve to the same library master now share one output media part, preventing repeated embeddings from pushing presentations past PowerPoint's practical package limits and triggering repair dialogs.
- **Cleanup identity continuity**: Quarantining a low-quality variant retains its SHA-256 as a family alias, and cross-family fingerprint matches with audio require decoded-audio correlation before merge. Batch delivery only prepares variants for changed/missing PPTX outputs and restores an exact quality-approved variant from quarantine when needed.
- **WMV migration and registered-deck rebinding**: Library upgrades can limit replacement to incompatible embedded formats, preserve compatible MP4 media byte-for-byte, and adopt a validated in-place PPTX upgrade by matching slide/shape anchors. WMV-to-MP4 part changes update deck hashes, paths, aliases, structure anchors, and original-version references without breaking one-video-to-many-PPTX associations. Sparse variable-frame-rate WMV transcodes disable H.264 B-frames so the MP4 container duration includes the final presentation timestamp instead of truncating playback.
- **Compression association continuity**: Video-library compression now registers or aliases the input deck before encoding, records non-overwrite outputs, and refreshes deck hashes and slide/shape anchors after safe in-place overwrite. Renamed source paths are accepted through `source_aliases`.
- **Explicit source management**: The video tree marks non-authoritative versions as candidates and shows PPTX reference counts. CLI/Agent workflows can import a video into a verified family with `import-video` and explicitly promote a reviewed version with `set-source`.
- **Same-resolution source candidates**: PPTX imports retain a content-matched candidate when its resolution is unchanged but bitrate is materially higher; candidates are never promoted without review.
- **WMV source delivery**: PPTX video upgrades now transcode authoritative WMV and other incompatible source formats to real H.264/AAC MP4 parts instead of incorrectly treating them as already compatible. Slide XML and video playback timelines remain unchanged while relationships and content types are updated.
- **Duplicate copy suffixes**: Library cleanup treats common copy suffixes such as `_1`, `(1)`, `copy`, and `副本` as naming hints, then still requires duration, aspect, frame, decoded-audio, and SSIM verification before merging families.
- **Video library filenames and truncated imports**: Managed video filenames now consistently include available resolution, duration, and an eight-character SHA-256 suffix. An undecodable embedded video is only associated with an existing same-name family when its complete bytes exactly match the prefix of one unique, larger healthy source, preventing partial cloud files from becoming duplicate authoritative sources.
- **Video names containing dots**: Renaming a family or variant only strips recognized video extensions, so names such as `5.2追溯` are no longer truncated to `5`.
- **Duplicate video selection and audio safety**: Cleanup now defaults to the healthy version with the highest resolution, then bitrate and size. Difficult same-name re-encodes use decoded-audio correlation instead of trusting a coarse spectrum hash, so true duplicates merge while visually identical videos with different soundtracks remain separate for exact PPTX restoration.
- **Damaged and ultrawide duplicate videos**: Cleanup detects truncated or partially corrupted same-name copies by combining full packet validation with aligned-byte evidence, prevents damaged candidates from becoming the retained source, and recognizes same-content ultrawide encodes with small crop/aspect differences while still requiring matching duration, sampled frames, and audio.
- **Renamed and re-encoded deck matching**: Identical PPTX copies now retain every source path as an alias, and direct deck registration falls back from known media hashes to conservative content fingerprints so externally re-encoded videos do not create duplicate families.
- **Deduplicated deck references**: Cleanup now migrates every PPTX asset from a removed video version to the retained version, recognizes common encoder cropping, computes missing family fingerprints, uses non-generic shared names as supporting evidence, and splits transitive matches into directly verified groups that the GUI can safely apply.
- **Manual version identity**: Adding a version to a selected family now rejects videos whose content fingerprint does not match, preventing an unrelated file from replacing the family's source and corrupting future PPTX matching.
- **Safe source selection**: A higher-resolution matched import is retained as a candidate instead of automatically replacing the authoritative source; users explicitly select it after reviewing quality. Successful fingerprint-based PPTX upgrades persist the compressed media hash, so later use resolves by exact alias without repeating content analysis.
- **Library cleanup safety**: Cleanup now validates family identity again when applying a plan, skips unsafe groups as a whole, keeps overlapping cross/within-family work in separate scans, commits each group with one manifest save, quarantines exact duplicate files dropped by family merges, isolates radio-button choices per group, and fails closed for corrupt or path-escaping quarantine indexes.
- **Process lifecycle**: ffmpeg/soffice child processes are registered in a shared process registry; compactor cancel and application quit terminate POSIX process groups with SIGTERM → SIGKILL escalation, while Windows cleanup uses `taskkill /T /F`, replacing `QThread.terminate()` which could orphan subprocesses and abort on exit.
- **Tab switching**: The shell no longer destroys and rebuilds tool windows on every tab switch; per-tab settings and logs are preserved.
- **Duplicate video storage**: Exact duplicate imports are rejected across video families, family merges retain all known compact-video hash aliases, and compressed aliases do not create low-quality physical copies in the library.
- **Re-encoded video identity**: Conservative matching now includes bounded duration tolerance and an audio-spectrum fingerprint. Different resolutions/encodes can share one family, while different audio tracks or edited lengths are not automatically merged or upgraded.
- **Video-family merge safety**: Merging families no longer marks the removed family's source as already high quality, and manifest save failures roll back the in-memory merge.
- **Video workflow UX**: Source archiving is opt-in on first use, selected PPTX names are visible and replaceable, stale selections are cleared after import, and no-match upgrades report that no output was created.
- **Video library organization**: Source imports can target nested relative categories. Family rename updates both the display name and authoritative source filename while preserving hashes and matching identity.
- **Silent unwatermarked videos**: Watermarking a video whose dimensions cannot be probed now fails loudly instead of exporting a 1×1-px overlay with no visible watermark.
- **Preview race**: Preview source folders used by an in-flight preview worker are no longer deleted mid-render; cleanup is deferred until the worker finishes.
- **Absolute relationship targets**: PPTX parts referenced by absolute pack paths (`/ppt/media/...`) are normalized so such decks no longer drop slides or crash with `KeyError`.
- **Audit robustness**: Quality-audit asset extraction rejects paths escaping the audit directory, and audit subprocesses are cancel-safe via the process registry.
- **Packaging validation**: `--dmg` argument errors are rejected during argument normalization instead of after a full PyInstaller build; transparent-PNG alpha channels are preserved during image compression; `pptx_quality_audit` is included in packaged wheels along with `config/` and `assets/` data files.

### Changed

- **Video metadata display**: The video library now shows resolution and duration in separate sortable columns instead of combining them into one specification field.
- **Version single-sourcing**: `pyproject.toml` reads the version dynamically from `pptx_tools.__version__`; the watermark package re-exports it.
- **PPTX output writing**: Compression/detach outputs are written via a temp-file + atomic replace path with streamed (chunked) zip copies, avoiding whole-file reads and partial-write corruption.
- **Release naming**: Build artifacts and workflows are unified under the "Doc Media Toolkit" name.

## [0.1.0] - 2026-05

### Added

- Initial unified toolkit: tabbed desktop GUI and CLI combining DOCX/PDF/PPTX watermark export and PPTX embedded-media compression.

### Fixed

- **macOS GUI**: Fixed a widespread issue where popup dialogs (`StyledDialog` and `HelpDialog`) would initially appear in the top-left corner due to macOS native window cascading behavior. Applied a robust opacity-toggle technique (`setWindowOpacity`) combined with a zero-delay `QTimer` in `showEvent` to guarantee perfect and flicker-free centering over the parent window across all sub-tools.
- **GUI Layout**: Fixed missing `layout.setSpacing(12)` duplicate line issue.
