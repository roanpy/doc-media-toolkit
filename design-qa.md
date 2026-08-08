# Desktop UI Design QA

final result: passed

## Current revalidation

- Watermark, compression, video library, and image library passed at a
  960x620 compact viewport on 2026-08-01.
- The four workspaces share the same shell, palette, typography, field/button
  geometry, scrollbar treatment, left rail alignment, and bottom log shelf.
- Watermark shows two complete 16:9 pages without stretching or cropping and
  keeps the horizontal thumbnail strip below the preview.
- Compression keeps the opt-in video-library archive flow, quality controls,
  audit actions, result table, and status shelf visible without overlap.
- Video-library core filters and version actions remain available in one row;
  its detail/player drawer stays inside the application window.
- Image-library filters, primary actions, list, preview, source details, and
  collapsed log shelf remain readable without overlap.
- File sizes use the shared KB/MB/GB formatter; raw byte counts are not shown
  in the GUI or CLI.

## Compared surfaces

- Current Qt application surfaces rendered with isolated settings.
- The three English captures under `docs/images/`, created with anonymous
  synthetic files and invalid example configuration.
- `docs/UI_DESIGN.md`, used as the canonical interaction specification.

## Result

- Shared typography, navy/orange palette, button priority, compact status/log
  treatment, main work-surface priority, and dialog hierarchy are consistent.
- Main surfaces use only the canonical 12/13/15/16/17/20/22 px role scale;
  custom empty states use pixel sizes rather than platform-dependent points.
- Watermark uses one responsive current-page preview and a multi-page
  thumbnail strip.
- Compression prioritizes settings, progress, and a per-file result table;
  logs are collapsed.
- Video library prioritizes the family/version list, compact PPTX chips, status
  filters, and an overlay detail drawer.
- The video-library toolbar is one line on wide workspaces and wraps only the
  frequent action group at narrower widths, without truncating filter labels.
- The detail drawer poster uses a responsive 16:9 container while preserving
  the source frame's native aspect ratio.
- Match and restore previews preserve aspect ratio and use explicit playback.
- Image library prioritizes the asset list, preview, source details, and
  recoverable cleanup; AI organization remains a secondary action.

## Intentional differences

- Standalone child windows omit the unified shell navigation.
- Single-page documents never show a fabricated second page.
- Real data and native file/permission dialogs replace illustrative mockup
  content.
- Three-frame contact sheets remain in identity-review dialogs; the library
  drawer uses a single representative poster.

## Verification

- Viewport: 960x620, plus existing geometry checks at larger sizes.
- Isolated automated GUI and media regression suite: 242 tests passed on
  2026-08-01.
- Ruff lint and full-repository format checks, `compileall`, dependency checks,
  public-safety scan, and `git diff --check` passed.
- Library doctor checks and release-package validation must use disposable
  fixtures or private release records; public QA documentation does not expose
  local library names, paths, or package fingerprints.
