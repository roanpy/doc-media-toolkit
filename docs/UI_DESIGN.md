# Doc Media Toolkit UI Design

## 1. Purpose

This document is the canonical UI specification for the four desktop tools:

- Document and media watermark export
- Document and media dynamic compression
- PPTX video asset library
- Document image asset library

The design may reorganize controls, reduce persistent chrome, and improve
review flows. It must not change media processing, matching, archive,
association, restore, cleanup, validation, or overwrite rules.

The code, `README.md`, `docs/HANDOFF.md`, and this document are authoritative.
Public screenshots must come from the running application and use only
synthetic files, generic labels, invalid example endpoints, and temporary
paths. The current verified captures are `images/doc-media-toolkit-compression-en.png`,
`images/doc-media-toolkit-watermark-en.png`, and
`images/doc-media-toolkit-settings-en.png`.

Empty, loading, running, completed, cancelled, disabled, and error states
follow the state rules in this document and do not require separate full-page
screenshots.

## 2. Design Principles

1. Keep the current dark navy and orange identity.
2. Give the primary work object the most space:
   - preview for watermark export;
   - file/result list for compression;
   - video family/version list for the video library.
   - image list and preview for the image library.
3. Keep frequent actions visible and move low-frequency actions into
   disclosure menus or expandable sections.
4. Use spacing, typography, indentation, and separators before adding cards.
5. Never reduce important text to fit more controls.
6. Describe consequences before destructive or identity-changing actions.
7. Preserve native file dialogs and native keyboard behavior.
8. Keep a one-line status and log shelf visible at the bottom. Hovering or
   clicking it opens the detailed log over the lower workspace; running work
   may open it briefly. The closed drawer must never reserve an empty panel.

## 3. Visual Tokens

### 3.1 Colors

| Token | Value | Usage |
| --- | --- | --- |
| `window` | `#0b1017` | Main window and dialog background |
| `surface` | `#0f1720` | Lists, previews, and primary work surfaces |
| `surface-raised` | `#111827` | Selected groups and raised sections |
| `control` | `#18212d` | Buttons, fields, and secondary controls |
| `border` | `#273244` | Normal dividers and panel boundaries |
| `border-strong` | `#334155` | Focused or emphasized control boundaries |
| `text-primary` | `#f8fafc` | Titles and primary values |
| `text-secondary` | `#cbd5e1` | Body text and control values |
| `text-muted` | `#94a3b8` | Help text, paths, and metadata |
| `text-disabled` | `#64748b` | Disabled controls only |
| `accent` | `#f97316` | Primary actions and active navigation |
| `accent-hover` | `#ea580c` | Primary hover state |
| `accent-border` | `#fb923c` | Primary focus and border |
| `selection` | `#12385f` | Selected list rows |
| `success` | `#22c55e` | Completed and safe states |
| `warning` | `#f59e0b` | Review required and non-blocking warnings |
| `danger` | `#ef4444` | Corruption and irreversible actions |
| `info` | `#3b82f6` | Informational and active utility states |

Do not use orange for warnings or destructive actions. Orange identifies the
product's primary action. Warnings use amber and destructive actions use red.

### 3.2 Typography

Use the existing platform-aware UI font configuration:

- macOS: PingFang SC
- Windows: Microsoft YaHei
- Linux: Noto Sans CJK SC

No screen may introduce a private type scale.

| Role | Size | Weight | Line height |
| --- | --- | --- | --- |
| Shared shell title | 18 px | 600 | 26 px |
| Page/dialog title | 16 px | 600 | 22 px |
| Section title | 13 px | 600 | 18 px |
| Brand eyebrow | 11 px | 700 | 16 px |
| Main controls | 12 px | 500-600 | 18 px |
| Main and dialog buttons | 13 px | 500-600 | 18 px |
| Table header | 11 px | 600 | 16 px |
| Video family/body row | 12 px | 500-600 | 18 px |
| Helper text | 11 px | 400 | 16 px |
| Footer/status text | 11 px | 400-600 | 16 px |

Repeated labels and controls use 12 px on every page and dialog; button labels
use 13 px. Only helper text, metadata, table headers, logs, badges, and the
bottom status bar use 11 px. Page-specific styles must not introduce another
private type scale.

Qt stylesheets and custom-painted text use pixel sizes from this table
(`font-size: Npx` or `QFont.setPixelSize(N)`). Do not translate these values
into unrelated point sizes. Generated mockup text is illustrative; when its
apparent size differs from this table, this table wins.

### 3.3 Spacing and Geometry

- Spacing scale: `4, 8, 12, 16, 20, 24, 32`
- Main content margin: 10-12 px at the 960 px reference window
- Dialog content margin: 20-24 px
- Normal control height: 30 px
- Compact toolbar control height: 28 px
- Primary button height: 30 px on main pages, 40 px in confirmations
- Family row height: 50-52 px
- Child/version row height: 40-42 px
- Normal table row height: 46-48 px
- Control radius: 8 px
- Panel/dialog radius: 10 px
- Borders: 1 px
- Active tab underline: 2 px

### 3.4 Effects and Control States

- Main surfaces are flat. A subtle window/header gradient is optional, but
  controls do not use glossy gradients.
- Normal controls use a 1 px border. Keyboard focus uses a visible 2 px
  `info` or `accent-border` outline without changing layout.
- Hover lightens the control surface one step; pressed darkens it one step.
- Selection uses `selection`; primary action uses `accent`; neither uses glow.
- Shadows are reserved for modal dialogs and the overlay detail drawer:
  24 px blur, low-opacity black, no colored shadow.
- Disabled controls keep readable text and never resemble selected controls.
- Motion is limited to 160-220 ms tab underline, drawer, disclosure, and
  opacity transitions. Processing progress is not decorative animation.

## 4. Shared Application Shell

All four main pages use the same header.

### 4.1 Main Header

- Left: current tool title and one-line subtitle.
- Center: equal-width text tabs:
  `水印导出 / 动态压缩 / 视频库 / 图片库`.
- Right: circular help button.
- Active tab: orange text and a 2 px underline.
- Inactive tabs: muted text.
- No filled active tab.
- No segmented-control pill around the tabs.
- No page-specific header height or title size.

### 4.2 Bottom Status Bar

- Height: 24-26 px.
- Left: output/library target and concise operation state.
- Right: `日志` disclosure.
- The expanded log may overlay or expand upward, but the closed state must not
  reserve a large empty panel.
- `操作记录` shows the current process session only. `日志目录` opens the
  persistent shared log directory; neither is stored as detailed history
  inside a video library.

### 4.3 Common Interaction States

Every interactive control must define:

- normal;
- hover;
- keyboard focus;
- pressed;
- selected;
- disabled;
- loading/running when applicable;
- validation error when applicable.

Focus must remain visible on macOS and Windows. Color is not the sole status
indicator; use text and, where useful, a small icon.

### 4.4 State and Feedback Contract

- Empty state: explain what can be added, accepted file types, and both button
  and drag/drop entry points.
- Drag hover: highlight the valid drop target and reject unsupported types
  without changing the current selection.
- Running state: disable actions that can invalidate the operation, keep the
  stop action visible, and show the current object and overall progress.
- Cancelled state: say what completed, what was not changed, and where any
  valid output was saved.
- Error state: preserve the queue/selection, show a concise cause, and keep a
  retry or safe recovery path when the existing logic supports one.
- Completed state: show output location and the next relevant action.
- Long operations must not block window repainting or hide cancellation.
- No success message may appear before output and association validation pass.

### 4.5 Keyboard and Accessibility

- Icon-only controls have an accessible name and tooltip.
- Tab order follows the visual workflow; hidden drawers and collapsed sections
  are removed from the focus chain.
- `Enter` activates the safe primary action, `Esc` closes a drawer or cancels
  a dialog, and `Delete` removes only the current removable queue item.
- Drag/drop always has an equivalent button path.
- Status never depends on color alone; labels and icons carry the same meaning.
- Minimum pointer target is 36 px and the visible keyboard focus ring is at
  least 2 px.

## 5. Watermark Export

### 5.1 Layout

- Left: file queue, 310 px at the reference 1440 px workspace. It may shrink
  only when required to preserve usable preview and settings columns, and must
  always keep the file name, type/size metadata, selection state, and status
  readable.
- Center: document preview, flexible and largest.
- Right: export and watermark settings, approximately 500-560 px.
- Bottom: minimal status bar.

The center shows the current page and, when available, the next page in one
vertical scroll area. Landscape PPTX pages display as a two-page vertical pair
that fills the available preview width; the second page is the page immediately
after the selected page. Portrait/A4 documents display one large page. A
horizontal strip below shows up to five page thumbnails and scrolls for the
remaining pages. Clicking a thumbnail starts the visible pair at that page,
while previous/next moves by two landscape pages or one portrait page.
Single-page media hides the strip and never shows an empty second page. The
page counter states the visible range and total page count.

The bottom status and log shelf is always visible. Hovering or clicking it
opens progress, output path, current file, and the detailed log as an overlay;
it must not shrink the preview or reserve a panel in the settings column.

### 5.2 Required Functions

The redesigned page must retain:

- multi-file add, remove, drag/drop, and selection;
- DOCX, PDF, PPTX, image, and video inputs;
- output format, mode, quality, and PPTX video re-embed;
- text/image watermark selection;
- watermark template, color, opacity, size, spacing, and angle;
- font inspection/fill;
- runtime dependency status;
- background selection, page navigation, refresh, original/preview comparison;
- progress, output path, running state, stop state, and logs.

Font and dependency controls belong in `高级设置`, collapsed by default.
Progress stays visible; verbose logs do not.

Mixed queues retain the current constraints:

- DOCX and PDF always export as PDF.
- PPTX follows the selected PDF/PPTX output format.
- Standalone images and videos use their supported media output path.
- Video reinsertion is enabled only for image-based PPTX output.
- Selection changes update the visible constraint hint; controls are disabled
  rather than silently ignored.

## 6. Dynamic Compression

### 6.1 Layout

- Left: file queue, approximately 310-330 px.
- Right top: compression settings in two or three readable rows.
- Right center: progress and result table.
- Bottom: minimal status bar.
- The assessment controls and actions share one row only when their measured
  size hints fit the available settings width. Otherwise actions move to the
  title row. The optional forced-output action participates in this calculation
  only while visible; no fixed window-width breakpoint may clip control text.
- Interactive controls use their style-derived size hint as a minimum height;
  embedded workspaces re-apply this guard after the shell stylesheet is copied.
- Embedded workspaces keep a 6 px top inset so the first panel's border remains
  visible below the shared shell/tab pane on macOS and Windows.
  The results table starts at least 8 px below the assessment row so zoom and
  platform style metrics cannot make its header collide with the controls.
- The dynamic-compression assessment row uses 16 px gaps; in a wide workspace
  the complete evaluation/action group is right-aligned, while compact layouts
  keep the primary controls left-aligned and move actions to the title row.
- The ordinary encoder selector uses the short labels `自动硬件`、`仅 CPU`、
  `优先 GPU` (and their concise English equivalents); the redundant `普通模式：`
  prefix is omitted.

### 6.2 Required Functions

The redesigned page must retain:

- multi-file add, remove, drag/drop, and selection;
- PPTX, image, and video inputs;
- archive modes: off, 1080p, and original;
- separate video/image SSIM thresholds;
- target-size GPU (off by default) and ordinary preset hardware strategy;
- automatic assessment and incremental upgrade;
- target size, video profile, and image profile;
- active video library and category;
- safe original-PPTX overwrite option;
- current file, overall progress, output location, and per-file result;
- manual quality assessment and incremental optimization;
- completion, failure, stop, and retry states;
- logs.

The bottom status and log shelf is always visible. Hovering or clicking it
opens progress, output path, current file, and the detailed log as an overlay;
the results table keeps its full height while the drawer is closed.

Source-video archiving is optional and off by default; ordinary compression
must not require a video library. Selecting 1080p or original enables the
library and category controls. `覆盖原 PPTX` remains constrained by the
existing validation rules: it requires archiving, video-only compression, and
automatic assessment to be off. The UI may explain those rules but must not
relax them.

The compression page accepts PPTX, DOCX/DOCM, PDF, XLSX/XLSM, image, and video
inputs in one queue. Document backends require an explicit target size and an
enabled image profile, and only replace embedded images after their
format-specific structure/layout gates pass. Automatic assessment may re-grade
and optimize PPTX output; PDF bypasses the ZIP-media audit, while standalone
images and videos are assessed only. Manual incremental optimization remains a
PPTX-only action and must explain why it is disabled for other inputs.

Standalone image/video handling is ancillary: the same add/drop queue, target,
presets, thresholds, progress, and report UI are reused; there is no separate
media-compressor page or professional codec panel. Safe target output that is
still above target may reveal `Try forced`; the button is otherwise hidden.
Clicking it must show a second confirmation, preserve the safe output, name the
forced output separately, state the absolute quality redlines, and warn that the
target can still be impossible.

Experimental builds show a persistent banner and `Experimental` application
name. Their settings, logs, temporary paths, bundle identifier, build output,
and `_experimental` user outputs must remain isolated from the formal app.

## 7. PPTX Video Asset Library

### 7.1 Main Layout

- Compact library path row.
- Compact PPTX workflow row.
- Search, status filters, and only the most frequent actions.
- Full-width family/version table.
- On-demand overlay detail drawer.
- Minimal status bar.

The library path row shows the active library name and path. Its menu retains
new library, switch/open library, health check, operation record, and log
directory access. Recently opened libraries, category, external-video
directory, and table sorting persist through the existing `QSettings`
behavior. Each library remains an independent directory with its own manifest,
media, cleanup, reports, and backups.

### 7.2 Multiple PPTX Selection

- The workflow area accepts multiple PPTX files by button or drag/drop.
- Show the first two selected PPTX files as removable chips.
- Show additional files as `还有 N 个`.
- Clicking `还有 N 个` opens a lightweight popover list.
- The popover supports removing one file and clearing the selection.
- Do not open a blocking modal merely to list selected files.
- Archive and HD restore operate on the explicit current selection. If a
  multi-file action requires per-file output choices, use the existing batch
  directory flow rather than repeated modal prompts.

### 7.3 Video Table

Family and version rows must be visually distinct:

- family row uses a tinted surface, stronger type, and collapse arrow;
- child versions are indented;
- a subtle hierarchy line connects child versions;
- leave 6 px between family groups;
- do not use global zebra striping that hides family boundaries.

Important columns remain visible:

- video family/version;
- resolution;
- duration;
- size;
- hash count;
- association/status.

The file location is lower priority and may be covered by the detail drawer.

### 7.4 Filters and Actions

Persistent controls:

- search;
- all;
- review required;
- no PPTX association;
- multiple versions;
- file anomaly;
- play;
- review association;
- more actions.

`更多操作` contains rename, move, add version, set HD source, merge, find
missing, review versions, quarantine anomaly, organize, pending cleanup, and
export association records.

At the reference desktop width, search, status-filter tags, and the frequent
actions share one toolbar row. Only review version, add version, organize
library, and More Actions remain visible; playback is available by double-click
or in the detail drawer, while low-frequency maintenance stays in More Actions.
The status filters use compact tag styling; they are not a second combo box.
Responsive fallback may wrap the complete action group, but never individual
labels or the association/status column.

External videos may be imported through the workflow button or dropped onto
the video table. Exact matches reuse the existing entity; unique content
matches become candidates in the matched family; ambiguous matches open the
review dialog; unsupported files are rejected without modifying the library.

Action availability follows the selected object:

- family: play current HD source, rename, move current source, add version,
  review versions, merge, and inspect association;
- version: play, move, set as HD source when valid, and inspect metadata;
- unreadable non-source, non-active, unreferenced version: quarantine anomaly;
- no selection or invalid selection: disable identity-changing actions and
  explain the requirement in a tooltip.

No-association and multiple-version states are not presented as corruption.
Only actual file anomalies use the danger color.

The status filters are clickable:

- `待核对` is the de-duplicated union of no-association, multiple-version, and
  file-anomaly families;
- `无关联` means no PPTX references and may be a valid standalone asset;
- `多版本` means the family has candidates and is not itself an error;
- `文件异常` means at least one registered entity is unavailable or unreadable.

All visible table headers support true-value sorting and toggle ascending or
descending order. Search covers display name, path, resolution, and hash
fragment. A zero-result state offers one action to clear search and filters.
Double-clicking a family plays its current HD source; double-clicking a version
plays that version.

### 7.5 Detail Drawer

- Closed by default.
- Single-clicking a family or version opens it.
- Clicking a different row updates it.
- The expand/collapse arrow only expands the family.
- The drawer overlays the right side of the table; it does not resize columns.
- It may cover the file-location column only.
- Close with outside click, `Esc`, collapse, or close button.
- Important left-side columns remain visible.
- The poster container targets 16:9 and shrinks with available height; source
  frames always preserve their native aspect ratio without cropping.
- The poster is a representative frame. Playback starts only through an
  explicit play action and replaces the poster inside the drawer; it must not
  open an external player or another window.

The primary drawer action is contextual:

- no association: review association;
- normal association: inspect association;
- file anomaly: repair or replace when the existing logic permits it.

The drawer width is 420-540 px and overlays only the low-priority right side of
the table. Its metadata uses aligned label/value rows, keeps resolution,
duration, size, references, path, hash, and status readable, and uses a single
representative poster with in-place playback.

### 7.6 Recovered Approved Adjustments

The following adjustments were reconfirmed on 2026-07-24 and override older
mockup details:

1. Shared navigation uses orange text plus a centered short underline, not a
   filled pill. Hover feedback and the underline transition use 180 ms.
2. The watermark file queue is not narrower than the table content at the
   reference workspace. The preview width is reduced before queue text becomes
   unreadable.
3. Landscape PPTX preview shows the selected page and the next page; portrait
   documents show one page. Both use a horizontal thumbnail strip.
4. Watermark and compression use a full-width bottom status and log shelf;
   detailed logs open above it on hover, click, or briefly during active work
   without reserving workspace.
5. Compression controls are grouped by preset, media quality, archive target,
   category, assessment, and overwrite safety. The active video library is a
   selectable persisted target, not a static label.
6. Video-library family and version rows remain visually distinct. The
   association/status column must not wrap at the reference width.
7. The video detail panel is an overlay drawer; opening it must not resize the
   table or permanently reduce list space.
8. Search, filter tags, and frequent actions remain on one row at the reference
   width. Low-frequency maintenance actions remain available from More Actions.
9. The implementation is accepted only after current Qt screenshots are
   compared against these rules at 1440x900 and 1280x800.

## 8. Document Image Asset Library

There is no dedicated canonical mockup for this newer workspace yet. Until one
is approved, the running Qt implementation must follow the shared tokens and
the rules below; it must not copy video-specific controls merely for visual
symmetry.

### 8.1 Main Layout

- Compact current-library row with create and switch/open actions.
- Collapsible import workflow accepting standalone images, PPTX, DOCX, and
  digital PDF files by button or drag/drop.
- Search and health filters above a resizable image list and preview/detail
  pane.
- The list remains the primary work surface. The detail pane shows a
  non-cropped preview, metadata, summary, and recorded origins.
- Low-frequency maintenance and AI suggestions stay under `更多操作`.

### 8.2 Identity and Review

- Exact SHA-256 duplicates reuse the existing asset and add origin records.
- Pixel hash and dHash only produce review candidates; similar images are not
  merged automatically.
- Filters distinguish duplicate origins, similar candidates, undersized
  assets, and missing origins without presenting all of them as corruption.
- `编辑信息` changes only name, category, tags, and summary. `查找相似图`
  presents evidence and requires an explicit keep/merge/ignore decision.
- AI can suggest metadata and merge groups but never applies identity changes,
  removal, or cleanup without confirmation.

### 8.3 Import, Cleanup, and Feedback

- PPTX/DOCX import extracts only relationship-referenced images. PDF import
  extracts embedded raster images and does not render full pages or run OCR.
- The workspace does not promise image reinsertion or document rewriting.
- Removal, merge, and orphan cleanup move files into `_cleanup/`; restore is a
  normal action and permanent empty is destructive with a safe default.
- Running imports and AI requests remain cancellable and keep the UI
  responsive. Completion names the imported/reused/skipped counts.
- Base URL validation, response limits, and API-key lifetime are security
  behavior, not configurable visual options.

## 9. Dialogs

### 9.1 Dialog Header

All custom dialogs use:

- 16 px title;
- optional 11 px subtitle;
- close button at top right where appropriate;
- no main-window navigation inside the dialog;
- no page-specific title scale.

### 9.2 Video Match Review

Retain:

- unresolved source identity;
- 10%/50%/90% source and candidate frames;
- full-video play actions;
- ranked candidate table;
- resolution, duration, frame difference, audio evidence, and decision;
- optional remembered association;
- link, new family, and skip.

The score is supporting evidence, not confirmation. Audio `不同/未知` cannot
be styled as an automatic match.

`关联并继续` writes only after explicit confirmation. `新建视频族` creates a
new identity instead of forcing a weak candidate. `跳过` leaves the source and
all existing associations unchanged.

### 9.3 PPTX HD Restore Review

Retain:

- complete list of embedded media;
- exact/content/unmatched state;
- references and planned action;
- current and target previews;
- target family search;
- keep current;
- replace for this output only;
- confirm same video and remember;
- save as a new PPTX.

Only the remembered option may write a new media hash association, and only
after output validation.

The review includes every embedded video, not only unresolved items. The
default plan keeps already-high-quality media unchanged, preserves uncertain
items, and proposes replacement only when a validated target exists. Cancel
must leave both the PPTX and library unchanged. Output always uses Save As.

### 9.4 Library Cleanup Review

Retain:

- duplicate groups;
- one selected keep candidate per group;
- resolution, duration, size, SSIM, and recommendation;
- safe/unsafe group state;
- keep selected, generate compatible 1080p, or skip;
- apply or cancel.

Unsafe groups default to skip. Cleanup moves files to pending cleanup and
does not immediately delete them.

When families are merged or versions removed, known hashes and every PPTX
reference migrate to a valid retained family/version before the manifest is
saved. A failed consistency check rolls the whole group back.

### 9.5 Pending Cleanup

Retain:

- file, source family, size, reason, and quarantine time;
- single-item restore;
- cleanup issue summary;
- permanent empty only when consistency checks pass.

Restore is a normal secondary action. Permanent empty uses the danger color,
states the count, and requires confirmation.

Restore reinstates both the file and its manifest snapshot. Permanent empty is
disabled when any live reference, invalid index, or unsafe path is detected.

### 9.6 Library Health

Retain:

- quick/full-check mode;
- families, variants, decks, and references;
- grouped errors, warnings, and information;
- full hash verification;
- prune stale output records only;
- save JSON report;
- close.

The dialog is read-only except for explicit pruning of stale output history.
It must not imply that warnings are repaired automatically.

Only red errors block safe restore or cleanup. Warnings and information remain
reviewable without being presented as data loss. Full hash verification is a
long-running, cancellable follow-up and does not modify files.

### 9.7 Help

- Four top-level help sections matching the four tools.
- A local table of contents for the current tool.
- 12 px body text and a readable 18-20 px line height.
- Short steps and safety callouts instead of long uninterrupted paragraphs.
- No invented online account, cloud service, or version number.

### 9.8 Standard Messages

Standard messages use one of four patterns:

| Type | Primary treatment | Default focus |
| --- | --- | --- |
| Information | Neutral button | Close/OK |
| Recoverable warning | Orange primary | Safe continuation |
| Confirmation | Orange primary | Cancel when risk exists |
| Destructive | Red outline/fill | Cancel |

Destructive confirmations must state:

- the object;
- the irreversible consequence;
- what remains protected;
- the safe default action.

Native `QFileDialog` and platform permission dialogs remain native and are
not restyled.

### 9.9 Selection and Confirmation Rules

- Rename, move, set-source, merge, quarantine, cleanup, and overwrite always
  name the affected object before execution.
- Merge confirmation summarizes moved variants, known hashes, PPTX decks, and
  media-reference counts.
- Removing or quarantining a referenced/current/source version is refused;
  the UI does not offer a bypass.
- A weak visual similarity suggestion never becomes an identity association
  without explicit user confirmation.
- Output validation failure leaves the source PPTX and existing associations
  unchanged and exposes the error in the operation record/log.

## 10. Responsive Rules

- Design target: 1440 x 900.
- Unified shell minimum: 760 x 540; individual workspaces target at least
  880-900 x 560-600, while 1280 x 800 remains the primary review baseline.
- At narrower widths:
  - preserve the canonical 12 px control/body and 11 px helper/table text;
  - collapse low-frequency controls into menus;
  - shorten paths with ellipsis and tooltips;
  - keep important columns before file paths;
  - use splitters where the current implementation already supports them;
  - never reduce important text to make one row fit.
- The video detail panel stays an overlay rather than forcing table reflow.

## 11. Implementation Approach

This is a PySide6 application. Do not convert the application to a web
frontend and do not slice the mockups into background images.

Implementation should use native Qt widgets, layouts, models, delegates,
stylesheets, focus handling, and accessibility metadata. The mockups are
measurement and hierarchy references only.

Prefer the smallest shared changes that remove inconsistency:

- one canonical type scale;
- one canonical main header;
- one dialog/button/status stylesheet;
- existing widgets and behavior wherever possible;
- local layout changes inside each current window.

Do not introduce a new UI framework, renderer, design-system package, or
parallel application shell.

Recommended implementation order:

1. Add the shared palette, type scale, button states, and dialog style.
2. Normalize the common header and collapsed status/log bar.
3. Re-layout each existing main window without changing its workers or
   processing calls.
4. Re-layout the video/image-library business dialogs and shared help/messages.
5. Add only the focused interaction changes specified here: clickable status
   filters, overlay detail drawer, compact multi-PPTX chips, and responsive
   action disclosure.
6. Capture and compare the required states before removing any old layout
   code.

## 12. Functional Traceability

| User task | Primary surface | Confirmation or result |
| --- | --- | --- |
| Batch watermark/export | Watermark main | Output path, per-file state, log |
| Batch compress/audit/optimize | Compression main | Result table, output path, log |
| Archive PPTX videos | Video library main | Library families and PPTX references |
| Import external videos | Video library main | Exact/unique result or match review |
| Review ambiguous identity | Video match review | Link, new family, or unchanged skip |
| Restore PPTX HD media | PPTX HD restore review | Validated Save As output |
| Review duplicate variants | Library cleanup review | Pending-cleanup migration or skip |
| Recover quarantined media | Pending cleanup | Restored file and manifest snapshot |
| Diagnose library | Library health | Read-only result or saved JSON report |
| Import document images | Image library main | Imported/reused/skipped result |
| Review similar images | Image similarity review | Merge, keep separate, or unchanged skip |
| Organize image metadata | Image detail/AI review | Explicitly applied metadata only |
| Understand workflow | Help | Local instructions and diagnostics entry |
| Confirm risky action | Standard message | Safe default, explicit consequence |

## 13. Verification

Before accepting the implementation:

1. Capture all four default main pages at the same viewport.
2. Capture populated, running, completed, empty, disabled, and error states.
3. Capture every custom dialog in this document.
4. Compare typography, header, spacing, button priority, and semantic colors.
5. Verify keyboard focus, `Esc`, Enter, double-click, drag/drop, and selection.
6. Verify no important text is clipped at the minimum supported size.
7. Verify macOS and Windows scaling behavior.
8. Run the existing automated tests.
9. Add focused GUI regression checks for changed interaction behavior.
10. Confirm no processing or data-association logic changed unintentionally.

## 14. Implemented Design Decisions

The design screenshots remain the layout baseline. The implementation makes
only these deliberate desktop-runtime adjustments:

- The shared navigation is owned by the unified shell. Standalone child
  windows show their page title but do not duplicate shell navigation.
- A single-page document shows one preview. Multi-page documents show the
  current and next page vertically plus a thumbnail strip; the UI never
  fabricates a second page when one does not exist.
- Video match evidence remains a 10%/50%/90% three-frame contact sheet because
  it is identity evidence rather than a decorative poster. The library detail
  drawer uses one representative frame and opens a larger view on click.
- Real paths, names, counts, disabled states, and processing results replace
  illustrative mockup data.
- The PPTX workflow keeps at most two removable file chips visible. Additional
  files appear in a non-blocking menu, and infrequent archive/restore settings
  remain collapsed under `工作流设置`.
- The library path row keeps only `切换 / 打开视频库` and `更多` visible.
  Creation, health, operation record, and logs reuse their existing actions
  inside the menu.
- The library search, review filters, counts, and frequent actions share one
  toolbar when at least 1680 px is available. At smaller widths the action
  group moves to a second right-aligned row instead of clipping labels.

These adjustments do not change processing, matching, association, cleanup,
validation, overwrite, or Save As behavior.
