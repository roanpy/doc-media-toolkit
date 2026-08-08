# Configuration Guide

[中文说明](README.md)

This directory stores compression strategy presets. The default strategy is:

- larger on-slide videos receive higher default resolution
- full-screen or near-full-screen videos keep more detail and higher audio bitrate
- small videos yield more aggressively to the overall file-size target

Project note: built with a GPT-5.5-assisted engineering workflow.

## Preset Files

- `default.json`
  Compatibility preset; currently the same as `balanced.json`.
- `balanced.json`
  Recommended default. Balances quality and output size.
- `high.json`
  Quality-first preset.
- `aggressive.json`
  Size-first preset.

By default, the tool removes intermediate artifacts and keeps only:

- the original `pptx`
- the compressed `pptx`

Use `--keep-artifacts` to keep the report JSON and compressed media folder. On ordinary processing failures, the tool also copies available diagnostic reports or process videos beside the source file. Manual stops and successful runs clean the temporary folders.

## `render_limits`

- `max_output_height`
  Maximum encoded video height. Default: `1080`.
- `max_long_edge`
  Maximum encoded long edge. Default: `1920`.

## `height_floor_rules`

Rules are evaluated in order. The first matching rule sets the minimum resolution bucket.

- `min_area_ratio`
  Area ratio of the video shape on the slide.
- `min_width_ratio`
  Width ratio of the video shape relative to the slide.
- `min_height_ratio`
  Height ratio of the video shape relative to the slide.
- `min_height`
  Minimum output height bucket when the rule matches.

The tool evaluates the video's placement size in the PPT slide, not the source video's pixel dimensions.

Default intent:

- near-full-screen or large placement: prefer `1080p` when allowed
- width or height greater than half the slide: at least `720p`
- other videos: at least `480p`

## `audio_rules`

Audio limits differ by preset:

- high quality: up to `192k`
- balanced: up to `128k`
- aggressive: up to `96k`

Fields:

- `min_kbps`
  Minimum audio bitrate.
- `max_kbps`
  Maximum audio bitrate.
- `rounding_step_kbps`
  Rounding step for source bitrate caps.
- `tiers`
  On-slide importance tiers for audio bitrate.
- `pressure_adjustments`
  Additional audio reductions under tight target-size pressure.

The tool never raises source audio bitrate.

## `bitrate_ladder`

Only three resolution buckets are used: `480p`, `720p`, and `1080p`.

- `min_video_kbps`
  Quality floor in fixed target-size mode.
- `best_video_kbps`
  Recommended bitrate in preset quality mode, and upper bound when the target-size budget is generous.

Current built-in video bitrate table:

| Preset | Resolution | Minimum | Recommended |
| --- | ---: | ---: | ---: |
| High | 480p | 700k | 1500k |
| High | 720p | 1700k | 3800k |
| High | 1080p | 3200k | 6500k |
| Balanced | 480p | 430k | 980k |
| Balanced | 720p | 980k | 2150k |
| Balanced | 1080p | 1950k | 3900k |
| Aggressive | 480p | 325k | 700k |
| Aggressive | 720p | 700k | 1500k |
| Aggressive | 1080p | 1400k | 2800k |

These values are upper targets and floors. The tool does not increase source resolution, source video bitrate, or source audio bitrate.
FFmpeg is used in variable-bitrate mode here: `-b:v` is the target average bitrate, while `-maxrate` and `-bufsize` bound short-term peaks. CPU two-pass encoding usually lands closer to the average target; GPU encoding is approximate.

Target size is closed against the final output bytes, not only bitrate estimates:
at most two correction encodes are allowed, plus one quality give-back when the
result is below 95% of target. Assets that fail dual-scale SSIM or structural
checks are restored and are not degraded further. Safe base thresholds are 0.95
for video and 0.99 for images. Forced output requires a second GUI confirmation
and still observes absolute redlines. Target mode defaults to CPU two-pass;
target-size GPU is off by default. Windows probes NVENC, QSV, AMF/MF and falls
back to CPU per asset.

## Frame Rate

- High: keep source frame rate.
- Balanced: cap high-FPS sources at `30fps`.
- Aggressive: cap high-FPS sources at `24fps`.
- Low-FPS source videos are not upsampled.

## Tuning Tips

- To make large videos clearer, raise large-placement `min_height` rules or the `720p/1080p` bitrate bounds.
- To compress small videos harder, lower the `480p` bitrate bounds.
- To preserve more audio quality, raise `audio_rules.tiers`.
- To hit target sizes more aggressively, tune `pressure_adjustments`.
