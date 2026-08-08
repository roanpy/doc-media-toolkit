from __future__ import annotations

import sys
from collections.abc import Sequence


HELP = """Doc Media Toolkit CLI

Usage:
  pptx-tools watermark [args...]   Export document/media files with optional watermark.
  pptx-tools compact [args...]     Compress PPTX embedded media or standalone media files.
  pptx-tools videos [args...]      Deduplicate PPTX video sources and upgrade embedded videos.
  pptx-tools images [args...]      Inspect a deduplicated document image library.

Examples:
  pptx-tools watermark --help
  pptx-tools compact --help
  pptx-tools videos --help
  pptx-tools images --help
"""


def _forward(argv: Sequence[str], module_name: str, main_name: str = "main") -> int:
    original_argv = sys.argv[:]
    sys.argv = [f"pptx-tools {argv[0]}", *argv[1:]]
    try:
        module = __import__(module_name, fromlist=[main_name])
        return int(getattr(module, main_name)())
    finally:
        sys.argv = original_argv


def main(argv: Sequence[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in {"-h", "--help", "help"}:
        print(HELP)
        return 0

    command = args[0].strip().lower()
    if command in {"watermark", "export", "wm"}:
        return _forward(args, "pptx_output_watermark.cli")
    if command in {"compact", "compress", "video", "media"}:
        return _forward(args, "pptx_video_compactor")
    if command in {"videos", "library", "assets"}:
        return _forward(args, "pptx_tools.video_manager")
    if command in {"images", "pictures"}:
        return _forward(args, "pptx_tools.image_manager")

    print(f"Unknown subcommand: {args[0]}\n")
    print(HELP)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
