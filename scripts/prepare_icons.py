#!/usr/bin/env python3
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QRectF
from PySide6.QtGui import QGuiApplication, QImage, QImageWriter, QPainter
from PySide6.QtSvg import QSvgRenderer


def render_svg(renderer: QSvgRenderer, size: int | tuple[int, int]) -> QImage:
    if isinstance(size, tuple):
        width, height = size
    else:
        width = height = size
    image = QImage(width, height, QImage.Format.Format_ARGB32)
    image.fill(0)
    painter = QPainter(image)
    renderer.render(painter, QRectF(0, 0, width, height))
    painter.end()
    return image


def save_image(image: QImage, path: Path, fmt: bytes | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if fmt is None:
        ok = image.save(str(path))
    else:
        writer = QImageWriter(str(path), fmt)
        ok = writer.write(image)
    if not ok:
        raise SystemExit(f"Failed to write icon: {path}")


def main() -> int:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    _app = QGuiApplication.instance() or QGuiApplication([])

    project_root = Path(__file__).resolve().parents[1]
    assets_dir = project_root / "assets"
    svg_path = assets_dir / "app_icon.svg"
    if not svg_path.exists():
        raise SystemExit(f"Missing SVG icon: {svg_path}")

    generated_paths = (
        assets_dir / "app_icon.png",
        assets_dir / "app_icon.icns",
        assets_dir / "app_icon.ico",
    )
    cover_svg_path = assets_dir / "app_cover.svg"
    if cover_svg_path.exists():
        generated_paths += (assets_dir / "app_cover.png",)
    if all(path.is_file() for path in generated_paths):
        for path in generated_paths:
            print(path)
        return 0

    renderer = QSvgRenderer(str(svg_path))
    if not renderer.isValid():
        raise SystemExit(f"Invalid SVG icon: {svg_path}")

    base_image = render_svg(renderer, 1024)
    save_image(base_image, assets_dir / "app_icon.png", b"png")
    save_image(base_image, assets_dir / "app_icon.icns", b"icns")
    save_image(render_svg(renderer, 256), assets_dir / "app_icon.ico", b"ico")

    if cover_svg_path.exists():
        cover_renderer = QSvgRenderer(str(cover_svg_path))
        if not cover_renderer.isValid():
            raise SystemExit(f"Invalid SVG cover: {cover_svg_path}")
        save_image(
            render_svg(cover_renderer, (1600, 900)),
            assets_dir / "app_cover.png",
            b"png",
        )

    print(assets_dir / "app_icon.png")
    print(assets_dir / "app_icon.icns")
    print(assets_dir / "app_icon.ico")
    if cover_svg_path.exists():
        print(assets_dir / "app_cover.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
