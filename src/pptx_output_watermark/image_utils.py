from __future__ import annotations

from PIL import Image


def normalize_int_setting(value: int | None, default: int, *, minimum: int = 1) -> int:
    try:
        normalized = int(value if value is not None else default)
    except (TypeError, ValueError):
        normalized = int(default)
    return max(minimum, normalized)


def scale_image_to_limits(
    image: Image.Image,
    *,
    max_edge: int,
    max_pixels: int,
) -> Image.Image:
    width, height = image.size
    if width <= 0 or height <= 0:
        return image

    scale = 1.0
    longest_edge = max(width, height)
    if longest_edge > max_edge:
        scale = min(scale, max_edge / float(longest_edge))

    total_pixels = width * height
    if total_pixels > max_pixels:
        scale = min(scale, (max_pixels / float(total_pixels)) ** 0.5)

    if scale >= 0.999:
        return image

    new_width = max(1, int(width * scale))
    new_height = max(1, int(height * scale))
    return image.resize((new_width, new_height), Image.Resampling.LANCZOS)
