from __future__ import annotations

import unittest
import tempfile
from pathlib import Path

from pptx_tools.media_rules import (
    fit_media_component,
    normalize_import_name,
    normalize_media_category,
    safe_media_name,
)
from pptx_tools.video_manager import _unique_path, _variant_filename


class MediaRulesTests(unittest.TestCase):
    def test_windows_device_names_are_safe_even_with_extensions(self) -> None:
        for value in ("CON", "con.txt", "AUX.tar.gz", "COM1.mp4", "LPT³.png"):
            self.assertFalse(
                safe_media_name(value).split(".", 1)[0].casefold()
                in {
                    "con",
                    "prn",
                    "aux",
                    "nul",
                    "com1",
                    "lpt3",
                    "com¹",
                    "com²",
                    "com³",
                    "lpt¹",
                    "lpt²",
                    "lpt³",
                }
            )

    def test_timestamp_suffix_survives_name_budget(self) -> None:
        result = normalize_import_name("1234567890-" + "设" * 100)
        self.assertTrue(result.endswith("_1234567890"))
        self.assertLessEqual(len(result), 80)

        astral = normalize_import_name("1234567890-" + "𐐷" * 100)
        self.assertTrue(astral.endswith("_1234567890"))
        self.assertLessEqual(len(astral.encode("utf-8")), 255)

    def test_component_budget_keeps_metadata_suffix(self) -> None:
        suffix = "_[1920x1080_12.0s]_deadbeef.mp4"
        result = fit_media_component("画" * 200, suffix)
        self.assertTrue(result.endswith(suffix))
        self.assertLessEqual(len(result.encode("utf-8")), 255)

    def test_empty_fallback_still_rejects_invalid_category_component(self) -> None:
        with self.assertRaises(ValueError):
            normalize_media_category("???")

    def test_astral_unicode_is_bounded_without_surrogate_splitting(self) -> None:
        result = safe_media_name("𐐷" * 100)
        self.assertLessEqual(len(result), 80)
        self.assertLessEqual(len(result.encode("utf-8")), 255)
        result.encode("utf-8").decode("utf-8")
        imported = normalize_import_name("1234567890-" + "𐐷" * 100)
        self.assertTrue(imported.endswith("_1234567890"))
        self.assertLessEqual(len(imported.encode("utf-8")), 255)

    def test_variant_and_collision_names_keep_metadata_suffix_under_budget(
        self,
    ) -> None:
        metadata = {"width": 1920, "height": 1080, "duration_sec": 12.0}
        digest = "deadbeef" * 8
        filename = _variant_filename("画" * 200, metadata, digest, ".mp4")
        suffix = "_[1920x1080_12.0s]_deadbeef.mp4"
        self.assertTrue(filename.endswith(suffix))
        self.assertLessEqual(len(filename.encode("utf-8")), 255)
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / filename
            path.touch()
            collision = _unique_path(path)
            self.assertTrue(collision.name.endswith("_deadbeef_2.mp4"))
            self.assertLessEqual(len(collision.name.encode("utf-8")), 255)


if __name__ == "__main__":
    unittest.main()
