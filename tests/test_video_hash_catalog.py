from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from pptx_tools.video_manager import VideoProject


class VideoHashCatalogTests(unittest.TestCase):
    def test_merge_requires_one_local_hash_anchor(self) -> None:
        local_hash = "a" * 64
        remote_alias = "b" * 64
        project = VideoProject(
            Path(tempfile.gettempdir()),
            {
                "project_id": "local",
                "families": [
                    {
                        "id": "family-1",
                        "name": "示例",
                        "known_hashes": [local_hash],
                        "source_hashes": [local_hash],
                        "variants": [{"sha256": local_hash}],
                    }
                ],
                "decks": [],
            },
        )
        project.record = Mock()
        project.save = Mock()

        report = project.merge_hash_catalog(
            {
                "format": "doc-media-video-hash-catalog",
                "version": 1,
                "families": [
                    {"name": "另一台机器", "hashes": [local_hash, remote_alias]},
                    {"name": "无锚点", "hashes": ["c" * 64]},
                ],
            }
        )

        self.assertEqual(
            report, {"matched": 1, "added": 1, "skipped": 1, "conflicts": 0}
        )
        self.assertIn(remote_alias, project.families()[0]["known_hashes"])
        project.save.assert_called_once()


if __name__ == "__main__":
    unittest.main()
