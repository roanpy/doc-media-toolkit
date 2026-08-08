from __future__ import annotations

import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from tests.test_video_manager import make_video_pptx, no_probe
from pptx_tools.video_library_health import (
    audit_video_project,
    prune_missing_output_records,
)
from pptx_tools.video_manager import VideoProject, main


class VideoLibraryHealthTest(unittest.TestCase):
    def test_archive_and_register_pptx_keeps_media_and_shape_associations_together(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            make_video_pptx(source, b"video-source", "Source")
            project = VideoProject.create(root / "library")

            with patch("pptx_tools.video_manager.probe_video", side_effect=no_probe):
                result = project.archive_and_register_pptx(source)

            self.assertEqual(result["added"], 1)
            self.assertEqual(len(project.families()), 1)
            self.assertEqual(len(project.decks()), 1)
            self.assertEqual(len(result["deck"]["assets"]), 1)
            self.assertEqual(
                result["deck"]["assets"][0]["family_id"],
                project.families()[0]["id"],
            )

    def test_archive_and_register_rolls_back_new_media_when_deck_registration_fails(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            make_video_pptx(source, b"video-source", "Source")
            project = VideoProject.create(root / "library")

            with (
                patch("pptx_tools.video_manager.probe_video", side_effect=no_probe),
                patch.object(project, "add_deck", side_effect=RuntimeError("boom")),
            ):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    project.archive_and_register_pptx(source)

            reopened = VideoProject.open(root / "library")
            self.assertEqual(reopened.families(), [])
            self.assertEqual(reopened.decks(), [])
            self.assertEqual(
                [
                    path
                    for path in (root / "library" / "media").rglob("*")
                    if path.is_file()
                ],
                [],
            )

            reopened.manifest_path.write_text("{broken", encoding="utf-8")
            recovered = VideoProject.open(root / "library")
            self.assertTrue(recovered.recovered_from_backup)
            self.assertEqual(recovered.families(), [])
            self.assertEqual(recovered.decks(), [])

    def test_save_validates_next_manifest_before_replacing_current_copy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            make_video_pptx(source, b"video-source", "Source")
            project = VideoProject.create(root / "library")
            with patch("pptx_tools.video_manager.probe_video", side_effect=no_probe):
                project.archive_and_register_pptx(source)
            valid_data = copy.deepcopy(project.data)
            project.decks()[0]["assets"][0]["family_id"] = "missing-family"

            with self.assertRaisesRegex(ValueError, "unknown family"):
                project.save()

            reopened = VideoProject.open(root / "library")
            self.assertEqual(reopened.data, valid_data)

    def test_save_refuses_to_overwrite_an_unreadable_current_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            project = VideoProject.create(root / "library")
            project.data["name"] = "second revision"
            project.save()
            project.manifest_path.write_text("{broken", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, "拒绝覆盖"):
                project.save()

    def test_health_report_and_prune_only_remove_missing_output_records(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            make_video_pptx(source, b"video-source", "Source")
            project = VideoProject.create(root / "library")
            with patch("pptx_tools.video_manager.probe_video", side_effect=no_probe):
                project.archive_and_register_pptx(source)
            deck = project.decks()[0]
            deck["optimized_outputs"].append(
                {
                    "id": "missing-output",
                    "kind": "optimized",
                    "path": str(root / "deleted-output.pptx"),
                    "sha256": "f" * 64,
                    "size_bytes": 1,
                    "mtime_ns": 1,
                    "created_at": "2026-01-01T00:00:00+00:00",
                }
            )
            project.save()
            family_id = deck["assets"][0]["family_id"]
            untracked = project.root / "media" / "untracked.mp4"
            untracked.write_bytes(b"untracked")

            report = audit_video_project(project)

            self.assertTrue(report["ok"])
            self.assertEqual(report["stats"]["stale_output_records"], 1)
            self.assertEqual(report["stats"]["untracked_media_files"], 1)
            self.assertEqual(prune_missing_output_records(project), 1)
            self.assertEqual(project.decks()[0]["optimized_outputs"], [])
            self.assertEqual(project.decks()[0]["assets"][0]["family_id"], family_id)
            self.assertTrue(untracked.is_file())

    def test_health_report_flags_ambiguous_hash_ownership(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = VideoProject.create(Path(temp_dir) / "library")
            digest = "a" * 64
            for index in range(2):
                path = project.root / "media" / f"{index}.mp4"
                path.write_bytes(str(index).encode())
                variant = {
                    "id": f"variant-{index}",
                    "label": "source",
                    "profile": "original",
                    "path": project.encode_path(path),
                    "sha256": f"{index}" * 64,
                    "size_bytes": path.stat().st_size,
                    "mtime_ns": path.stat().st_mtime_ns,
                }
                project.families().append(
                    {
                        "id": f"family-{index}",
                        "name": f"Family {index}",
                        "active_variant_id": variant["id"],
                        "source_variant_id": variant["id"],
                        "known_hashes": [digest, variant["sha256"]],
                        "source_hashes": [variant["sha256"]],
                        "variants": [variant],
                    }
                )

            report = audit_video_project(project)

            self.assertFalse(report["ok"])
            self.assertEqual(report["issue_counts"]["ambiguous_known_hash"], 1)
            with self.assertRaisesRegex(RuntimeError, "同时属于多个视频族"):
                project.family_by_known_hash(digest)

    def test_variant_hash_conflicting_with_another_family_alias_is_reported(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = VideoProject.create(Path(temp_dir) / "library")
            first_path = project.root / "media" / "first.mp4"
            second_path = project.root / "media" / "second.mp4"
            first_path.write_bytes(b"first")
            second_path.write_bytes(b"second")
            shared_digest = "a" * 64
            for index, (path, digest, aliases) in enumerate(
                (
                    (first_path, shared_digest, [shared_digest]),
                    (second_path, "b" * 64, [shared_digest, "b" * 64]),
                )
            ):
                variant = {
                    "id": f"variant-{index}",
                    "label": "source",
                    "profile": "original",
                    "path": project.encode_path(path),
                    "sha256": digest,
                    "size_bytes": path.stat().st_size,
                    "mtime_ns": path.stat().st_mtime_ns,
                }
                project.families().append(
                    {
                        "id": f"family-{index}",
                        "name": f"Family {index}",
                        "active_variant_id": variant["id"],
                        "source_variant_id": variant["id"],
                        "known_hashes": aliases,
                        "source_hashes": [digest],
                        "variants": [variant],
                    }
                )

            report = audit_video_project(project)

            self.assertFalse(report["ok"])
            self.assertEqual(report["issue_counts"]["ambiguous_known_hash"], 1)
            with self.assertRaisesRegex(RuntimeError, "同时属于多个视频族"):
                project.family_by_known_hash(shared_digest)

    def test_full_hash_audit_distinguishes_mtime_drift_from_content_change(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.pptx"
            make_video_pptx(source, b"video-source", "Source")
            project = VideoProject.create(root / "library")
            with patch("pptx_tools.video_manager.probe_video", side_effect=no_probe):
                project.archive_and_register_pptx(source)
            variant = project.families()[0]["variants"][0]
            path = project.variant_path(variant)
            path.touch()

            fast = audit_video_project(project)
            full = audit_video_project(project, verify_hashes=True)

            self.assertFalse(fast["ok"])
            self.assertEqual(fast["stats"]["modified_variants"], 1)
            self.assertTrue(full["ok"])
            self.assertEqual(full["stats"]["modified_variants"], 0)
            self.assertEqual(full["stats"]["metadata_drift_variants"], 1)
            self.assertEqual(full["issue_counts"]["variant_metadata_drift"], 1)

    def test_doctor_cli_writes_machine_readable_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            project = VideoProject.create(root / "library")
            report_path = root / "health.json"

            result = main(
                [
                    "doctor",
                    str(project.root),
                    "--verify-hashes",
                    "--report",
                    str(report_path),
                ]
            )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(result, 0)
            self.assertTrue(report["ok"])
            self.assertEqual(report["mode"], "full_hash")
            self.assertEqual(report["stats"]["families"], 0)
            self.assertEqual(report["pruned_stale_output_records"], 0)


if __name__ == "__main__":
    unittest.main()
