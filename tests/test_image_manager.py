from __future__ import annotations

import io
import json
import shutil
import tempfile
import threading
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from unittest.mock import patch

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from pptx_tools.image_manager import ImageProject, main


def image_bytes(color: tuple[int, int, int], size: tuple[int, int] = (32, 20)) -> bytes:
    output = io.BytesIO()
    Image.new("RGB", size, color).save(output, format="PNG")
    return output.getvalue()


def office_document(
    path: Path, prefix: str, owner: str, target: str, data: bytes
) -> None:
    owner_dir = Path(owner).parent.as_posix()
    owner_name = Path(owner).name
    rels = f"{owner_dir}/_rels/{owner_name}.rels"
    relationship = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="{target}"/>'
        "</Relationships>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(owner, "<root/>")
        archive.writestr(rels, relationship)
        archive.writestr(f"{prefix}/image1.png", data)
        archive.writestr(f"{prefix}/unused.png", image_bytes((1, 2, 3)))


class ImageManagerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_import_exact_duplicate_stores_one_file(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "first.png"
        second = self.root / "second.png"
        data = image_bytes((200, 30, 20))
        first.write_bytes(data)
        second.write_bytes(data)

        result = project.import_paths([first, second])

        self.assertEqual(result["added"], 1)
        self.assertEqual(result["reused"], 1)
        self.assertEqual(len(project.assets()), 1)
        self.assertEqual(len(project.assets()[0]["origins"]), 2)
        self.assertEqual(len(list((project.root / "images").rglob("*.png"))), 1)

    def test_lossless_reencoding_with_same_pixels_reuses_asset(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "first.png"
        second = self.root / "second.bmp"
        image = Image.new("RGB", (32, 20), (200, 30, 20))
        image.save(first, format="PNG")
        image.save(second, format="BMP")

        result = project.import_paths([first, second])

        self.assertEqual(result["added"], 1)
        self.assertEqual(result["reused"], 1)
        self.assertEqual(len(project.assets()), 1)
        self.assertEqual(len(project.assets()[0]["origins"]), 2)

    def test_import_can_skip_images_below_configured_dimensions(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "small.png"
        source.write_bytes(image_bytes((200, 30, 20), (32, 48)))

        result = project.import_paths([source], min_width=64, min_height=64)

        self.assertEqual(result["added"], 0)
        self.assertEqual(result["failed"], [])
        self.assertEqual(len(result["skipped"]), 1)
        self.assertEqual(result["skipped"][0]["width"], 32)
        self.assertEqual(result["skipped"][0]["height"], 48)
        self.assertEqual(project.assets(), [])

    def test_import_category_preserves_reused_image_metadata(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((20, 30, 40)))

        project.import_paths([source], category="示例项目/2026")
        project.import_paths([source], category="复用分类")

        self.assertEqual(project.assets()[0]["category"], "示例项目/2026")

    def test_import_normalizes_name_and_rejects_unsafe_category(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "1662562042-示例设备_fault.png"
        source.write_bytes(image_bytes((20, 30, 40)))

        project.import_paths([source], category="示例分类/异常样本")

        self.assertEqual(
            project.assets()[0]["name"],
            "示例设备_异常_1662562042",
        )
        self.assertEqual(project.assets()[0]["category"], "示例分类/异常样本")
        with self.assertRaises(ValueError):
            project.import_paths([source], category="../outside")

    def test_office_import_only_keeps_referenced_images(self) -> None:
        project = ImageProject.create(self.root / "library")
        pptx = self.root / "sample.pptx"
        docx = self.root / "sample.docx"
        office_document(
            pptx,
            "ppt/media",
            "ppt/slides/slide1.xml",
            "../media/image1.png",
            image_bytes((220, 50, 40)),
        )
        office_document(
            docx,
            "word/media",
            "word/document.xml",
            "media/image1.png",
            image_bytes((40, 80, 220)),
        )

        result = project.import_paths([pptx, docx])

        self.assertEqual(result["added"], 2)
        self.assertEqual(result["reused"], 0)
        self.assertEqual(
            {item["origins"][0]["source_type"] for item in project.assets()},
            {"pptx", "docx"},
        )

    def test_pdf_import_extracts_embedded_image_without_rendering_pages(self) -> None:
        project = ImageProject.create(self.root / "library")
        pdf_path = self.root / "sample.pdf"
        image_data = image_bytes((60, 180, 80), (80, 50))
        canvas = Canvas(str(pdf_path), pagesize=(300, 200))
        canvas.drawImage(ImageReader(io.BytesIO(image_data)), 20, 20, 80, 50)
        canvas.drawString(20, 150, "digital PDF text")
        canvas.save()

        result = project.import_paths([pdf_path])

        self.assertEqual(result["added"], 1)
        self.assertEqual(len(project.assets()), 1)
        self.assertEqual(project.assets()[0]["origins"][0]["source_type"], "pdf")

    def test_webp_library_remains_valid_after_directory_move(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.webp"
        Image.new("RGB", (96, 64), (30, 80, 160)).save(source, format="WEBP")
        project.import_paths([source])
        self.assertEqual(project.assets()[0]["format"], "WEBP")
        project.assets()[0]["format"] = ""
        project.save()
        moved_root = self.root / "moved" / "library"
        moved_root.parent.mkdir()
        shutil.copytree(project.root, moved_root)

        moved = ImageProject.open(moved_root)

        self.assertEqual(moved.assets()[0]["format"], "WEBP")
        self.assertTrue(moved.asset_path(moved.assets()[0]).is_file())
        self.assertTrue(moved.health_report(verify_hashes=True)["ok"])

    def test_health_report_detects_missing_modified_and_orphan_files(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "first.png"
        second = self.root / "second.png"
        first.write_bytes(image_bytes((10, 20, 30)))
        second.write_bytes(image_bytes((40, 50, 60)))
        project.import_paths([first, second])
        first_asset, second_asset = project.assets()
        first_path = project.asset_path(first_asset)
        first_path.write_bytes(b"x" * first_asset["size_bytes"])
        project.asset_path(second_asset).unlink()
        orphan = project.root / "images" / "orphan.bin"
        orphan.write_bytes(b"orphan")

        report = project.health_report(verify_hashes=True)

        self.assertFalse(report["ok"])
        self.assertEqual(report["modified_files"], [first_asset["path"]])
        self.assertEqual(report["missing_files"], [second_asset["path"]])
        self.assertEqual(report["orphan_files"], ["images/orphan.bin"])

    def test_doctor_cli_writes_machine_readable_report(self) -> None:
        project = ImageProject.create(self.root / "library")
        report_path = self.root / "image-health.json"

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
        self.assertTrue(report["hashes_verified"])

    def test_similar_candidates_are_suggestions_not_automatic_merges(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "one.png"
        second = self.root / "two.png"
        first.write_bytes(image_bytes((100, 100, 100)))
        second.write_bytes(image_bytes((101, 101, 101)))
        project.import_paths([first, second])

        candidates = project.similar_candidates(project.assets()[0]["id"])

        self.assertEqual(len(project.assets()), 2)
        self.assertTrue(candidates)
        self.assertEqual(candidates[0]["asset_id"], project.assets()[1]["id"])

    def test_image_health_and_ignored_similar_pairs_persist(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "one.png"
        first_copy = self.root / "one-copy.png"
        second = self.root / "two.png"
        small = self.root / "small.png"
        first.write_bytes(image_bytes((100, 100, 100)))
        first_copy.write_bytes(first.read_bytes())
        second.write_bytes(image_bytes((101, 101, 101)))
        small.write_bytes(image_bytes((20, 30, 40), (8, 8)))
        project.import_paths([first, first_copy, second, small])
        primary, candidate, small_asset = project.assets()
        small_asset["origins"] = []
        project.save()

        self.assertEqual(
            project.health_counts(min_width=16, min_height=16),
            {
                "all": 3,
                "duplicate_origins": 1,
                "similar": 2,
                "undersized": 1,
                "no_origin": 1,
            },
        )

        project.ignore_similar_pair(primary["id"], candidate["id"])
        reopened = ImageProject.open(project.root)
        self.assertEqual(reopened.similar_candidates(primary["id"]), [])
        self.assertEqual(reopened.health_counts()["similar"], 0)
        self.assertEqual(reopened.reset_ignored_similar_pairs(), 1)
        self.assertTrue(reopened.similar_candidates(primary["id"]))

    def test_cleanup_only_removes_unreferenced_library_files(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        referenced = project.asset_path(project.assets()[0])
        orphan = referenced.parent / "orphan.tmp"
        orphan.write_bytes(b"unused")

        self.assertEqual(project.orphan_paths(), [orphan])
        removed = project.cleanup_orphans()

        self.assertEqual(removed, [orphan])
        self.assertFalse(orphan.exists())
        self.assertEqual(len(project.pending_cleanup()), 1)
        self.assertEqual(
            project.restore_cleanup_entry(project.pending_cleanup()[0]["token"]),
            orphan,
        )
        self.assertTrue(orphan.is_file())
        self.assertTrue(referenced.is_file())

    def test_pixel_dedupe_logs_unreadable_legacy_asset(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        asset = project.assets()[0]
        metadata = {
            key: asset[key] for key in ("width", "height", "dhash", "pixel_sha256")
        }
        asset["pixel_sha256"] = ""
        project.asset_path(asset).unlink()

        with self.assertLogs("pptx_tools.image_manager", level="WARNING"):
            self.assertIsNone(project.find_by_pixels(metadata))

    def test_asset_health_stops_after_first_similar_match(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        asset = project.assets()[0]

        def matches():
            yield asset, 0, 0.0
            raise AssertionError("asset_health scanned after finding a match")

        with patch.object(project, "_similar_matches", return_value=matches()):
            self.assertTrue(project.asset_health(asset["id"])["similar"])

    def test_restore_recovers_after_manifest_saved_before_cleanup_index(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        project.remove_asset(project.assets()[0]["id"])
        pending = project.pending_cleanup()[0]
        quarantined = Path(pending["quarantined_path"])
        target = project.root / pending["original_path"]

        target.parent.mkdir(parents=True, exist_ok=True)
        quarantined.replace(target)
        project.assets().append(pending["asset"])
        project.save()

        self.assertEqual(project.restore_cleanup_entry(pending["token"]), target)
        self.assertEqual(project.pending_cleanup(), [])
        self.assertEqual(len(project.assets()), 1)

    def test_open_recovers_interrupted_quarantine_move_and_restores(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        asset = project.assets()[0]
        original = project.asset_path(asset)
        quarantined = project.cleanup_dir / f"recovery_{original.name}"
        entry = {
            "token": "recovery-token",
            "asset": asset,
            "original_path": original.relative_to(project.root).as_posix(),
            "quarantined_path": quarantined.relative_to(project.root).as_posix(),
            "sha256": asset["sha256"],
            "quarantined_at": "2026-07-31T00:00:00+00:00",
            "reason": "test",
            "state": "moving",
        }
        project._write_cleanup_index([entry])
        shutil.move(str(original), quarantined)

        reopened = ImageProject.open(project.root)

        pending = reopened.pending_cleanup()
        self.assertEqual(pending[0]["state"], "quarantined")
        self.assertEqual(reopened.restore_cleanup_entry("recovery-token"), original)
        self.assertTrue(original.is_file())
        self.assertEqual(reopened.pending_cleanup(), [])

    def test_cleanup_index_paths_are_restricted_to_managed_directories(self) -> None:
        project = ImageProject.create(self.root / "library")
        outside = self.root / "outside.png"
        outside.write_bytes(b"keep")
        project.cleanup_dir.mkdir(exist_ok=True)
        project.cleanup_index_path.write_text(
            json.dumps(
                [
                    {
                        "token": "escape",
                        "original_path": "images/restored.png",
                        "quarantined_path": str(outside),
                    }
                ]
            ),
            encoding="utf-8",
        )
        with self.assertRaises(RuntimeError):
            project.empty_cleanup()
        self.assertTrue(outside.is_file())

        project.cleanup_index_path.write_text(
            json.dumps(
                [
                    {
                        "token": "escape",
                        "original_path": str(outside),
                        "quarantined_path": "_cleanup/held.png",
                        "asset": {},
                    }
                ]
            ),
            encoding="utf-8",
        )
        with self.assertRaises(RuntimeError):
            project.restore_cleanup_entry("escape")

    def test_failed_remove_keeps_quarantine_index_when_original_is_recreated(
        self,
    ) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        asset = project.assets()[0]
        stored = project.asset_path(asset)

        def fail_after_recreate() -> None:
            stored.parent.mkdir(parents=True, exist_ok=True)
            stored.write_bytes(b"concurrent replacement")
            raise RuntimeError("stale revision")

        with (
            patch.object(project, "save", side_effect=fail_after_recreate),
            self.assertRaisesRegex(RuntimeError, "stale revision"),
        ):
            project.remove_asset(asset["id"])

        pending = project.pending_cleanup()
        self.assertEqual(len(pending), 1)
        self.assertTrue(Path(pending[0]["quarantined_path"]).is_file())

    def test_open_recovers_primary_manifest_and_can_save_again(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        project.manifest_path.write_text("{broken", encoding="utf-8")

        recovered = ImageProject.open(project.root)
        self.assertTrue(recovered.recovered_from_backup)
        self.assertTrue(recovered.recovery_detail)
        recovered.data["name"] = "已恢复"
        recovered.save()

        self.assertEqual(ImageProject.open(project.root).data["name"], "已恢复")
        json.loads(project.manifest_path.read_text(encoding="utf-8"))

    def test_concurrent_saves_do_not_both_overwrite_the_same_revision(self) -> None:
        project = ImageProject.create(self.root / "library")
        left = ImageProject.open(project.root)
        right = ImageProject.open(project.root)
        barrier = threading.Barrier(2)
        outcomes: list[str] = []

        def save(instance: ImageProject, name: str) -> None:
            instance.data["name"] = name
            barrier.wait()
            try:
                instance.save()
            except RuntimeError:
                outcomes.append("stale")
            else:
                outcomes.append("saved")

        threads = [
            threading.Thread(target=save, args=(left, "left")),
            threading.Thread(target=save, args=(right, "right")),
        ]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join(timeout=3)

        self.assertEqual(sorted(outcomes), ["saved", "stale"])
        self.assertIn(ImageProject.open(project.root).data["name"], {"left", "right"})

    def test_corrupt_source_does_not_abort_other_imports_or_leave_orphans(
        self,
    ) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        corrupt = self.root / "corrupt.pptx"
        corrupt.write_bytes(b"not-a-zip")

        result = project.import_paths([source, corrupt])

        self.assertEqual(result["added"], 1)
        self.assertEqual(len(result["failed"]), 1)
        self.assertEqual(len(ImageProject.open(project.root).assets()), 1)
        self.assertEqual(len(list((project.root / "images").rglob("*.png"))), 1)

    def test_save_failure_rolls_back_new_files_and_memory(self) -> None:
        project = ImageProject.create(self.root / "library")
        stale = ImageProject.open(project.root)
        project.data["name"] = "new revision"
        project.save()
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))

        with self.assertRaises(RuntimeError):
            stale.import_paths([source])

        self.assertEqual(stale.assets(), [])
        self.assertEqual(list((stale.root / "images").rglob("*.png")), [])

    def test_stale_metadata_and_removal_changes_are_rolled_back(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        stale = ImageProject.open(project.root)
        asset_id = stale.assets()[0]["id"]
        original_name = stale.assets()[0]["name"]
        project.data["name"] = "new revision"
        project.save()

        with self.assertRaises(RuntimeError):
            stale.update_metadata(asset_id, name="不应保留")
        self.assertEqual(stale.asset(asset_id)["name"], original_name)

        with self.assertRaises(RuntimeError):
            stale.remove_asset(asset_id)
        self.assertEqual(stale.asset(asset_id)["name"], original_name)
        self.assertTrue(stale.asset_path(stale.asset(asset_id)).is_file())

    def test_confirmed_image_merge_preserves_origins_and_removes_duplicate_file(
        self,
    ) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "first.png"
        second = self.root / "second.png"
        first.write_bytes(image_bytes((100, 100, 100)))
        second.write_bytes(image_bytes((101, 101, 101)))
        project.import_paths([first, second])
        primary, duplicate = project.assets()
        duplicate_path = project.asset_path(duplicate)

        result = project.merge_assets(
            primary["id"],
            [primary["id"], duplicate["id"]],
            confirmed_same_content=True,
        )

        self.assertEqual(result["removed_ids"], [duplicate["id"]])
        self.assertEqual(len(project.assets()), 1)
        self.assertEqual(len(project.assets()[0]["origins"]), 2)
        self.assertFalse(duplicate_path.exists())
        pending = project.pending_cleanup()
        self.assertEqual(len(pending), 1)
        project.restore_cleanup_entry(pending[0]["token"])
        self.assertTrue(duplicate_path.is_file())
        self.assertEqual(len(project.assets()), 2)

    def test_remove_asset_is_quarantined_until_permanent_cleanup(self) -> None:
        project = ImageProject.create(self.root / "library")
        source = self.root / "source.png"
        source.write_bytes(image_bytes((10, 20, 30)))
        project.import_paths([source])
        asset = project.assets()[0]
        stored = project.asset_path(asset)

        project.remove_asset(asset["id"])

        self.assertFalse(stored.exists())
        self.assertEqual(project.assets(), [])
        pending = project.pending_cleanup()
        self.assertEqual(len(pending), 1)
        self.assertEqual(project.empty_cleanup(), 1)
        self.assertEqual(project.pending_cleanup(), [])

    def test_image_merge_requires_explicit_confirmation(self) -> None:
        project = ImageProject.create(self.root / "library")
        first = self.root / "first.png"
        second = self.root / "second.png"
        first.write_bytes(image_bytes((100, 100, 100)))
        second.write_bytes(image_bytes((101, 101, 101)))
        project.import_paths([first, second])
        ids = [asset["id"] for asset in project.assets()]

        with self.assertRaises(ValueError):
            project.merge_assets(ids[0], ids)


if __name__ == "__main__":
    unittest.main()
