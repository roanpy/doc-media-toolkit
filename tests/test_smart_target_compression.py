from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image

from pptx_video_compactor import (
    ImageAsset,
    VideoAsset,
    allocate_media_budgets,
    append_frame_rate_mode,
    applied_quality_threshold,
    audio_stream_is_usable,
    assign_image_plan,
    assign_quality_plan,
    audit_encoded_assets,
    build_output_pptx,
    classify_image_content,
    compact_input_path,
    consolidate_exact_duplicate_images,
    dynamic_package_reserve_bytes,
    default_media_output_path,
    encode_asset,
    is_supported_standalone_image,
    images_pixel_identical,
    mb_to_bytes,
    load_runtime_config,
    media_plan_signature,
    next_target_media_budget,
    parse_pptx_assets,
    quality_variant_output_path,
    should_copy,
    target_report_fields,
    target_size_filename_label,
    write_standalone_report,
)
from pptx_video_compactor_gui import CompressionWorker, build_namespace


class SmartTargetContractTest(unittest.TestCase):
    def test_target_capacity_uses_decimal_megabytes(self) -> None:
        self.assertEqual(mb_to_bytes(10), 10_000_000)
        self.assertEqual(mb_to_bytes(0.25), 250_000)
        self.assertEqual(target_size_filename_label(10), "10")
        self.assertEqual(target_size_filename_label(10.25), "10_25")

    def test_worker_progress_never_moves_backward_during_retries(self) -> None:
        worker = CompressionWorker([], None, "high", "high", "zh", {})
        emitted: list[int] = []
        worker.progress.connect(lambda percent, _label: emitted.append(percent))
        worker.total_bytes = 100
        worker._progress(100, 8, 10, "first")
        worker._progress(100, 2, 10, "retry")
        self.assertEqual(emitted, [80, 80])

    def test_media_plan_signature_ignores_runtime_status_changes(self) -> None:
        asset = VideoAsset(media_path="video.mp4", zip_size=1_000, status="encoded")
        before = media_plan_signature({asset.media_path: asset}, {})
        asset.status = "planned"
        self.assertEqual(media_plan_signature({asset.media_path: asset}, {}), before)

    def test_target_capacity_rejects_nonpositive_and_nonfinite_values(self) -> None:
        for value in (0, -1, float("inf"), float("nan")):
            with self.subTest(value=value), self.assertRaises(ValueError):
                mb_to_bytes(value)

    def test_experimental_outputs_are_visibly_isolated(self) -> None:
        with patch.dict(os.environ, {"PPTX_TOOLS_EXPERIMENTAL": "1"}):
            output = default_media_output_path(Path("photo.jpg"), 10)
        self.assertEqual(output.name, "photo_compressed_10MB_experimental.jpg")
        self.assertEqual(
            quality_variant_output_path(output, "forced").name,
            "photo_compressed_10MB_forced_experimental.jpg",
        )

    def test_ancillary_image_queue_accepts_safe_first_version_formats(self) -> None:
        for suffix in (".jpg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".gif"):
            self.assertTrue(is_supported_standalone_image(Path(f"image{suffix}")))

    def test_target_mode_defaults_to_cpu_unless_gpu_is_enabled(self) -> None:
        source = Path("source.pptx")
        self.assertEqual(
            build_namespace(source, 10, "high", "high").encoder,
            "cpu",
        )
        self.assertEqual(
            build_namespace(
                source, 10, "high", "high", target_gpu_enabled=True
            ).encoder,
            "gpu",
        )
        self.assertEqual(
            build_namespace(source, None, "high", "high").encoder,
            "auto",
        )
        self.assertEqual(
            build_namespace(
                source,
                None,
                "high",
                "high",
                standard_encoder_strategy="cpu",
            ).encoder,
            "cpu",
        )

    def test_source_already_under_target_is_reported_without_duplicate_output(
        self,
    ) -> None:
        from PIL import Image

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "small.png"
            Image.new("RGB", (16, 16), "red").save(source)

            result = compact_input_path(
                build_namespace(source, 1.0, "none", "high"),
                logger=lambda _message: None,
            )

            self.assertTrue(result["skipped"])
            self.assertEqual(result["output_pptx"], source.resolve())
            report_path = Path(result["report_path"])
            self.assertTrue(report_path.is_file())
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(report["target"]["status"], "source_already_meets")
            self.assertFalse((source.parent / "small_compressed_1MB.png").exists())

    def test_dynamic_package_reserve_is_small_and_bounded(self) -> None:
        self.assertEqual(dynamic_package_reserve_bytes(10_000_000, 2_000_000), 64_000)
        self.assertEqual(
            dynamic_package_reserve_bytes(1_000_000_000, 100_000_000),
            1_000_000,
        )

    def test_image_plan_protects_larger_displayed_images(self) -> None:
        assets = {
            "large": ImageAsset(
                media_path="ppt/media/large.jpg",
                zip_size=1_000,
                max_area_ratio=1.0,
                content_type="photo",
            ),
            "small": ImageAsset(
                media_path="ppt/media/small.jpg",
                zip_size=1_000,
                max_area_ratio=0.01,
                content_type="photo",
            ),
        }

        assign_image_plan(assets, "high", target_image_bytes=1_600)

        self.assertGreater(assets["large"].quality, assets["small"].quality)

    def test_document_image_plan_keeps_a_quality_fallback_original(self) -> None:
        asset = ImageAsset(
            media_path="image.jpg",
            zip_size=1_000,
            quality_status="restored_original",
        )
        assign_image_plan(
            {asset.media_path: asset},
            "aggressive",
            100,
            preserve_quality_fallbacks=True,
        )
        self.assertEqual(asset.status, "copy_requested")
        self.assertEqual(asset.quality, 100)
        self.assertEqual(asset.scale, 1.0)

    def test_applied_quality_threshold_protects_large_reused_media(self) -> None:
        small = applied_quality_threshold(0.95, 0.01, 1, 0.90)
        large_reused = applied_quality_threshold(0.95, 1.0, 4, 0.90)
        self.assertGreater(large_reused, small)
        self.assertLess(large_reused, 1.0)

    def test_target_plan_never_copies_video_that_requires_downscale(self) -> None:
        asset = VideoAsset(
            media_path="ppt/media/video.mp4",
            zip_size=1_000_000,
            width=1920,
            height=1080,
            selected_height=720,
            target_bytes=990_000,
            target_total_kbps=9_900,
            original_total_kbps=10_000,
        )

        self.assertFalse(should_copy(asset))

    def test_vfr_source_preserves_timestamps_when_not_downsampling(self) -> None:
        asset = VideoAsset(
            media_path="video.mp4", zip_size=1, original_fps=18.1101, target_fps=18.1101
        )
        command = ["ffmpeg"]
        append_frame_rate_mode(command, asset)
        self.assertEqual(command[-2:], ["-fps_mode", "passthrough"])

        reduced = VideoAsset(
            media_path="video.mp4", zip_size=1, original_fps=30.0, target_fps=15.0
        )
        command = ["ffmpeg"]
        append_frame_rate_mode(command, reduced)
        self.assertNotIn("-fps_mode", command)

    def test_empty_audio_stream_is_preserved_instead_of_reencoded(self) -> None:
        self.assertFalse(
            audio_stream_is_usable({"duration": "0.000000", "nb_frames": "1"})
        )
        self.assertTrue(audio_stream_is_usable({"codec_name": "aac"}))
        self.assertTrue(audio_stream_is_usable({"duration": "0.000000"}))
        self.assertTrue(audio_stream_is_usable({"nb_frames": "1"}))
        asset = VideoAsset(
            media_path="ppt/media/video.mp4",
            zip_size=1_000_000,
            width=960,
            height=540,
            has_audio=True,
            audio_stream_usable=False,
            selected_height=480,
            target_bytes=200_000,
            target_total_kbps=1_000,
            original_total_kbps=2_000,
        )
        self.assertTrue(should_copy(asset))

    def test_unusable_audio_plan_consumes_original_budget(self) -> None:
        asset = VideoAsset(
            media_path="ppt/media/video.mp4",
            zip_size=2_000,
            width=960,
            height=540,
            duration_sec=2.0,
            has_audio=True,
            audio_stream_usable=False,
            original_video_kbps=800,
            original_total_kbps=900,
            original_audio_kbps=100,
            selected_height=360,
        )
        assign_quality_plan(
            {asset.media_path: asset},
            target_video_bytes=1_000,
            min_height=360,
            profile="balanced",
            config=load_runtime_config(None, "balanced"),
        )
        self.assertEqual(asset.target_bytes, asset.zip_size)
        self.assertEqual(asset.selected_height, asset.height)
        self.assertEqual(asset.target_total_kbps, asset.original_total_kbps)

    def test_unusable_audio_encode_path_copies_source_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.mp4"
            output_dir = root / "output"
            output_dir.mkdir()
            source.write_bytes(b"video-with-declared-empty-audio")
            asset = VideoAsset(
                media_path="ppt/media/source.mp4",
                zip_size=source.stat().st_size,
                extracted_path=str(source),
                output_media_path="ppt/media/source.mp4",
                has_audio=True,
                audio_stream_usable=False,
                status="planned",
            )
            encode_asset(
                asset,
                output_dir,
                "medium",
                load_runtime_config(None, "balanced"),
                encoder_mode="cpu",
            )
            self.assertEqual(asset.status, "copied")
            self.assertEqual(
                (output_dir / "source.mp4").read_bytes(), source.read_bytes()
            )
            self.assertEqual(asset.quality_reason, "unusable_audio_stream_preserved")

    def test_rgba_rgb_difference_is_not_pixel_identical(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            left = Path(temp_dir) / "left.png"
            right = Path(temp_dir) / "right.png"
            Image.new("RGBA", (4, 4), (255, 0, 0, 255)).save(left)
            Image.new("RGBA", (4, 4), (0, 0, 255, 255)).save(right)
            self.assertFalse(images_pixel_identical(left, right))

    def test_failed_video_quality_restores_original(self) -> None:
        asset = VideoAsset(
            media_path="ppt/media/video1.mp4",
            zip_size=1_000,
            width=640,
            height=360,
            display_width_px=640,
            display_height_px=360,
            max_area_ratio=1.0,
            extracted_path="/tmp/original.mp4",
            output_path="/tmp/candidate.mp4",
            output_media_path="ppt/media/video1.mp4",
            status="encoded",
        )
        with patch("pptx_video_compactor.measure_media_ssim", return_value=0.80):
            audit_encoded_assets(
                {asset.media_path: asset},
                {},
                video_threshold=0.95,
                image_threshold=0.99,
            )
        self.assertEqual(asset.status, "copied")
        self.assertEqual(asset.output_path, asset.extracted_path)
        self.assertEqual(asset.quality_status, "restored_original")

    def test_target_quality_failure_keeps_compressed_candidate(self) -> None:
        asset = VideoAsset(
            media_path="ppt/media/video1.mp4",
            zip_size=1_000,
            width=640,
            height=360,
            display_width_px=640,
            display_height_px=360,
            max_area_ratio=1.0,
            extracted_path="/tmp/original.mp4",
            output_path="/tmp/candidate.mp4",
            output_media_path="ppt/media/video1.mp4",
            status="encoded",
        )
        with patch("pptx_video_compactor.measure_media_ssim", return_value=0.80):
            audit_encoded_assets(
                {asset.media_path: asset},
                {},
                video_threshold=0.95,
                image_threshold=0.99,
                restore_failed=False,
            )
        self.assertEqual(asset.status, "encoded")
        self.assertEqual(asset.quality_status, "below_threshold")
        self.assertEqual(asset.output_path, "/tmp/candidate.mp4")

    def test_target_quality_audit_failure_still_restores_original(self) -> None:
        asset = VideoAsset(
            media_path="ppt/media/video1.mp4",
            zip_size=1_000,
            width=640,
            height=360,
            extracted_path="/tmp/original.mp4",
            output_path="/tmp/candidate.mp4",
            status="encoded",
        )
        with patch(
            "pptx_video_compactor.measure_media_ssim",
            side_effect=OSError("ffmpeg unavailable"),
        ):
            audit_encoded_assets(
                {asset.media_path: asset},
                {},
                video_threshold=0.95,
                image_threshold=0.99,
                restore_failed=False,
            )
        self.assertEqual(asset.status, "copied")
        self.assertEqual(asset.quality_status, "restored_original")
        self.assertEqual(asset.output_path, asset.extracted_path)

    def test_forced_mode_still_enforces_absolute_video_redline(self) -> None:
        asset = VideoAsset(
            media_path="video.mp4",
            zip_size=1_000,
            width=640,
            height=360,
            display_width_px=64,
            display_height_px=36,
            max_area_ratio=0.01,
            extracted_path="/tmp/original.mp4",
            output_path="/tmp/candidate.mp4",
            status="encoded",
        )
        with patch("pptx_video_compactor.measure_media_ssim", return_value=0.89):
            audit_encoded_assets(
                {asset.media_path: asset},
                {},
                video_threshold=0.95,
                image_threshold=0.99,
                forced=True,
            )
        self.assertEqual(asset.status, "copied")
        self.assertEqual(asset.applied_threshold, 0.9)

    def test_image_metadata_change_restores_original(self) -> None:
        from PIL import Image, PngImagePlugin

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.png"
            candidate = root / "candidate.png"
            metadata = PngImagePlugin.PngInfo()
            metadata.add_text("Copyright", "Example")
            Image.new("RGB", (32, 32), "red").save(source, pnginfo=metadata)
            Image.new("RGB", (32, 32), "red").save(candidate)
            asset = ImageAsset(
                media_path="source.png",
                zip_size=source.stat().st_size,
                width=32,
                height=32,
                display_width_px=32,
                display_height_px=32,
                max_area_ratio=1.0,
                content_type="line_art",
                extracted_path=str(source),
                output_path=str(candidate),
                status="encoded",
            )
            with patch("pptx_video_compactor.measure_media_ssim", return_value=1.0):
                audit_encoded_assets(
                    {},
                    {asset.media_path: asset},
                    video_threshold=0.95,
                    image_threshold=0.99,
                )
        self.assertEqual(asset.quality_status, "restored_original")
        self.assertFalse(asset.metadata_preserved)
        self.assertEqual(asset.quality_reason, "image_metadata_changed")

    def test_joint_budget_reduces_video_and_images_together(self) -> None:
        video, image = allocate_media_budgets(
            8_000_000,
            2_000_000,
            7_500_000,
            "balanced",
            "balanced",
        )
        self.assertLess(video, 8_000_000)
        self.assertLess(image, 2_000_000)
        self.assertEqual(video + image, 7_500_000)

    def test_joint_budget_does_not_reduce_disabled_media_type(self) -> None:
        video, image = allocate_media_budgets(
            8_000_000,
            2_000_000,
            8_000_000,
            "balanced",
            "none",
        )
        self.assertEqual(image, 2_000_000)
        self.assertEqual(video, 6_000_000)

    def test_image_content_classification_is_local_and_deterministic(self) -> None:
        from PIL import Image

        transparent = Image.new("RGBA", (32, 32), (10, 20, 30, 0))
        line_art = Image.new("RGB", (32, 32), "white")
        self.assertEqual(classify_image_content(transparent), "transparent")
        self.assertEqual(classify_image_content(line_art), "line_art")

    def test_target_report_uses_real_output_size(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "output.bin"
            output.write_bytes(b"x" * 950_000)

            target = target_report_fields(1.0, output)

            self.assertEqual(target["target_bytes"], 1_000_000)
            self.assertEqual(target["actual_bytes"], 950_000)
            self.assertEqual(target["delta_bytes"], -50_000)
            self.assertEqual(target["target_ratio"], 0.95)
            self.assertEqual(target["status"], "met")

    def test_target_capacity_attempts_are_bounded(self) -> None:
        correction = next_target_media_budget(
            actual_bytes=1_100_000,
            target_bytes=1_000_000,
            current_media_budget=800_000,
            maximum_media_budget=900_000,
            correction_rounds=0,
            giveback_used=False,
        )
        self.assertEqual(correction[1], "correction")
        self.assertLess(correction[0], 800_000)
        self.assertIsNotNone(
            next_target_media_budget(
                actual_bytes=1_100_000,
                target_bytes=1_000_000,
                current_media_budget=800_000,
                maximum_media_budget=900_000,
                correction_rounds=0,
                giveback_used=False,
            )
        )
        self.assertIsNone(
            next_target_media_budget(
                actual_bytes=1_100_000,
                target_bytes=1_000_000,
                current_media_budget=800_000,
                maximum_media_budget=900_000,
                correction_rounds=2,
                giveback_used=False,
            )
        )

    def test_target_capacity_gives_quality_back_once(self) -> None:
        giveback = next_target_media_budget(
            actual_bytes=800_000,
            target_bytes=1_000_000,
            current_media_budget=600_000,
            maximum_media_budget=900_000,
            correction_rounds=0,
            giveback_used=False,
        )
        self.assertEqual(giveback, (731_250, "quality_giveback"))
        self.assertIsNone(
            next_target_media_budget(
                actual_bytes=800_000,
                target_bytes=1_000_000,
                current_media_budget=600_000,
                maximum_media_budget=900_000,
                correction_rounds=0,
                giveback_used=True,
            )
        )

    def test_json_report_also_writes_human_markdown(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.png"
            output = root / "output.png"
            source.write_bytes(b"source")
            output.write_bytes(b"output")
            report_path = root / "output.report.json"
            asset = ImageAsset(
                media_path=source.name,
                zip_size=source.stat().st_size,
                output_path=str(output),
                output_media_path=output.name,
                status="encoded",
            )

            write_standalone_report(
                report_path,
                source,
                output,
                image_asset=asset,
                target_size_mb=0.00001,
            )

            data = json.loads(report_path.read_text(encoding="utf-8"))
            markdown = report_path.with_suffix(".md")
            self.assertEqual(data["target"]["actual_bytes"], output.stat().st_size)
            self.assertTrue(markdown.is_file())
            self.assertIn("# 压缩报告", markdown.read_text(encoding="utf-8"))
            self.assertNotIn(str(root), markdown.read_text(encoding="utf-8"))

    def test_pptx_images_use_relationships_and_slide_display_area(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            deck = Path(temp_dir) / "images.pptx"
            with ZipFile(deck, "w", ZIP_DEFLATED) as archive:
                archive.writestr(
                    "ppt/presentation.xml",
                    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                    '<p:sldSz cx="1000" cy="500"/></p:presentation>',
                )
                archive.writestr(
                    "ppt/slides/slide1.xml",
                    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    '<p:cSld><p:spTree><p:pic><p:blipFill><a:blip r:embed="rId1"/>'
                    '</p:blipFill><p:spPr><a:xfrm><a:ext cx="500" cy="250"/>'
                    "</a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>",
                )
                archive.writestr(
                    "ppt/slides/_rels/slide1.xml.rels",
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    'Target="../media/image1.png"/></Relationships>',
                )
                archive.writestr("ppt/media/image1.png", b"referenced")
                archive.writestr("ppt/media/orphan.png", b"orphan")

            _, images, meta = parse_pptx_assets(
                deck,
                render_width=1000,
                render_height=500,
                overscan=1.0,
                min_height=480,
                max_height=1080,
                config=load_runtime_config(None, "high"),
                include_videos=False,
            )

            self.assertEqual(set(images), {"ppt/media/image1.png"})
            image = images["ppt/media/image1.png"]
            self.assertEqual(image.max_area_ratio, 0.25)
            self.assertEqual(
                [image.display_width_px, image.display_height_px], [500, 250]
            )
            self.assertEqual(meta["orphan_image_paths"], ["ppt/media/orphan.png"])

    def test_exact_duplicate_and_orphan_images_are_safely_removed(self) -> None:
        from PIL import Image

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            image = root / "image.png"
            Image.new("RGB", (8, 8), "blue").save(image)
            source = root / "source.pptx"
            output = root / "output.pptx"
            with ZipFile(source, "w", ZIP_DEFLATED) as archive:
                archive.writestr("ppt/media/canonical.png", image.read_bytes())
                archive.writestr("ppt/media/duplicate.png", image.read_bytes())
                archive.writestr("ppt/media/orphan.png", image.read_bytes())
                archive.writestr(
                    "ppt/slides/_rels/slide1.xml.rels",
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="image" Target="../media/duplicate.png"/>'
                    "</Relationships>",
                )
            assets = {
                name: ImageAsset(
                    media_path=name,
                    zip_size=image.stat().st_size,
                    extracted_path=str(image),
                    output_path=str(image),
                    status="copied",
                )
                for name in (
                    "ppt/media/canonical.png",
                    "ppt/media/duplicate.png",
                )
            }
            duplicate_map = consolidate_exact_duplicate_images(assets)
            build_output_pptx(
                source,
                output,
                {},
                assets,
                relationship_path_map=duplicate_map,
                remove_paths={*duplicate_map, "ppt/media/orphan.png"},
            )
            with ZipFile(output) as archive:
                self.assertIn("ppt/media/canonical.png", archive.namelist())
                self.assertNotIn("ppt/media/duplicate.png", archive.namelist())
                self.assertNotIn("ppt/media/orphan.png", archive.namelist())
                relationships = archive.read(
                    "ppt/slides/_rels/slide1.xml.rels"
                ).decode()
            self.assertIn("../media/canonical.png", relationships)


if __name__ == "__main__":
    unittest.main()
