from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image
from pptx import Presentation
from reportlab.pdfgen import canvas
from pptx_video_compactor import (
    IMAGE_EXTENSIONS,
    MP4_CONTAINER_EXTENSIONS,
    VIDEO_EXTENSIONS,
)

from .dependencies import missing_dependency_message
from .models import ExportArtifacts, ExportOptions
from .pdf_io import open_pdf_reader
from .pdf_rendering import batch_render_pdf_slides
from .pdf_watermark import apply_watermark_to_pdf
from .pptx_fonts import replace_pptx_fonts
from .pptx_video_support import (
    export_sidecar_videos,
    extract_video_poster_frame,
    prepare_videos_for_image_pptx,
    reinsert_videos_into_pptx,
    scan_fidelity_warnings,
    slide_visibility_counts,
    visible_slide_number_map,
    watermark_video_file,
    watermark_videos_in_editable_pptx,
)
from .pptx_rebuild import (
    add_watermark_overlay_to_pptx,
    build_image_pptx_from_images,
    copy_pptx,
)
from .presentation_rendering import convert_document_to_pdf
from .runtime_temp import create_runtime_temp_dir
from .watermarking import apply_watermark_to_image

Logger = callable


def is_standalone_image_input(input_path: Path) -> bool:
    return input_path.suffix.lower() in IMAGE_EXTENSIONS


def is_standalone_video_input(input_path: Path) -> bool:
    return input_path.suffix.lower() in VIDEO_EXTENSIONS


def _render_limits_for_dpi(dpi: int) -> tuple[int, int]:
    if dpi >= 220:
        return (4096, 10_000_000)
    if dpi >= 192:
        return (3072, 6_500_000)
    return (2400, 4_000_000)


def default_output_path(input_path: Path, output_format: str, output_mode: str) -> Path:
    if is_standalone_image_input(input_path) or is_standalone_video_input(input_path):
        return input_path.with_name(f"{input_path.stem}_watermarked.{output_format}")
    suffix = f"_{output_mode}_watermarked.{output_format}"
    return input_path.with_name(f"{input_path.stem}{suffix}")


def effective_output_format(input_path: Path, requested_output_format: str) -> str:
    if is_standalone_image_input(input_path):
        return input_path.suffix.lower().lstrip(".")
    if is_standalone_video_input(input_path):
        source_suffix = input_path.suffix.lower()
        if source_suffix in MP4_CONTAINER_EXTENSIONS:
            return source_suffix.lstrip(".")
        return "mp4"
    return requested_output_format if input_path.suffix.lower() == ".pptx" else "pdf"


def _copy_to_final(src: Path, dst: Path) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def _render_pdf_pages(
    pdf_path: Path,
    *,
    options: ExportOptions,
    artifacts: ExportArtifacts,
) -> list[Path]:
    reader = open_pdf_reader(str(pdf_path))
    rendered_dir = create_runtime_temp_dir(
        "pptx_output_watermark_rendered_",
        purpose="rendered_pdf_pages",
    )
    artifacts.rendered_dir = rendered_dir
    max_edge, max_pixels = _render_limits_for_dpi(options.dpi)
    rendered = batch_render_pdf_slides(
        pdf_path,
        num_slides=len(reader.pages),
        output_dir=rendered_dir,
        dpi=options.dpi,
        jpeg_quality=options.jpeg_quality,
        max_edge=max_edge,
        max_pixels=max_pixels,
    )
    return [rendered[idx] for idx in sorted(rendered)]


def _watermark_images(
    image_paths: list[Path],
    *,
    options: ExportOptions,
    artifacts: ExportArtifacts,
) -> list[Path]:
    if not options.watermark.enabled:
        return image_paths
    output_dir = create_runtime_temp_dir(
        "pptx_output_watermark_images_",
        purpose="watermarked_page_images",
    )
    artifacts.watermarked_dir = output_dir
    result: list[Path] = []
    for image_path in image_paths:
        output_path = output_dir / image_path.name
        result.append(
            apply_watermark_to_image(
                image_path,
                output_path,
                options.watermark,
                jpeg_quality=options.jpeg_quality,
            )
        )
    return result


def _pdf_page_sizes(pdf_path: Path) -> list[tuple[float, float]]:
    reader = open_pdf_reader(str(pdf_path))
    return [
        (float(page.mediabox.width), float(page.mediabox.height))
        for page in reader.pages
    ]


def _build_pdf_from_images(
    image_paths: list[Path],
    output_path: Path,
    *,
    page_sizes: list[tuple[float, float]] | None = None,
) -> Path:
    if not image_paths:
        raise RuntimeError("No images available to build image-based PDF.")
    if page_sizes and len(page_sizes) != len(image_paths):
        raise RuntimeError("Rendered page count does not match source PDF page count.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(output_path))
    try:
        for index, image_path in enumerate(image_paths):
            if page_sizes:
                width, height = page_sizes[index]
            else:
                with Image.open(image_path) as image:
                    width, height = image.size
            pdf.setPageSize((width, height))
            pdf.drawImage(str(image_path), 0, 0, width=width, height=height)
            pdf.showPage()
        pdf.save()
        return output_path
    except Exception:
        try:
            output_path.unlink()
        except FileNotFoundError:
            pass
        raise


def cleanup_artifacts(artifacts: ExportArtifacts) -> None:
    if artifacts.temp_pdf:
        try:
            artifacts.temp_pdf.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass
        try:
            artifacts.temp_pdf.parent.rmdir()
        except OSError:
            pass
    for path in (artifacts.rendered_dir, artifacts.watermarked_dir):
        if not path:
            continue
        shutil.rmtree(path, ignore_errors=True)
    if artifacts.preprocessed_dir:
        shutil.rmtree(artifacts.preprocessed_dir, ignore_errors=True)
    if artifacts.video_work_dir:
        shutil.rmtree(artifacts.video_work_dir, ignore_errors=True)


def _prepare_input_pptx(
    input_path: Path, options: ExportOptions, artifacts: ExportArtifacts
) -> Path:
    if not options.replace_source_fonts:
        return input_path
    preprocessed_dir = create_runtime_temp_dir(
        "pptx_output_watermark_preprocessed_",
        purpose="preprocessed_pptx_copy",
    )
    artifacts.preprocessed_dir = preprocessed_dir
    output_path = preprocessed_dir / input_path.name
    return replace_pptx_fonts(
        input_path,
        output_path,
        replacement_family=options.source_font_family,
        font_names=options.source_font_names,
    )


def _export_standalone_image(
    input_path: Path,
    output_path: Path,
    *,
    options: ExportOptions,
) -> Path:
    if options.watermark.enabled:
        return apply_watermark_to_image(
            input_path,
            output_path,
            options.watermark,
            jpeg_quality=options.jpeg_quality,
        )
    return _copy_to_final(input_path, output_path)


def _export_standalone_video(
    input_path: Path,
    output_path: Path,
    *,
    options: ExportOptions,
    logger,
    artifacts: ExportArtifacts,
) -> Path:
    if not options.watermark.enabled:
        return _copy_to_final(input_path, output_path)
    video_work_dir = create_runtime_temp_dir(
        "pptx_output_watermark_video_",
        purpose="standalone_video_watermark",
    )
    artifacts.video_work_dir = video_work_dir
    poster_path = video_work_dir / "poster.jpg"
    extract_video_poster_frame(input_path, poster_path)
    logger(f"Watermarking standalone video ({options.video_quality_profile})")
    return watermark_video_file(
        input_path,
        output_path,
        options.watermark,
        encoder_mode=options.video_encoder,
        quality_profile=options.video_quality_profile,
        logger=logger,
    )


def export_document(
    options: ExportOptions,
    *,
    logger=print,
) -> Path:
    input_path = options.input_path.resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    dependency_message = missing_dependency_message(options)
    if dependency_message:
        raise RuntimeError(dependency_message)

    is_pdf = input_path.suffix.lower() == ".pdf"
    is_pptx = input_path.suffix.lower() == ".pptx"
    is_image = is_standalone_image_input(input_path)
    is_video = is_standalone_video_input(input_path)
    resolved_output_format = effective_output_format(input_path, options.output_format)
    if options.output_path is None:
        output_path = default_output_path(
            input_path,
            resolved_output_format,
            options.output_mode,
        ).resolve()
    else:
        output_path = options.output_path.resolve()
        expected_suffix = f".{resolved_output_format}"
        if output_path.suffix.lower() != expected_suffix:
            output_path = output_path.with_suffix(expected_suffix)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    artifacts = ExportArtifacts()

    try:
        if is_image:
            logger("Exporting standalone image")
            return _export_standalone_image(input_path, output_path, options=options)
        if is_video:
            return _export_standalone_video(
                input_path,
                output_path,
                options=options,
                logger=logger,
                artifacts=artifacts,
            )
        if is_pptx:
            prepared_input = _prepare_input_pptx(input_path, options, artifacts)
        else:
            prepared_input = input_path

        final_output_format = resolved_output_format

        if (
            options.output_mode == "editable"
            and final_output_format == "pptx"
            and is_pptx
        ):
            logger("Exporting editable PPTX")
            if not options.watermark.enabled:
                return copy_pptx(prepared_input, output_path)
            # Step 1: overlay page watermark image onto every slide.
            add_watermark_overlay_to_pptx(
                prepared_input, output_path, options.watermark
            )
            # Step 2: watermark embedded videos via media replacement, which
            # preserves all slide XML (timing, playback settings, audio, z-order)
            # — unlike the image-PPTX rebuild path.
            logger("Watermarking embedded videos in editable PPTX")
            watermark_videos_in_editable_pptx(
                output_path,
                output_path,
                options.watermark,
                encoder_mode=options.video_encoder,
                quality_profile=options.video_quality_profile,
                logger=logger,
            )
            return output_path

        if is_pdf:
            logger("Input is already PDF, skipping conversion")
            temp_pdf = prepared_input
        else:
            logger(f"Converting {input_path.suffix[1:].upper()} to PDF")
            temp_pdf = convert_document_to_pdf(prepared_input, logger=logger)
            artifacts.temp_pdf = temp_pdf

        if options.output_mode == "editable" and final_output_format == "pdf":
            logger("Exporting editable PDF")
            if options.watermark.enabled:
                return apply_watermark_to_pdf(temp_pdf, output_path, options.watermark)
            return _copy_to_final(temp_pdf, output_path)

        logger("Rendering PDF pages to images")
        page_images = _render_pdf_pages(temp_pdf, options=options, artifacts=artifacts)
        page_images = _watermark_images(
            page_images, options=options, artifacts=artifacts
        )

        if final_output_format == "pdf":
            logger("Building image-based PDF")
            return _build_pdf_from_images(
                page_images,
                output_path,
                page_sizes=_pdf_page_sizes(temp_pdf),
            )

        logger("Building image-based PPTX")
        source_prs = Presentation(str(prepared_input))
        built_output = build_image_pptx_from_images(
            page_images,
            output_path,
            slide_width_emu=int(source_prs.slide_width),
            slide_height_emu=int(source_prs.slide_height),
        )
        fidelity_warnings = scan_fidelity_warnings(prepared_input)
        if fidelity_warnings:
            logger("保真度提示：以下内容在图片 PPTX 模式下不会保留：")
            for warning in fidelity_warnings:
                logger(f"  · {warning}")
        if not options.preserve_videos_in_image_pptx:
            return built_output

        logger("Preparing embedded videos for image-based PPTX")
        video_work_dir = create_runtime_temp_dir(
            "pptx_output_watermark_video_",
            purpose="image_pptx_video_processing",
        )
        artifacts.video_work_dir = video_work_dir
        source_slide_count, hidden_slide_count = slide_visibility_counts(prepared_input)
        slide_number_map = visible_slide_number_map(prepared_input)
        if hidden_slide_count:
            logger(
                "Hidden slides detected: "
                f"{hidden_slide_count}/{source_slide_count}; mapping videos to visible exported pages"
            )
        assets = prepare_videos_for_image_pptx(
            prepared_input,
            video_work_dir,
            options.watermark,
            encoder_mode=options.video_encoder,
            slide_number_map=slide_number_map,
            logger=logger,
        )
        if not assets:
            logger("No embedded videos found to reinsert")
            return built_output

        try:
            logger("Reinserting embedded videos into image-based PPTX")
            inserted = reinsert_videos_into_pptx(built_output, assets, logger=logger)
            logger(f"Reinserted {inserted} video placement(s)")
            return built_output
        except Exception as exc:
            sidecar = export_sidecar_videos(built_output, assets)
            if sidecar is not None:
                logger(
                    f"Video reinsertion failed; exported watermarked sidecar videos to {sidecar.directory}"
                )
                logger(f"Video sidecar manifest: {sidecar.manifest_path}")
                logger(f"Embedded video reinsertion fallback reason: {exc}")
                return built_output
            raise RuntimeError(f"Embedded video reinsertion failed: {exc}") from exc
    finally:
        if not options.keep_artifacts:
            cleanup_artifacts(artifacts)
