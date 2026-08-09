from __future__ import annotations

import json
import shutil
import tempfile
from defusedxml import ElementTree as ET
from pathlib import Path
from typing import Any, Callable
from zipfile import BadZipFile, ZipFile

import pypdfium2 as pdfium
from PIL import Image, UnidentifiedImageError
from pypdf import PdfReader

from pptx_output_watermark.presentation_rendering import convert_document_to_pdf
from pptx_video_compactor import (
    ImageAsset,
    allocate_media_budgets,
    assign_image_plan,
    audit_encoded_assets,
    classify_image_content,
    dynamic_package_reserve_bytes,
    encode_image_asset,
    experimental_output_stem,
    image_detail_metrics,
    image_report_entry,
    mb_to_bytes,
    media_plan_signature,
    measure_media_ssim,
    next_target_media_budget,
    patch_output_pptx,
    referenced_ooxml_images,
    target_report_fields,
    write_markdown_report,
    write_target_skip_report,
    zip_member_size,
)

Logger = Callable[[str], None]
SUPPORTED_SUFFIXES = {".docx", ".docm"}


def _render_pdf_pages(pdf_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    document = pdfium.PdfDocument(str(pdf_path))
    paths: list[Path] = []
    try:
        for index in range(len(document)):
            path = output_dir / f"page-{index + 1}.png"
            document[index].render(scale=1.5).to_pil().save(path, format="PNG")
            paths.append(path)
    finally:
        document.close()
    return paths


def validate_render_layout(
    source: Path,
    output: Path,
    *,
    logger: Logger = print,
    page_ssim_threshold: float = 0.985,
) -> dict[str, Any]:
    source = source.expanduser().resolve()
    output = output.expanduser().resolve()
    rendered: list[Path] = []
    try:
        rendered = [
            convert_document_to_pdf(path, logger=logger) for path in (source, output)
        ]
        readers = [PdfReader(path) for path in rendered]
        geometries = [
            [
                (
                    round(float(page.mediabox.width), 3),
                    round(float(page.mediabox.height), 3),
                    int(page.rotation or 0) % 360,
                )
                for page in reader.pages
            ]
            for reader in readers
        ]
        texts = [
            [page.extract_text() or "" for page in reader.pages] for reader in readers
        ]
        if not geometries[0] or geometries[0] != geometries[1] or texts[0] != texts[1]:
            raise RuntimeError("Rendered DOCX page geometry or text changed")
        with tempfile.TemporaryDirectory(prefix="docx_render_validation_") as temp_dir:
            root = Path(temp_dir)
            pages = [
                _render_pdf_pages(path, root / str(index))
                for index, path in enumerate(rendered)
            ]
            if len(pages[0]) != len(pages[1]):
                raise RuntimeError("Rendered DOCX page count changed")
            scores: list[float] = []
            for original, candidate in zip(pages[0], pages[1], strict=True):
                with Image.open(original) as first, Image.open(candidate) as second:
                    if first.size != second.size:
                        raise RuntimeError("Rendered DOCX page dimensions changed")
                    width, height = first.size
                score = measure_media_ssim(
                    original,
                    candidate,
                    is_video=False,
                    width=width,
                    height=height,
                )
                edge_score, _ = image_detail_metrics(original, candidate)
                if score < page_ssim_threshold or edge_score < 0.98:
                    raise RuntimeError(
                        f"Rendered DOCX quality changed: SSIM={score:.4f}, "
                        f"edge={edge_score:.4f}"
                    )
                scores.append(score)
        return {
            "status": "passed",
            "page_count": len(geometries[0]),
            "minimum_page_ssim": round(min(scores), 6),
        }
    finally:
        for pdf in rendered:
            if pdf.parent.name.startswith("pptx_output_watermark_pdf_"):
                shutil.rmtree(pdf.parent, ignore_errors=True)


def _is_signed(archive: ZipFile) -> bool:
    names = set(archive.namelist())
    if any(name.lower().startswith("_xmlsignatures/") for name in names):
        return True
    for name in names:
        if not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(archive.read(name))
        except (ET.ParseError, KeyError):
            continue
        if any(
            "/digital-signature/" in relationship.attrib.get("Type", "").lower()
            for relationship in root
        ):
            return True
    return False


def _default_output_path(source: Path, target_size_mb: float, forced: bool) -> Path:
    label = f"{target_size_mb:.6f}".rstrip("0").rstrip(".").replace(".", "_")
    stem = f"{source.stem}_compressed_{label}MB"
    if forced:
        stem += "_forced"
    return source.with_name(f"{experimental_output_stem(stem)}{source.suffix.lower()}")


def default_output_path(
    source: Path, target_size_mb: float, *, forced: bool = False
) -> Path:
    """Public naming helper so the dispatcher can derive safe/forced outputs."""
    return _default_output_path(source, target_size_mb, forced)


def _validate_structure(
    source: Path,
    output: Path,
    replaced_media: set[str],
) -> None:
    with ZipFile(source) as original, ZipFile(output) as candidate:
        original_names = set(original.namelist())
        candidate_names = set(candidate.namelist())
        if original_names != candidate_names:
            raise RuntimeError("DOCX package members changed")
        for name in original_names - replaced_media:
            if original.read(name) != candidate.read(name):
                raise RuntimeError(f"DOCX non-media part changed: {name}")
        for macro_path in ("word/vbaProject.bin", "word/vbaData.xml"):
            if macro_path in original_names and original.read(
                macro_path
            ) != candidate.read(macro_path):
                raise RuntimeError(f"DOCM macro part changed: {macro_path}")


def _load_assets(source: Path, work_dir: Path) -> tuple[dict[str, ImageAsset], int]:
    assets: dict[str, ImageAsset] = {}
    with ZipFile(source) as archive:
        names = set(archive.namelist())
        if _is_signed(archive):
            raise RuntimeError(
                "Digitally signed Word documents are refused in safe compression"
            )
        referenced = {
            name
            for name in referenced_ooxml_images(archive, "word/", "word/media/")
            if name.startswith("word/media/") and name in names
        }
        media_bytes = 0
        originals = work_dir / "original_images"
        originals.mkdir(parents=True, exist_ok=True)
        for name in sorted(referenced):
            info = archive.getinfo(name)
            media_bytes += zip_member_size(info)
            extracted = originals / name
            extracted.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(name) as reader, extracted.open("wb") as writer:
                shutil.copyfileobj(reader, writer, length=1024 * 1024)
            asset = ImageAsset(
                media_path=name,
                zip_size=zip_member_size(info),
                extracted_path=str(extracted),
                output_media_path=name,
                max_area_ratio=1.0,
                max_width_ratio=1.0,
                max_height_ratio=1.0,
            )
            try:
                with Image.open(extracted) as image:
                    asset.width = image.width
                    asset.height = image.height
                    asset.display_width_px = image.width
                    asset.display_height_px = image.height
                    asset.image_format = str(image.format or "").upper()
                    asset.mode = image.mode
                    asset.content_type = classify_image_content(image)
            except (UnidentifiedImageError, OSError) as exc:
                asset.status = "unsupported"
                asset.reason = f"Cannot read image: {exc}"
            assets[name] = asset
    return assets, media_bytes


def compact_docx(
    source: Path,
    target_size_mb: float,
    *,
    output: Path | None = None,
    image_profile: str = "high",
    image_ssim_threshold: float = 0.99,
    forced: bool = False,
    safe_output: Path | None = None,
    confirm_forced: bool = False,
    validate_render: bool = True,
    logger: Logger = print,
) -> dict[str, Any]:
    source = source.expanduser().resolve()
    if source.suffix.lower() not in SUPPORTED_SUFFIXES or not source.is_file():
        raise ValueError(f"Expected an existing DOCX/DOCM file: {source}")
    try:
        with ZipFile(source) as archive:
            bad_member = archive.testzip()
            if bad_member is not None:
                raise RuntimeError(f"Corrupt Word package member: {bad_member}")
    except BadZipFile as exc:
        raise RuntimeError(
            "Encrypted or invalid Word packages are not processed without a safe re-encryption runtime"
        ) from exc

    target_bytes = mb_to_bytes(target_size_mb)
    if target_bytes >= source.stat().st_size:
        report_path = write_target_skip_report(
            source, target_size_mb, "Source Word document already meets the target"
        )
        return {
            "input": source,
            "output": source,
            "report_path": report_path,
            "skipped": True,
        }
    output = (
        output.expanduser().resolve()
        if output is not None
        else _default_output_path(source, target_size_mb, forced)
    )
    if output.suffix.lower() != source.suffix.lower():
        raise ValueError("Word output must preserve the source extension")
    if output == source:
        raise ValueError("Word output must not overwrite the source file")
    if output.exists():
        raise FileExistsError(f"Word output already exists: {output}")
    if forced:
        if not confirm_forced or safe_output is None:
            raise ValueError(
                "Forced Word compression requires an explicit confirmation and safe output"
            )
        safe_output = safe_output.expanduser().resolve()
        if not safe_output.is_file() or safe_output.stat().st_size <= target_bytes:
            raise ValueError(
                "Forced Word compression requires a safe output above target"
            )

    work_dir = Path(tempfile.mkdtemp(prefix="docx_compact_experimental_"))
    completed = False
    try:
        assets, media_bytes = _load_assets(source, work_dir)
        if not assets:
            raise RuntimeError("No safely referenced Word images were found")
        non_media_bytes = source.stat().st_size - media_bytes
        media_budget = (
            target_bytes
            - non_media_bytes
            - dynamic_package_reserve_bytes(target_bytes, non_media_bytes)
        )
        if media_budget <= 0:
            raise RuntimeError("Word non-media content already exceeds the target")

        encoded_dir = work_dir / "encoded_images"
        encoded_dir.mkdir()

        def encode_and_package(budget: int) -> None:
            _, image_budget = allocate_media_budgets(
                0, media_bytes, budget, "none", image_profile
            )
            assign_image_plan(
                assets,
                image_profile,
                max(1, image_budget),
                preserve_quality_fallbacks=True,
            )
            for index, asset in enumerate(assets.values()):
                asset_dir = encoded_dir / str(index)
                asset_dir.mkdir(exist_ok=True)
                encode_image_asset(asset, asset_dir)
            audit_encoded_assets(
                {},
                assets,
                video_threshold=0.95,
                image_threshold=image_ssim_threshold,
                forced=forced,
                logger=logger,
            )
            patch_output_pptx(
                source,
                output,
                {name: Path(asset.output_path) for name, asset in assets.items()},
            )
            _validate_structure(source, output, set(assets))

        encode_and_package(media_budget)
        attempts = [
            {
                "kind": "initial",
                "media_budget_bytes": media_budget,
                "actual_bytes": output.stat().st_size,
            }
        ]
        correction_rounds = 0
        giveback_used = False
        while True:
            next_attempt = next_target_media_budget(
                actual_bytes=output.stat().st_size,
                target_bytes=target_bytes,
                current_media_budget=media_budget,
                maximum_media_budget=media_bytes,
                correction_rounds=correction_rounds,
                giveback_used=giveback_used,
            )
            if next_attempt is None:
                break
            media_budget, kind = next_attempt
            correction_rounds += kind == "correction"
            giveback_used |= kind == "quality_giveback"
            previous_plan = media_plan_signature({}, assets)
            _, image_budget = allocate_media_budgets(
                0, media_bytes, media_budget, "none", image_profile
            )
            assign_image_plan(
                assets,
                image_profile,
                max(1, image_budget),
                preserve_quality_fallbacks=True,
            )
            if media_plan_signature({}, assets) == previous_plan:
                logger("Target capacity retry skipped: media plan is unchanged")
                break
            encode_and_package(media_budget)
            attempts.append(
                {
                    "kind": kind,
                    "media_budget_bytes": media_budget,
                    "actual_bytes": output.stat().st_size,
                }
            )

        render_validation = (
            validate_render_layout(source, output, logger=logger)
            if validate_render
            else {"status": "not_requested"}
        )
        report_path = output.with_suffix(".report.json")
        report = {
            "input_kind": "docm" if source.suffix.lower() == ".docm" else "docx",
            "input_pptx": str(source),
            "output_pptx": str(output),
            "target_size_mb": target_size_mb,
            "target": target_report_fields(target_size_mb, output),
            "presentation": {
                "quality_mode": "forced" if forced else "safe",
                "target_capacity_attempts": attempts,
                "non_media_parts_preserved": True,
                "macro_parts_preserved": source.suffix.lower() == ".docm",
                "render_validation": render_validation,
            },
            "videos": [],
            "images": [image_report_entry(asset) for asset in assets.values()],
        }
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        write_markdown_report(report_path, report)
        completed = True
        return {
            "input": source,
            "output": output,
            "report_path": report_path,
            "skipped": False,
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
        if not completed:
            output.unlink(missing_ok=True)
            output.with_suffix(".report.json").unlink(missing_ok=True)
            output.with_suffix(".report.md").unlink(missing_ok=True)
