from __future__ import annotations

import hashlib
import io
import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Callable

import pypdfium2 as pdfium
import pikepdf
from PIL import Image, UnidentifiedImageError
from pypdf import PdfReader

from pptx_video_compactor import (
    ImageAsset,
    ImageOccurrence,
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
    target_report_fields,
    write_markdown_report,
    write_target_skip_report,
)

Logger = Callable[[str], None]
ImageKey = tuple[int, int]


def _default_output_path(source: Path, target_size_mb: float, forced: bool) -> Path:
    label = f"{target_size_mb:.6f}".rstrip("0").rstrip(".").replace(".", "_")
    stem = f"{source.stem}_compressed_{label}MB"
    if forced:
        stem += "_forced"
    return source.with_name(f"{experimental_output_stem(stem)}.pdf")


def default_output_path(
    source: Path, target_size_mb: float, *, forced: bool = False
) -> Path:
    """Public naming helper so the dispatcher can derive safe/forced outputs."""
    return _default_output_path(source, target_size_mb, forced)


def _is_signed(reader: PdfReader) -> bool:
    root = reader.trailer["/Root"]
    if root.get("/Perms") is not None:
        return True
    return any(
        str(field.get("/FT")) == "/Sig" and field.get("/V") is not None
        for field in (reader.get_fields() or {}).values()
    )


def _filters(image_object: Any) -> set[str]:
    value = image_object.get("/Filter")
    if value is None:
        return set()
    return {str(item) for item in value} if isinstance(value, list) else {str(value)}


def _safe_image_object(image_object: Any) -> tuple[bool, str]:
    if "/JBIG2Decode" in _filters(image_object):
        return False, "JBIG2 images are preserved; lossy JBIG2 is disabled"
    for key in (
        "/SMask",
        "/Mask",
        "/ImageMask",
        "/Alternates",
        "/Metadata",
        "/OC",
        "/OPI",
    ):
        if image_object.get(key) is not None:
            return False, f"Complex PDF image property preserved: {key}"
    color_space = image_object.get("/ColorSpace")
    if isinstance(color_space, list) or str(color_space) not in {
        "/DeviceGray",
        "/DeviceRGB",
        "/DeviceCMYK",
    }:
        return False, "ICCBased, indexed, separation, or unknown color space preserved"
    return True, ""


def _pdfimages_layout(
    source: Path, reader: PdfReader
) -> dict[ImageKey, list[dict[str, Any]]]:
    executable = shutil.which("pdfimages")
    if not executable:
        raise RuntimeError(
            "Poppler pdfimages is required for safe PDF image area classification"
        )
    result = subprocess.run(
        [executable, "-list", str(source)],
        capture_output=True,
        text=True,
        check=True,
    )
    layout: dict[ImageKey, list[dict[str, Any]]] = {}
    for line in result.stdout.splitlines():
        parts = line.split()
        if len(parts) < 16 or not parts[0].isdigit() or not parts[10].isdigit():
            continue
        try:
            page_index = int(parts[0]) - 1
            width = int(parts[3])
            height = int(parts[4])
            x_ppi = float(parts[12])
            y_ppi = float(parts[13])
            page = reader.pages[page_index]
            page_width = float(page.cropbox.width)
            page_height = float(page.cropbox.height)
            display_width = width * 72.0 / max(1.0, x_ppi)
            display_height = height * 72.0 / max(1.0, y_ppi)
        except (IndexError, TypeError, ValueError):
            continue
        layout.setdefault((int(parts[10]), int(parts[11])), []).append(
            {
                "page_index": page_index,
                "area_ratio": min(
                    1.0,
                    display_width * display_height / max(1.0, page_width * page_height),
                ),
                "width_ratio": min(1.0, display_width / max(1.0, page_width)),
                "height_ratio": min(1.0, display_height / max(1.0, page_height)),
                "display_width_px": max(1, round(display_width * 2)),
                "display_height_px": max(1, round(display_height * 2)),
            }
        )
    return layout


def _load_assets(
    source: Path, work_dir: Path
) -> tuple[
    dict[ImageKey, ImageAsset],
    int,
    list[dict[str, Any]],
    list[dict[str, Any]],
]:
    reader = PdfReader(source, strict=False)
    layout = _pdfimages_layout(source, reader)
    maximum_area_by_page: dict[int, float] = {}
    for occurrences in layout.values():
        for occurrence in occurrences:
            page_index = occurrence["page_index"]
            maximum_area_by_page[page_index] = max(
                maximum_area_by_page.get(page_index, 0.0),
                occurrence["area_ratio"],
            )
    page_analysis: list[dict[str, Any]] = []
    for page_index, page in enumerate(reader.pages):
        try:
            text_present = bool((page.extract_text() or "").strip())
        except Exception:
            text_present = False
        maximum_area = maximum_area_by_page.get(page_index, 0.0)
        page_analysis.append(
            {
                "page": page_index + 1,
                "kind": (
                    "mixed"
                    if maximum_area >= 0.75 and text_present
                    else "scanned"
                    if maximum_area >= 0.75
                    else "digital"
                ),
                "text_layer_preserved": text_present,
                "maximum_image_area_ratio": round(maximum_area, 6),
            }
        )
    assets: dict[ImageKey, ImageAsset] = {}
    skipped: list[dict[str, Any]] = []
    raw_bytes = 0
    originals = work_dir / "original_images"
    originals.mkdir()

    for page_index, page in enumerate(reader.pages):
        try:
            page_images = list(page.images)
        except (KeyError, TypeError, ValueError) as exc:
            skipped.append({"page": page_index + 1, "reason": str(exc)})
            continue
        for image_file in page_images:
            reference = image_file.indirect_reference
            if reference is None:
                skipped.append(
                    {
                        "page": page_index + 1,
                        "name": image_file.name,
                        "reason": "Inline image preserved",
                    }
                )
                continue
            key = (reference.idnum, reference.generation)
            if key in assets:
                continue
            image_object = reference.get_object()
            safe, reason = _safe_image_object(image_object)
            if not safe:
                skipped.append(
                    {"page": page_index + 1, "name": image_file.name, "reason": reason}
                )
                continue
            try:
                decoded = image_file.image.copy()
                decoded.load()
            except (UnidentifiedImageError, OSError, ValueError) as exc:
                skipped.append(
                    {
                        "page": page_index + 1,
                        "name": image_file.name,
                        "reason": f"Cannot decode image: {exc}",
                    }
                )
                continue
            if decoded.mode not in {"1", "L", "RGB", "CMYK"}:
                skipped.append(
                    {
                        "page": page_index + 1,
                        "name": image_file.name,
                        "reason": f"Image mode preserved: {decoded.mode}",
                    }
                )
                continue
            content_type = classify_image_content(decoded.convert("RGB"))
            if decoded.mode != "1" and content_type != "photo":
                skipped.append(
                    {
                        "page": page_index + 1,
                        "name": image_file.name,
                        "reason": f"Non-bilevel {content_type} preserved losslessly",
                    }
                )
                continue

            occurrences = layout.get(key) or [
                {
                    "page_index": page_index,
                    "area_ratio": 1.0,
                    "width_ratio": 1.0,
                    "height_ratio": 1.0,
                    "display_width_px": decoded.width,
                    "display_height_px": decoded.height,
                }
            ]
            reference_path = originals / f"{key[0]}_{key[1]}.png"
            decoded.convert("RGB" if decoded.mode == "CMYK" else decoded.mode).save(
                reference_path, format="PNG"
            )
            object_bytes = len(getattr(image_object, "_data", b""))
            if object_bytes <= 0:
                skipped.append(
                    {
                        "page": page_index + 1,
                        "name": image_file.name,
                        "reason": "PDF image stream size unavailable",
                    }
                )
                continue
            raw_bytes += object_bytes
            asset = ImageAsset(
                media_path=f"pdf/image_{key[0]}_{key[1]}.{'png' if decoded.mode == '1' else 'jpg'}",
                zip_size=object_bytes,
                occurrences=[
                    ImageOccurrence(
                        owner_path=f"page_{item['page_index'] + 1}",
                        slide_number=item["page_index"] + 1,
                        media_path=image_file.name,
                        area_ratio=item["area_ratio"],
                        width_ratio=item["width_ratio"],
                        height_ratio=item["height_ratio"],
                    )
                    for item in occurrences
                ],
                width=decoded.width,
                height=decoded.height,
                display_width_px=max(item["display_width_px"] for item in occurrences),
                display_height_px=max(
                    item["display_height_px"] for item in occurrences
                ),
                max_area_ratio=max(item["area_ratio"] for item in occurrences),
                max_width_ratio=max(item["width_ratio"] for item in occurrences),
                max_height_ratio=max(item["height_ratio"] for item in occurrences),
                content_type="line_art" if decoded.mode == "1" else content_type,
                image_format="PNG",
                mode=decoded.mode,
                extracted_path=str(reference_path),
                output_media_path=image_file.name,
            )
            assets[key] = asset
    return assets, raw_bytes, skipped, page_analysis


def _encoded_pdf_image_stream(image: Image.Image, quality: int) -> bytes:
    buffer = io.BytesIO()
    if image.mode != "1":
        image.save(buffer, "JPEG", quality=quality, optimize=True)
        return buffer.getvalue()
    image.save(buffer, "PDF")
    candidate = PdfReader(io.BytesIO(buffer.getvalue()))
    reference = candidate.pages[0].images[0].indirect_reference
    if reference is None:
        raise RuntimeError("Pillow produced an inline PDF image")
    return bytes(getattr(reference.get_object(), "_data", b""))


def _content_hash(page: Any) -> str:
    contents = page.get_contents()
    return hashlib.sha256(
        contents.get_data() if contents is not None else b""
    ).hexdigest()


def _annotation_fingerprint(reader: PdfReader) -> list[list[tuple[Any, ...]]]:
    result: list[list[tuple[Any, ...]]] = []
    for page in reader.pages:
        page_items: list[tuple[Any, ...]] = []
        for reference in page.get("/Annots", []):
            item = reference.get_object()
            page_items.append(
                (
                    str(item.get("/Subtype")),
                    tuple(float(value) for value in item.get("/Rect", [])),
                    str(item.get("/Contents", "")),
                    int(item.get("/F", 0)),
                    str(item.get("/T", "")),
                    str(item.get("/Dest", "")),
                    str((item.get("/A") or {}).get("/S", "")),
                    str((item.get("/A") or {}).get("/URI", "")),
                )
            )
        result.append(page_items)
    return result


def _field_fingerprint(reader: PdfReader) -> dict[str, tuple[str, ...]]:
    return {
        name: tuple(str(field.get(key, "")) for key in ("/FT", "/V", "/DV", "/Ff"))
        for name, field in (reader.get_fields() or {}).items()
    }


def _attachment_fingerprint(reader: PdfReader) -> dict[str, list[str]]:
    return {
        name: [hashlib.sha256(value).hexdigest() for value in values]
        for name, values in reader.attachments.items()
    }


def _outline_fingerprint(
    reader: PdfReader, items: list[Any] | None = None
) -> list[Any]:
    result: list[Any] = []
    for item in reader.outline if items is None else items:
        if isinstance(item, list):
            result.append(_outline_fingerprint(reader, item))
        else:
            try:
                page_number = reader.get_destination_page_number(item)
            except Exception:
                page_number = None
            result.append((str(getattr(item, "title", item)), page_number))
    return result


def _document_fingerprint(reader: PdfReader) -> dict[str, Any]:
    root = reader.trailer["/Root"]
    metadata_object = root.get("/Metadata")
    xmp_hash = (
        hashlib.sha256(metadata_object.get_object().get_data()).hexdigest()
        if metadata_object is not None
        else None
    )
    return {
        "page_geometry": [
            (
                tuple(float(value) for value in page.mediabox),
                tuple(float(value) for value in page.cropbox),
                tuple(float(value) for value in page.bleedbox),
                tuple(float(value) for value in page.trimbox),
                tuple(float(value) for value in page.artbox),
                int(page.rotation or 0) % 360,
            )
            for page in reader.pages
        ],
        "page_content": [_content_hash(page) for page in reader.pages],
        "text": [page.extract_text() or "" for page in reader.pages],
        "annotations": _annotation_fingerprint(reader),
        "fields": _field_fingerprint(reader),
        "attachments": _attachment_fingerprint(reader),
        "outline": _outline_fingerprint(reader),
        "page_labels": list(reader.page_labels),
        "metadata": dict(reader.metadata or {}),
        "xmp_hash": xmp_hash,
        "catalog_keys": sorted(str(key) for key in root.keys()),
        "catalog_settings": {
            key: str(root.get(key, ""))
            for key in (
                "/PageMode",
                "/PageLayout",
                "/Lang",
                "/ViewerPreferences",
                "/OpenAction",
                "/AA",
                "/MarkInfo",
                "/OutputIntents",
                "/OCProperties",
                "/Collection",
            )
        },
    }


def _validate_structure(source: Path, output: Path) -> None:
    before = PdfReader(source, strict=False)
    after = PdfReader(output, strict=False)
    if _document_fingerprint(before) != _document_fingerprint(after):
        raise RuntimeError("PDF non-image structure, text, or page content changed")


def _render_pages(pdf_path: Path, output_dir: Path, *, renderer: str) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    if renderer == "poppler":
        executable = shutil.which("pdftocairo")
        if not executable:
            raise RuntimeError("Poppler pdftocairo is required for PDF validation")
        prefix = output_dir / "page"
        subprocess.run(
            [executable, "-png", "-r", "144", str(pdf_path), str(prefix)],
            capture_output=True,
            text=True,
            check=True,
        )
        return sorted(output_dir.glob("page-*.png"))

    document = pdfium.PdfDocument(str(pdf_path))
    paths: list[Path] = []
    try:
        document.init_forms()
        for index in range(len(document)):
            path = output_dir / f"page-{index + 1}.png"
            document[index].render(scale=2).to_pil().save(path, format="PNG")
            paths.append(path)
    finally:
        document.close()
    return paths


def validate_render_layout(
    source: Path, output: Path, *, page_ssim_threshold: float = 0.985
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="pdf_render_validation_") as temp_dir:
        root = Path(temp_dir)
        scores: list[float] = []
        for renderer in ("poppler", "pdfium"):
            originals = _render_pages(
                source, root / renderer / "source", renderer=renderer
            )
            candidates = _render_pages(
                output, root / renderer / "output", renderer=renderer
            )
            if not originals or len(originals) != len(candidates):
                raise RuntimeError(f"{renderer} rendered PDF page count changed")
            for original, candidate in zip(originals, candidates, strict=True):
                with Image.open(original) as first, Image.open(candidate) as second:
                    if first.size != second.size:
                        raise RuntimeError(
                            f"{renderer} rendered PDF page dimensions changed"
                        )
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
                        f"{renderer} rendered PDF quality changed: "
                        f"SSIM={score:.4f}, edge={edge_score:.4f}"
                    )
                scores.append(score)
        return {
            "status": "passed",
            "renderers": ["poppler-cairo", "pdfium"],
            "page_count": len(scores) // 2,
            "minimum_page_ssim": round(min(scores), 6),
        }


def _write_pdf(
    source: Path,
    output: Path,
    assets: dict[ImageKey, ImageAsset],
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f".{output.stem}.", suffix=".pdf", dir=output.parent, delete=False
    ) as handle:
        temp = Path(handle.name)
    try:
        with pikepdf.open(source) as document:
            for key, asset in assets.items():
                if asset.status != "encoded" or asset.quality_status != "passed":
                    continue
                with (
                    Image.open(asset.extracted_path) as original,
                    Image.open(asset.output_path) as planned,
                ):
                    candidate = original.copy()
                    if planned.size != original.size:
                        candidate = candidate.resize(
                            planned.size, Image.Resampling.LANCZOS
                        )
                    encoded = _encoded_pdf_image_stream(candidate, asset.quality)
                    if len(encoded) >= asset.zip_size:
                        asset.status = "copied"
                        asset.reason = "PDF image replacement was not smaller"
                        asset.quality_status = "unchanged"
                        continue
                    image_object = document.get_object(key)
                    image_object["/Width"] = candidate.width
                    image_object["/Height"] = candidate.height
                    image_object["/BitsPerComponent"] = (
                        1 if candidate.mode == "1" else 8
                    )
                    image_object["/ColorSpace"] = pikepdf.Name(
                        "/DeviceGray"
                        if candidate.mode in {"1", "L"}
                        else "/DeviceCMYK"
                        if candidate.mode == "CMYK"
                        else "/DeviceRGB"
                    )
                    if "/Decode" in image_object:
                        del image_object["/Decode"]
                    if candidate.mode == "CMYK":
                        image_object["/Decode"] = pikepdf.Array([1, 0] * 4)
                    if candidate.mode == "1":
                        image_object.write(
                            encoded,
                            filter=pikepdf.Array([pikepdf.Name("/CCITTFaxDecode")]),
                            decode_parms=pikepdf.Array(
                                [
                                    pikepdf.Dictionary(
                                        {
                                            "/K": -1,
                                            "/BlackIs1": True,
                                            "/Columns": candidate.width,
                                            "/Rows": candidate.height,
                                        }
                                    )
                                ]
                            ),
                        )
                    else:
                        image_object.write(encoded, filter=pikepdf.Name("/DCTDecode"))
            document.save(temp)
        _validate_structure(source, temp)
        temp.replace(output)
    finally:
        temp.unlink(missing_ok=True)


def compact_pdf(
    source: Path,
    target_size_mb: float,
    *,
    output: Path | None = None,
    image_profile: str = "high",
    image_ssim_threshold: float = 0.99,
    forced: bool = False,
    safe_output: Path | None = None,
    confirm_forced: bool = False,
    logger: Logger = print,
) -> dict[str, Any]:
    source = source.expanduser().resolve()
    if source.suffix.lower() != ".pdf" or not source.is_file():
        raise ValueError(f"Expected an existing PDF file: {source}")
    reader = PdfReader(source, strict=False)
    if reader.is_encrypted:
        raise RuntimeError(
            "Encrypted PDFs are refused without safe encryption preservation"
        )
    if _is_signed(reader):
        raise RuntimeError("Digitally signed PDFs are refused in safe compression")

    target_bytes = mb_to_bytes(target_size_mb)
    if target_bytes >= source.stat().st_size:
        report_path = write_target_skip_report(
            source, target_size_mb, "Source PDF already meets the target"
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
    if output.suffix.lower() != ".pdf" or output == source:
        raise ValueError("PDF output must be a separate .pdf file")
    if output.exists():
        raise FileExistsError(f"PDF output already exists: {output}")
    if forced:
        if not confirm_forced or safe_output is None:
            raise ValueError(
                "Forced PDF compression requires an explicit confirmation and safe output"
            )
        safe_output = safe_output.expanduser().resolve()
        if not safe_output.is_file() or safe_output.stat().st_size <= target_bytes:
            raise ValueError(
                "Forced PDF compression requires a safe output above target"
            )

    work_dir = Path(tempfile.mkdtemp(prefix="pdf_compact_experimental_"))
    completed = False
    try:
        assets, image_bytes, skipped_images, page_analysis = _load_assets(
            source, work_dir
        )
        if not assets:
            raise RuntimeError("No safely replaceable PDF images were found")
        non_image_bytes = source.stat().st_size - image_bytes
        image_budget = (
            target_bytes
            - non_image_bytes
            - dynamic_package_reserve_bytes(target_bytes, non_image_bytes)
        )
        if image_budget <= 0:
            raise RuntimeError("PDF non-image content already exceeds the target")

        encoded_dir = work_dir / "encoded_images"
        encoded_dir.mkdir()

        def encode_and_package(budget: int) -> None:
            _, allocated = allocate_media_budgets(
                0, image_bytes, budget, "none", image_profile
            )
            assign_image_plan(
                assets,
                image_profile,
                max(1, allocated),
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
                preserve_image_metadata=False,
                logger=logger,
            )
            _write_pdf(source, output, assets)

        encode_and_package(image_budget)
        attempts = [
            {
                "kind": "initial",
                "media_budget_bytes": image_budget,
                "actual_bytes": output.stat().st_size,
            }
        ]
        correction_rounds = 0
        giveback_used = False
        while True:
            next_attempt = next_target_media_budget(
                actual_bytes=output.stat().st_size,
                target_bytes=target_bytes,
                current_media_budget=image_budget,
                maximum_media_budget=image_bytes,
                correction_rounds=correction_rounds,
                giveback_used=giveback_used,
            )
            if next_attempt is None:
                break
            image_budget, kind = next_attempt
            correction_rounds += kind == "correction"
            giveback_used |= kind == "quality_giveback"
            previous_plan = media_plan_signature({}, assets)
            _, allocated = allocate_media_budgets(
                0, image_bytes, image_budget, "none", image_profile
            )
            assign_image_plan(
                assets,
                image_profile,
                max(1, allocated),
                preserve_quality_fallbacks=True,
            )
            if media_plan_signature({}, assets) == previous_plan:
                logger("Target capacity retry skipped: media plan is unchanged")
                break
            output.unlink(missing_ok=True)
            encode_and_package(image_budget)
            attempts.append(
                {
                    "kind": kind,
                    "media_budget_bytes": image_budget,
                    "actual_bytes": output.stat().st_size,
                }
            )

        render_validation = validate_render_layout(source, output)
        report_path = output.with_suffix(".report.json")
        report = {
            "input_kind": "pdf",
            "input_pptx": str(source),
            "output_pptx": str(output),
            "target_size_mb": target_size_mb,
            "target": target_report_fields(target_size_mb, output),
            "presentation": {
                "quality_mode": "forced" if forced else "safe",
                "target_capacity_attempts": attempts,
                "non_image_structure_preserved": True,
                "render_validation": render_validation,
                "skipped_images": skipped_images,
                "page_analysis": page_analysis,
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
