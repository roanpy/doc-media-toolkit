#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from pptx_output_watermark.pptx_video_support import scan_embedded_videos
from pptx_tools.video_manager import VideoProject, probe_video, sha256_file


def _wmv_parts(path: Path) -> list[str]:
    with ZipFile(path) as archive:
        return [
            name
            for name in archive.namelist()
            if name.startswith("ppt/media/") and name.lower().endswith(".wmv")
        ]


def _anchors(asset: object) -> set[tuple[str, int]]:
    return {
        (occurrence.slide_path, occurrence.shape_id) for occurrence in asset.occurrences
    }


def _slide_xml(archive: ZipFile) -> dict[str, bytes]:
    return {
        name: archive.read(name)
        for name in archive.namelist()
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
    }


def _probe_member(
    archive: ZipFile,
    member: str,
    work: Path,
    cache: dict[str, dict[str, object]],
) -> tuple[str, dict[str, object]]:
    payload = archive.read(member)
    digest = hashlib.sha256(payload).hexdigest()
    if digest not in cache:
        target = work / f"{digest}{Path(member).suffix.lower()}"
        target.write_bytes(payload)
        cache[digest] = probe_video(target)
    return digest, cache[digest]


def validate_upgrade(
    project: VideoProject,
    deck: dict[str, object],
    source: Path,
    upgraded: Path,
    work: Path,
    probe_cache: dict[str, dict[str, object]],
) -> dict[str, object]:
    original_scan = scan_embedded_videos(source)
    upgraded_scan = scan_embedded_videos(upgraded)
    if len(original_scan) != len(upgraded_scan):
        raise RuntimeError(f"Video count changed: {source}")
    if any(Path(part).suffix.lower() == ".wmv" for part in upgraded_scan):
        raise RuntimeError(f"WMV remains in upgraded PPTX: {source}")

    old_by_anchor = {
        anchor: asset for asset in original_scan.values() for anchor in _anchors(asset)
    }
    new_by_anchor = {
        anchor: asset for asset in upgraded_scan.values() for anchor in _anchors(asset)
    }
    if set(old_by_anchor) != set(new_by_anchor):
        raise RuntimeError(f"Video shape anchors changed: {source}")
    family_by_anchor = {
        (occurrence["slide_path"], occurrence["shape_id"]): item["family_id"]
        for item in deck["assets"]
        for occurrence in item["occurrences"]
    }

    converted = 0
    with ZipFile(source) as old_zip, ZipFile(upgraded) as new_zip:
        if old_zip.testzip() is not None or new_zip.testzip() is not None:
            raise BadZipFile(f"PPTX ZIP validation failed: {source}")
        if _slide_xml(old_zip) != _slide_xml(new_zip):
            raise RuntimeError(f"Slide XML changed during media migration: {source}")

        checked_pairs: set[tuple[str, str]] = set()
        for anchor, old_asset in old_by_anchor.items():
            new_asset = new_by_anchor[anchor]
            pair = (old_asset.media_path, new_asset.media_path)
            if pair in checked_pairs:
                continue
            checked_pairs.add(pair)
            old_digest, old_meta = _probe_member(
                old_zip, old_asset.media_path, work, probe_cache
            )
            new_digest, new_meta = _probe_member(
                new_zip, new_asset.media_path, work, probe_cache
            )
            if Path(old_asset.media_path).suffix.lower() != ".wmv":
                if old_digest != new_digest:
                    raise RuntimeError(
                        f"Compatible media changed unexpectedly: {old_asset.media_path}"
                    )
                continue

            converted += 1
            if Path(new_asset.media_path).suffix.lower() != ".mp4":
                raise RuntimeError(f"WMV did not become MP4: {old_asset.media_path}")
            found = project.find_variant_by_hash(new_digest)
            expected_family = family_by_anchor.get(anchor)
            if found is None or found[0]["id"] != expected_family:
                raise RuntimeError(
                    f"Converted media is not registered to its family: {new_asset.media_path}"
                )
            if (old_meta.get("width"), old_meta.get("height")) != (
                new_meta.get("width"),
                new_meta.get("height"),
            ):
                raise RuntimeError(f"Video dimensions changed: {old_asset.media_path}")
            old_duration = float(old_meta.get("duration_sec") or 0)
            new_duration = float(new_meta.get("duration_sec") or 0)
            if abs(old_duration - new_duration) > max(0.35, old_duration * 0.01):
                raise RuntimeError(f"Video duration changed: {old_asset.media_path}")
            if bool(old_meta.get("has_audio")) != bool(new_meta.get("has_audio")):
                raise RuntimeError(
                    f"Video audio presence changed: {old_asset.media_path}"
                )
            if str(new_meta.get("video_codec", "")).lower() != "h264":
                raise RuntimeError(
                    f"Converted video is not H.264: {new_asset.media_path}"
                )
            if str(new_meta.get("audio_codec", "")).lower() not in {"", "aac"}:
                raise RuntimeError(
                    f"Converted audio is not AAC: {new_asset.media_path}"
                )
    return {"converted_parts": converted, "video_count": len(upgraded_scan)}


def _libreoffice_pages(path: Path, soffice: str, pdfinfo: str) -> int:
    with tempfile.TemporaryDirectory(prefix="wmv-lo-") as temp_dir:
        temp = Path(temp_dir)
        profile = temp / "profile"
        output = temp / "output"
        output.mkdir()
        completed = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output),
                str(path),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        pdf = output / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file():
            raise RuntimeError(
                f"LibreOffice could not open {path}: {completed.stderr.strip()}"
            )
        details = subprocess.run(
            [pdfinfo, str(pdf)], capture_output=True, text=True, check=True
        ).stdout
        return next(
            (
                int(line.split(":", 1)[1])
                for line in details.splitlines()
                if line.startswith("Pages:")
            ),
            0,
        )


def validate_with_libreoffice(source: Path, upgraded: Path) -> int:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pdfinfo = shutil.which("pdfinfo")
    if not soffice or not pdfinfo:
        raise RuntimeError("LibreOffice and pdfinfo are required for migration")
    source_pages = _libreoffice_pages(source, soffice, pdfinfo)
    upgraded_pages = _libreoffice_pages(upgraded, soffice, pdfinfo)
    if source_pages != upgraded_pages:
        raise RuntimeError(
            f"LibreOffice page count changed for {source}: "
            f"{source_pages} != {upgraded_pages}"
        )
    return upgraded_pages


def _atomic_copy(source: Path, target: Path) -> None:
    temporary = target.with_name(f".{target.name}.wmv-migration.tmp")
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("project", type=Path)
    parser.add_argument("scan_root", type=Path)
    parser.add_argument("output_root", type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    project = VideoProject.open(args.project)
    scan_root = args.scan_root.expanduser().resolve()
    output_root = args.output_root.expanduser().resolve()
    if output_root.exists():
        raise FileExistsError(output_root)
    originals = output_root / "原始PPTX"
    converted_dir = output_root / "转换后PPTX"
    output_root.mkdir(parents=True)
    shutil.copy2(project.manifest_path, output_root / "video-project.before.json")

    physical_wmv = []
    for path in scan_root.rglob("*.pptx"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        try:
            if _wmv_parts(path):
                physical_wmv.append(path.resolve())
        except BadZipFile:
            continue

    path_to_deck: dict[Path, dict[str, object]] = {}
    for deck in project.decks():
        paths = [project.deck_source_path(deck)] + [
            project.resolve_path(value) for value in deck.get("source_aliases", [])
        ]
        for path in paths:
            path_to_deck[path] = deck
    unregistered = [str(path) for path in physical_wmv if path not in path_to_deck]
    if unregistered:
        raise RuntimeError(f"WMV PPTX is not registered: {unregistered}")

    references_before = {
        family["id"]: sum(
            asset["family_id"] == family["id"]
            for deck in project.decks()
            for asset in deck["assets"]
        )
        for family in project.families()
    }
    converted_families = []
    for family in project.families():
        source_variant = project.source_variant(family)
        source_path = project.variant_path(source_variant)
        if source_path.suffix.lower() == ".wmv":
            wmv_variant = source_variant
            mp4_variant = project.create_unified_version(family["id"])
            project.activate_variant(mp4_variant["id"])
        elif source_variant.get("profile") == "1080p_source" and source_variant.get(
            "source_variant_id"
        ):
            _, wmv_variant = project.find_variant(source_variant["source_variant_id"])
            if project.variant_path(wmv_variant).suffix.lower() != ".wmv":
                continue
            mp4_variant = source_variant
        else:
            continue
        converted_families.append(
            {
                "family_id": family["id"],
                "name": family["name"],
                "wmv_variant_id": wmv_variant["id"],
                "mp4_variant_id": mp4_variant["id"],
                "mp4_path": str(project.variant_path(mp4_variant)),
                "references": references_before[family["id"]],
            }
        )

    grouped: dict[str, list[Path]] = {}
    for path in physical_wmv:
        grouped.setdefault(path_to_deck[path]["id"], []).append(path)

    records = []
    probe_cache: dict[str, dict[str, object]] = {}
    with tempfile.TemporaryDirectory(prefix="wmv-migration-probe-") as temp_dir:
        probe_work = Path(temp_dir)
        for index, (deck_id, paths) in enumerate(sorted(grouped.items()), start=1):
            deck = project.deck(deck_id)
            source = paths[0]
            original_hashes = {sha256_file(path) for path in paths}
            if len(original_hashes) != 1:
                raise RuntimeError(
                    f"Deck aliases no longer have identical bytes: {paths}"
                )
            for path in paths:
                relative = path.relative_to(scan_root)
                backup = originals / relative
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, backup)

            output = converted_dir / source.relative_to(scan_root)
            output.parent.mkdir(parents=True, exist_ok=True)
            result = project.upgrade_pptx_from_library(
                source,
                output_path=output,
                incompatible_only=True,
                progress_callback=lambda message, i=index, n=len(grouped): print(
                    f"[{i}/{n}] {message}", flush=True
                ),
            )
            if result["output_pptx"] is None:
                raise RuntimeError(f"No WMV media was upgraded: {source}")
            validation = validate_upgrade(
                project, deck, source, output, probe_work, probe_cache
            )
            pages = validate_with_libreoffice(source, output)
            for alias in paths[1:]:
                alias_output = converted_dir / alias.relative_to(scan_root)
                alias_output.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(output, alias_output)
            records.append(
                {
                    "deck_id": deck_id,
                    "sources": [str(path) for path in paths],
                    "backup_paths": [
                        str(originals / path.relative_to(scan_root)) for path in paths
                    ],
                    "converted_paths": [
                        str(converted_dir / path.relative_to(scan_root))
                        for path in paths
                    ],
                    "source_sha256": next(iter(original_hashes)),
                    "converted_sha256": sha256_file(output),
                    "libreoffice_pages": pages,
                    **validation,
                }
            )

    manifest_before_overwrite = output_root / "video-project.before-overwrite.json"
    shutil.copy2(project.manifest_path, manifest_before_overwrite)
    if args.overwrite:
        try:
            for record in records:
                sources = [Path(value) for value in record["sources"]]
                outputs = [Path(value) for value in record["converted_paths"]]
                for output, source in zip(outputs, sources, strict=True):
                    _atomic_copy(output, source)
                project.adopt_upgraded_deck_source(record["deck_id"], sources[0])
        except Exception:
            for record in records:
                sources = [Path(value) for value in record["sources"]]
                backups = [Path(value) for value in record["backup_paths"]]
                for backup, source in zip(backups, sources, strict=True):
                    _atomic_copy(backup, source)
            _atomic_copy(manifest_before_overwrite, project.manifest_path)
            raise

    project.reload()
    references_after = {
        family["id"]: sum(
            asset["family_id"] == family["id"]
            for deck in project.decks()
            for asset in deck["assets"]
        )
        for family in project.families()
    }
    if references_before != references_after:
        raise RuntimeError("Family reference counts changed during migration")
    remaining_wmv = [str(path) for path in physical_wmv if _wmv_parts(path)]
    if args.overwrite and remaining_wmv:
        raise RuntimeError(f"WMV remains after overwrite: {remaining_wmv}")

    report = {
        "project": str(project.root),
        "scan_root": str(scan_root),
        "overwritten": args.overwrite,
        "converted_families": converted_families,
        "pptx_files": sum(len(record["sources"]) for record in records),
        "registered_decks": len(records),
        "wmv_parts": sum(record["converted_parts"] for record in records),
        "records": records,
        "remaining_wmv": remaining_wmv,
        "references_preserved": references_before == references_after,
    }
    (output_root / "迁移报告.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    shutil.copy2(project.manifest_path, output_root / "video-project.after.json")
    print(
        json.dumps(
            {
                key: report[key]
                for key in (
                    "pptx_files",
                    "registered_decks",
                    "wmv_parts",
                    "remaining_wmv",
                    "references_preserved",
                )
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
