#!/usr/bin/env python3

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from pathlib import Path
from zipfile import ZipFile

from pptx_output_watermark.pptx_video_support import scan_embedded_videos
from pptx_tools.video_manager import VideoProject, sha256_file
from pptx_video_compactor import (
    VideoAsset,
    choose_output_media_path,
    media_needs_mp4,
    patch_output_pptx,
)


MIN_VIDEO_SAVINGS = 0.15
MIN_QUALITY_THRESHOLD = 0.90
POLICY = "ssim095-video-15pct-and-pptx-5pct-or-10mib-v4"


def policy_for_threshold(threshold: float) -> str:
    return f"ssim{round(threshold * 100):03d}-video-15pct-and-pptx-5pct-or-10mib-v4"


def save_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(f"{path.suffix}.tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temporary.replace(path)


def member_sha256(archive: ZipFile, name: str) -> str:
    digest = hashlib.sha256()
    with archive.open(name) as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def deck_for_path(project: VideoProject, path: Path) -> dict:
    encoded = project.encode_path(path)
    matches = [
        deck
        for deck in project.decks()
        if encoded == deck["source_path"] or encoded in deck.get("source_aliases", [])
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one registered PPTX record for {path}")
    return matches[0]


def inventory(project: VideoProject, source_root: Path) -> list[dict]:
    targets: list[dict] = []
    for path in sorted(source_root.rglob("*.pptx")):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        scanned = scan_embedded_videos(path)
        if not scanned:
            continue
        assets = []
        with ZipFile(path) as archive:
            for media_path, asset in sorted(scanned.items()):
                digest = member_sha256(archive, media_path)
                family = project.family_by_known_hash(digest)
                if family is None:
                    raise RuntimeError(f"Unmatched media {media_path} in {path}")
                assets.append(
                    {
                        "part": media_path,
                        "sha256": digest,
                        "family_id": family["id"],
                        "anchors": sorted(
                            [item.slide_path, item.shape_id]
                            for item in asset.occurrences
                        ),
                    }
                )
        deck = deck_for_path(project, path)
        targets.append(
            {
                "source": str(path),
                "relative": path.relative_to(source_root).as_posix(),
                "source_sha256": sha256_file(path),
                "source_size_bytes": path.stat().st_size,
                "deck_id": deck["id"],
                "assets": assets,
            }
        )
    return targets


def valid_variant(project: VideoProject, variant: dict) -> bool:
    try:
        project.require_variant_path(variant)
    except (FileNotFoundError, ValueError):
        return False
    return True


def compatibility_errors(source: dict, compact: dict) -> list[str]:
    errors = []
    source_duration = float(source.get("duration_sec") or 0)
    compact_duration = float(compact.get("duration_sec") or 0)
    if (
        source_duration <= 0
        or compact_duration <= 0
        or abs(source_duration - compact_duration) > max(0.35, source_duration * 0.002)
    ):
        errors.append(f"时长不一致：{source_duration:g} -> {compact_duration:g} 秒")
    source_width = int(source.get("width") or 0)
    source_height = int(source.get("height") or 0)
    compact_width = int(compact.get("width") or 0)
    compact_height = int(compact.get("height") or 0)
    if (
        min(source_width, source_height, compact_width, compact_height) <= 0
        or abs(source_width / source_height - compact_width / compact_height)
        > max(source_width / source_height, compact_width / compact_height) * 0.015
    ):
        errors.append("宽高比不一致")
    if bool(source.get("has_audio")) != bool(compact.get("has_audio")):
        errors.append("音轨存在性不一致")
    if str(compact.get("video_codec") or "").lower() not in {"h264", "avc1"}:
        errors.append("视频不是 H.264")
    if (
        compact.get("has_audio")
        and str(compact.get("audio_codec") or "").lower() != "aac"
    ):
        errors.append("音频不是 AAC")
    return errors


def should_replace_video(
    media_path: str,
    current_size_bytes: int,
    replacement_size_bytes: int,
    *,
    compatible: bool,
) -> bool:
    if not compatible:
        return False
    return media_needs_mp4(media_path) or replacement_size_bytes < (
        current_size_bytes * (1 - MIN_VIDEO_SAVINGS)
    )


def source_matches_report(target: dict, existing: dict | None) -> bool:
    return bool(
        existing
        and existing.get("source_sha256") == target.get("source_sha256")
        and existing.get("source_size_bytes") == target.get("source_size_bytes")
    )


def target_needs_build(
    target: dict,
    existing: dict | None,
    output: Path,
    *,
    policy: str = POLICY,
) -> bool:
    if not source_matches_report(target, existing) or not existing:
        return True
    if existing.get("policy") != policy:
        return True
    if existing.get("status") == "skipped":
        return False
    return not (
        existing.get("status") == "validated"
        and output.is_file()
        and sha256_file(output) == existing.get("output_sha256")
    )


def compressed_variants(
    project: VideoProject,
    family_ids: set[str],
    profile: str,
    delivery_root: Path,
    report: dict,
    report_path: Path,
) -> tuple[dict[str, dict], dict[str, str]]:
    results: dict[str, dict] = {}
    blocked: dict[str, str] = {}
    total = len(family_ids)
    for index, family_id in enumerate(sorted(family_ids), start=1):
        family = project.family(family_id)
        source = project.source_variant(family)
        family_report = report["families"].get(family_id, {})
        reported_variant_id = family_report.get(
            "selected_variant_id"
        ) or family_report.get("compressed_variant_id")
        variant = next(
            (
                item
                for item in family["variants"]
                if item["id"] == reported_variant_id
                and (
                    item["id"] == source["id"]
                    or item.get("source_variant_id") == source["id"]
                )
                and valid_variant(project, item)
            ),
            None,
        )
        if variant is None:
            variant = next(
                (
                    item
                    for item in family["variants"]
                    if item.get("profile") == profile
                    and item.get("source_variant_id") == source["id"]
                    and valid_variant(project, item)
                ),
                None,
            )
        if variant is None:
            print(f"[{index}/{total}] 压缩视频族：{family['name']}", flush=True)
            variant = project.compress_variant(source["id"], profile, activate=False)
        else:
            print(f"[{index}/{total}] 复用压缩版本：{family['name']}", flush=True)
        source_path = project.require_variant_path(variant)
        category = Path(str(family.get("category") or "未分类"))
        delivery = delivery_root / category / f"{family_id[:8]}_{source_path.name}"
        delivery.parent.mkdir(parents=True, exist_ok=True)
        if not delivery.is_file() or sha256_file(delivery) != variant["sha256"]:
            temporary = delivery.with_suffix(f"{delivery.suffix}.tmp")
            shutil.copy2(source_path, temporary)
            if sha256_file(temporary) != variant["sha256"]:
                temporary.unlink(missing_ok=True)
                raise RuntimeError(f"Delivery video hash mismatch: {delivery}")
            temporary.replace(delivery)
        results[family_id] = variant
        errors = compatibility_errors(source, variant)
        if errors:
            blocked[family_id] = "压缩版本兼容性验证失败"
            print(f"  阻止回填：{'；'.join(errors)}", flush=True)
        report["families"][family_id] = {
            "name": family["name"],
            "source_variant_id": source["id"],
            "compressed_variant_id": variant["id"],
            "compressed_sha256": variant["sha256"],
            "managed_path": str(source_path),
            "delivery_path": str(delivery),
            "compatibility_errors": errors,
        }
        save_json(report_path, report)
    return results, blocked


def apply_quality_selection(
    project: VideoProject,
    variants: dict[str, dict],
    blocked: dict[str, str],
    selection: dict,
    delivery_root: Path,
    report: dict,
) -> None:
    threshold = float(selection.get("threshold") or 0)
    if threshold < MIN_QUALITY_THRESHOLD or threshold > 1.0:
        raise RuntimeError(
            f"Quality selection must use SSIM {MIN_QUALITY_THRESHOLD:.2f}-1.00"
        )
    selected_items = selection.get("items")
    if not isinstance(selected_items, dict):
        raise RuntimeError("Quality selection is missing family decisions")
    for family_id in variants:
        family = project.family(family_id)
        source = project.source_variant(family)
        decision = selected_items.get(family_id)
        if not isinstance(decision, dict):
            blocked[family_id] = "视频尚未完成画质评估"
            continue
        if decision.get("source_sha256") != source["sha256"]:
            raise RuntimeError(f"Quality selection source changed: {family['name']}")
        selected_profile = str(decision.get("selected_profile") or "")
        selected_source = (
            decision.get("selected_variant_id") == source["id"]
            and decision.get("selected_sha256") == source["sha256"]
        )
        if selected_profile == "source" or selected_source:
            blocked[family_id] = f"所有压缩档位均未达到 SSIM {threshold:.2f}"
            report["families"][family_id].update(
                {"selected_profile": "source", "selected_ssim": 1.0}
            )
            continue
        if selected_profile not in {"aggressive", "balanced", "high"}:
            raise RuntimeError(f"Invalid quality profile for {family['name']}")
        score = decision.get("ssim")
        if not isinstance(score, (int, float)) or float(score) < threshold:
            raise RuntimeError(f"Invalid quality score for {family['name']}")
        selected = next(
            (
                item
                for item in family["variants"]
                if item["id"] == decision.get("selected_variant_id")
                and item["sha256"] == decision.get("selected_sha256")
                and item.get("profile") == selected_profile
                and item.get("source_variant_id") == source["id"]
                and valid_variant(project, item)
            ),
            None,
        )
        if selected is None:
            raise RuntimeError(f"Selected quality variant is missing: {family['name']}")
        errors = compatibility_errors(source, selected)
        if errors:
            blocked[family_id] = "压缩版本兼容性验证失败"
            continue
        variants[family_id] = selected
        category = Path(str(family.get("category") or "未分类"))
        selected_path = project.require_variant_path(selected)
        delivery = delivery_root / category / f"{family_id[:8]}_{selected_path.name}"
        delivery.parent.mkdir(parents=True, exist_ok=True)
        if not delivery.is_file() or sha256_file(delivery) != selected["sha256"]:
            temporary = delivery.with_suffix(f"{delivery.suffix}.tmp")
            shutil.copy2(selected_path, temporary)
            if sha256_file(temporary) != selected["sha256"]:
                temporary.unlink(missing_ok=True)
                raise RuntimeError(f"Quality delivery hash mismatch: {delivery}")
            temporary.replace(delivery)
        report["families"][family_id].update(
            {
                "selected_profile": selected_profile,
                "selected_ssim": float(score),
                "selected_variant_id": selected["id"],
                "selected_sha256": selected["sha256"],
                "quality_delivery_path": str(delivery),
            }
        )


def validate_output(
    source: Path,
    output: Path,
    assets: list[dict],
    decisions: dict[str, dict],
) -> None:
    expected = {
        decisions[item["part"]]["part"]: (
            item["family_id"],
            decisions[item["part"]]["sha256"],
            item["anchors"],
        )
        for item in assets
    }
    output_scan = scan_embedded_videos(output)
    if set(output_scan) != set(expected):
        raise RuntimeError(f"Media part set changed unexpectedly: {output}")
    with ZipFile(source) as before, ZipFile(output) as after:
        if after.testzip() is not None:
            raise RuntimeError(f"Invalid output ZIP: {output}")
        source_parts = {item["part"] for item in assets}
        output_media = set(expected)
        changed_container = any(
            old != decision["part"] for old, decision in decisions.items()
        )
        allowed_xml = {"[Content_Types].xml"} if changed_container else set()
        for name in set(before.namelist()) - source_parts:
            if name in allowed_xml or (changed_container and name.endswith(".rels")):
                continue
            if name not in after.namelist():
                raise RuntimeError(f"Non-media member removed: {name}")
            left = before.getinfo(name)
            right = after.getinfo(name)
            if (left.file_size, left.CRC) != (right.file_size, right.CRC):
                raise RuntimeError(f"Non-media member changed: {name}")
        for part, asset in output_scan.items():
            family_id, digest, anchors = expected[part]
            if member_sha256(after, part) != digest:
                raise RuntimeError(f"Compressed media hash mismatch: {part}")
            actual_anchors = sorted(
                [item.slide_path, item.shape_id] for item in asset.occurrences
            )
            if actual_anchors != anchors:
                raise RuntimeError(f"Video anchors changed: {part} / {family_id}")
        unexpected = (
            set(after.namelist())
            - output_media
            - (set(before.namelist()) - source_parts)
        )
        if unexpected:
            raise RuntimeError(f"Unexpected package members: {sorted(unexpected)}")


def build_pptx_outputs(
    project: VideoProject,
    targets: list[dict],
    variants: dict[str, dict],
    blocked_families: dict[str, str],
    output_root: Path,
    report: dict,
    report_path: Path,
    *,
    policy: str = POLICY,
) -> None:
    total = len(targets)
    for index, target in enumerate(targets, start=1):
        source = Path(target["source"])
        output = output_root / target["relative"]
        existing = report["files"].get(target["relative"])
        if not target_needs_build(target, existing, output, policy=policy):
            continue
        print(f"[{index}/{total}] 回填 PPTX：{target['relative']}", flush=True)
        replacements: dict[str, Path] = {}
        relationship_map: dict[str, str] = {}
        replacement_infos = {}
        remove_paths: set[str] = set()
        video_assets: dict[str, VideoAsset] = {}
        decisions: dict[str, dict] = {}
        with ZipFile(source) as archive:
            reserved = set(archive.namelist())
            for asset in target["assets"]:
                old_part = asset["part"]
                source_info = archive.getinfo(old_part)
                variant = variants[asset["family_id"]]
                compatible = asset["family_id"] not in blocked_families
                should_replace = should_replace_video(
                    old_part,
                    source_info.file_size,
                    int(variant["size_bytes"]),
                    compatible=compatible,
                )
                if not should_replace:
                    decisions[old_part] = {
                        "part": old_part,
                        "sha256": asset["sha256"],
                        "replaced": False,
                        "source_size_bytes": source_info.file_size,
                        "replacement_size_bytes": int(variant["size_bytes"]),
                        "reason": (
                            blocked_families[asset["family_id"]]
                            if not compatible
                            else "压缩收益不足 15%"
                        ),
                    }
                    continue
                new_part = choose_output_media_path(old_part, reserved)
                decisions[old_part] = {
                    "part": new_part,
                    "sha256": variant["sha256"],
                    "replaced": True,
                    "source_size_bytes": source_info.file_size,
                    "replacement_size_bytes": int(variant["size_bytes"]),
                }
                replacements[new_part] = project.require_variant_path(variant)
                replacement_infos[new_part] = source_info
                reserved.add(new_part)
                if new_part != old_part:
                    relationship_map[old_part] = new_part
                    remove_paths.add(old_part)
                    video_assets[old_part] = VideoAsset(
                        media_path=old_part,
                        zip_size=archive.getinfo(old_part).file_size,
                        output_media_path=new_part,
                    )
        if not replacements:
            output.unlink(missing_ok=True)
            report["files"][target["relative"]] = {
                **target,
                "policy": policy,
                "decisions": decisions,
                "status": "skipped",
                "reason": "没有视频达到至少 15% 的压缩收益",
            }
            save_json(report_path, report)
            continue
        patch_output_pptx(
            source,
            output,
            replacements,
            relationship_path_map=relationship_map,
            replacement_infos=replacement_infos,
            remove_paths=remove_paths,
            video_assets=video_assets,
        )
        validate_output(source, output, target["assets"], decisions)
        output_digest = sha256_file(output)
        saved_bytes = target["source_size_bytes"] - output.stat().st_size
        reduction = saved_bytes / target["source_size_bytes"]
        if reduction < 0.05 and saved_bytes < 10 * 1024 * 1024:
            candidate_size = output.stat().st_size
            output.unlink()
            report["files"][target["relative"]] = {
                **target,
                "policy": policy,
                "decisions": decisions,
                "candidate_output_size_bytes": candidate_size,
                "status": "skipped",
                "reason": "整份 PPTX 节省不足 5% 且不足 10 MiB",
            }
            save_json(report_path, report)
            continue
        project.register_optimized_output(source, output, output_digest)
        report["files"][target["relative"]] = {
            **target,
            "policy": policy,
            "output": str(output),
            "output_sha256": output_digest,
            "output_size_bytes": output.stat().st_size,
            "decisions": decisions,
            "status": "validated",
        }
        save_json(report_path, report)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build deduplicated compact-video PPTX delivery outputs."
    )
    parser.add_argument("--project", required=True, type=Path)
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--output-root", required=True, type=Path)
    parser.add_argument(
        "--profile", choices=("high", "balanced", "aggressive"), default="aggressive"
    )
    parser.add_argument(
        "--quality-selection",
        type=Path,
        help="JSON decisions produced by the SSIM fallback audit",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project = VideoProject.open(args.project)
    output_root = args.output_root.expanduser().resolve()
    report_path = output_root / "00_报告" / "batch-report.json"
    report = (
        json.loads(report_path.read_text(encoding="utf-8"))
        if report_path.is_file()
        else {
            "version": 1,
            "project": str(project.root),
            "source_root": str(args.source_root.expanduser().resolve()),
            "output_root": str(output_root),
            "profile": args.profile,
            "families": {},
            "files": {},
        }
    )
    if report["profile"] != args.profile:
        raise RuntimeError("Existing batch report uses a different profile")
    source_root = args.source_root.expanduser().resolve()
    if Path(report["project"]).resolve() != project.root:
        raise RuntimeError("Existing batch report uses a different video project")
    if Path(report["source_root"]).resolve() != source_root:
        raise RuntimeError("Existing batch report uses a different source root")
    targets = inventory(project, source_root)
    report["target_count"] = len(targets)
    report["source_size_bytes"] = sum(item["source_size_bytes"] for item in targets)
    save_json(report_path, report)
    selection_path = (
        args.quality_selection.expanduser().resolve()
        if args.quality_selection
        else output_root / "00_报告" / "quality-fallback-selection.json"
    )
    if not selection_path.is_file():
        raise RuntimeError(f"Quality selection report is required: {selection_path}")
    selection = json.loads(selection_path.read_text(encoding="utf-8"))
    threshold = float(selection.get("threshold") or 0)
    if threshold < MIN_QUALITY_THRESHOLD or threshold > 1.0:
        raise RuntimeError(
            f"Quality selection must use SSIM {MIN_QUALITY_THRESHOLD:.2f}-1.00"
        )
    policy = policy_for_threshold(threshold)
    report["quality_threshold"] = threshold
    pending_targets = [
        target
        for target in targets
        if target_needs_build(
            target,
            report["files"].get(target["relative"]),
            output_root / "03_压缩PPTX" / target["relative"],
            policy=policy,
        )
    ]
    for target in targets:
        existing = report["files"].get(target["relative"])
        if (
            source_matches_report(target, existing)
            and existing.get("policy") == policy
            and existing.get("status") == "skipped"
        ):
            (output_root / "03_压缩PPTX" / target["relative"]).unlink(missing_ok=True)
    family_ids = {
        asset["family_id"] for target in pending_targets for asset in target["assets"]
    }
    pending_cleanup = project.pending_cleanup()
    for family_id in family_ids:
        decision = selection.get("items", {}).get(family_id, {})
        variant_id = decision.get("selected_variant_id")
        if variant_id and not any(
            item["id"] == variant_id for item in project.family(family_id)["variants"]
        ):
            entry = next(
                (
                    item
                    for item in pending_cleanup
                    if item.get("family_id") == family_id
                    and item.get("variant", {}).get("id") == variant_id
                ),
                None,
            )
            if entry:
                project.restore_cleanup_entry(entry["token"])
    variants, blocked_families = compressed_variants(
        project,
        family_ids,
        args.profile,
        output_root / "02_压缩视频",
        report,
        report_path,
    )
    apply_quality_selection(
        project,
        variants,
        blocked_families,
        selection,
        output_root / "02_压缩视频-画质达标",
        report,
    )
    build_pptx_outputs(
        project,
        pending_targets,
        variants,
        blocked_families,
        output_root / "03_压缩PPTX",
        report,
        report_path,
        policy=policy,
    )
    report["status"] = "validated"
    report["output_size_bytes"] = sum(
        item.get("output_size_bytes", 0) for item in report["files"].values()
    )
    save_json(report_path, report)
    print(f"完成：{report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
