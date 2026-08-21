from __future__ import annotations

import copy
from collections import Counter, defaultdict
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pptx_tools.manager_i18n import tr

if TYPE_CHECKING:
    from pptx_tools.video_manager import VideoProject


def _issue(
    severity: str,
    code: str,
    message: str,
    **details: Any,
) -> dict[str, Any]:
    return {
        "severity": severity,
        "code": code,
        "message": message,
        **{key: value for key, value in details.items() if value not in (None, "")},
    }


def audit_video_project(
    project: VideoProject,
    *,
    verify_hashes: bool = False,
    progress_callback: Callable[[str], None] | None = None,
    cancel_callback: Callable[[], bool] | None = None,
) -> dict[str, Any]:
    """Return a read-only health report for one video library.

    Fast mode checks paths, sizes, mtimes, references, aliases, and quarantine
    metadata. Full mode additionally hashes every registered video entity.
    """
    issues: list[dict[str, Any]] = []
    reference_counts: Counter[str] = Counter()
    known_hash_owners: defaultdict[str, set[str]] = defaultdict(set)
    variant_hash_owners: defaultdict[str, set[str]] = defaultdict(set)
    variant_path_owners: defaultdict[str, list[str]] = defaultdict(list)
    tracked_paths: set[Path] = set()
    variant_count = 0
    available_count = 0
    missing_count = 0
    modified_count = 0
    metadata_drift_count = 0
    unreadable_count = 0

    for deck in project.decks():
        for asset in deck.get("assets", []):
            reference_counts[str(asset.get("family_id") or "")] += 1

    families = project.families()
    for family_index, family in enumerate(families, start=1):
        if cancel_callback is not None and cancel_callback():
            raise RuntimeError("Operation cancelled")
        if progress_callback is not None and (
            family_index == 1 or family_index == len(families) or family_index % 25 == 0
        ):
            progress_callback(
                f"{tr('正在核对视频族 ')}{family_index}/{len(families)}{tr('：')}{family['name']}"
            )
        family_id = family["id"]
        variants = family.get("variants", [])
        variant_count += len(variants)
        variant_ids = {variant["id"] for variant in variants}
        for key, label in (
            ("source_variant_id", tr("高清源")),
            ("active_variant_id", tr("当前版本")),
        ):
            if family.get(key) not in variant_ids:
                issues.append(
                    _issue(
                        "error",
                        "invalid_family_pointer",
                        f"{family['name']}{tr(' 的')}{label}{tr('不存在。')}",
                        family_id=family_id,
                    )
                )
        for digest in family.get("known_hashes", []):
            known_hash_owners[digest].add(family_id)
        for variant in variants:
            variant_id = variant["id"]
            known_hash_owners[variant["sha256"]].add(family_id)
            variant_hash_owners[variant["sha256"]].add(family_id)
            path = project.variant_path(variant)
            resolved = path.resolve()
            tracked_paths.add(resolved)
            variant_path_owners[str(resolved)].append(variant_id)
            status = project.status(variant)
            if status == "missing":
                missing_count += 1
                issues.append(
                    _issue(
                        "error",
                        "missing_variant",
                        f"{tr('视频文件不存在：')}{path}",
                        family_id=family_id,
                        variant_id=variant_id,
                        path=str(path),
                    )
                )
            elif verify_hashes:
                try:
                    project.require_variant_path(variant)
                except ValueError:
                    modified_count += 1
                    issues.append(
                        _issue(
                            "error",
                            "hash_mismatch",
                            f"{tr('视频文件哈希与清单不一致：')}{path}",
                            family_id=family_id,
                            variant_id=variant_id,
                            path=str(path),
                        )
                    )
                else:
                    available_count += 1
                    if status == "metadata_drift":
                        metadata_drift_count += 1
                        issues.append(
                            _issue(
                                "warning",
                                "variant_metadata_drift",
                                f"{tr('视频内容未变，但文件时间戳与清单不同：')}{path}",
                                family_id=family_id,
                                variant_id=variant_id,
                                path=str(path),
                            )
                        )
            elif status == "metadata_drift":
                metadata_drift_count += 1
                issues.append(
                    _issue(
                        "warning",
                        "variant_metadata_drift",
                        f"{tr('文件时间戳与清单不同，需执行哈希核验：')}{path}",
                        family_id=family_id,
                        variant_id=variant_id,
                        path=str(path),
                    )
                )
            elif status == "modified":
                modified_count += 1
                issues.append(
                    _issue(
                        "error",
                        "modified_variant",
                        f"{tr('视频文件大小在入库后发生变化：')}{path}",
                        family_id=family_id,
                        variant_id=variant_id,
                        path=str(path),
                    )
                )
            else:
                available_count += 1
            if variant.get("probe_error"):
                unreadable_count += 1
                severity = (
                    "error"
                    if variant_id
                    in {
                        family.get("source_variant_id"),
                        family.get("active_variant_id"),
                    }
                    else "warning"
                )
                issues.append(
                    _issue(
                        severity,
                        "unreadable_variant",
                        f"{tr('媒体元数据不可读：')}{path}",
                        family_id=family_id,
                        variant_id=variant_id,
                        path=str(path),
                    )
                )

    for digest, owners in known_hash_owners.items():
        if len(owners) > 1:
            issues.append(
                _issue(
                    "error",
                    "ambiguous_known_hash",
                    f"{tr('同一已知哈希属于 ')}{len(owners)}{tr(' 个视频族：')}{digest[:12]}",
                    sha256=digest,
                    family_ids=sorted(owners),
                )
            )
    for digest, owners in variant_hash_owners.items():
        if len(owners) > 1:
            issues.append(
                _issue(
                    "error",
                    "duplicate_variant_hash",
                    f"{tr('同一实体视频哈希跨 ')}{len(owners)}{tr(' 个视频族重复：')}{digest[:12]}",
                    sha256=digest,
                    family_ids=sorted(owners),
                )
            )
    for path, variant_ids in variant_path_owners.items():
        if len(variant_ids) > 1:
            issues.append(
                _issue(
                    "error",
                    "duplicate_variant_path",
                    f"{tr('同一视频文件路径被 ')}{len(variant_ids)}{tr(' 个版本共同占用：')}{path}",
                    path=path,
                    variant_ids=variant_ids,
                )
            )

    missing_deck_sources = 0
    stale_output_records = 0
    changed_output_records = 0
    output_record_count = 0
    decks = project.decks()
    for deck_index, deck in enumerate(decks, start=1):
        if cancel_callback is not None and cancel_callback():
            raise RuntimeError("Operation cancelled")
        if progress_callback is not None and (
            deck_index == 1 or deck_index == len(decks) or deck_index % 50 == 0
        ):
            progress_callback(
                f"{tr('正在核对 PPTX 关联 ')}{deck_index}/{len(decks)}{tr('：')}{deck['name']}"
            )
        source = project.deck_source_path(deck)
        aliases = [
            project.resolve_path(value) for value in deck.get("source_aliases", [])
        ]
        if not source.is_file():
            missing_deck_sources += 1
            alias_available = next((path for path in aliases if path.is_file()), None)
            issues.append(
                _issue(
                    "warning",
                    (
                        "deck_primary_missing_alias_available"
                        if alias_available
                        else "deck_source_missing"
                    ),
                    (
                        f"{tr('PPTX 主路径已失效，但存在可用别名：')}{alias_available}"
                        if alias_available
                        else f"{tr('PPTX 来源文件不存在：')}{source}"
                    ),
                    deck_id=deck["id"],
                    path=str(source),
                )
            )
        for kind in (
            "optimized_outputs",
            "detached_outputs",
            "restored_outputs",
        ):
            for record in deck.get(kind, []):
                output_record_count += 1
                path = project.resolve_path(record["path"])
                if not path.is_file():
                    stale_output_records += 1
                    issues.append(
                        _issue(
                            "info",
                            "missing_output_record",
                            f"{tr('历史输出已不存在：')}{path}",
                            deck_id=deck["id"],
                            output_id=record.get("id"),
                            output_kind=kind,
                            path=str(path),
                        )
                    )
                    continue
                expected_size = int(record.get("size_bytes") or 0)
                if expected_size and path.stat().st_size != expected_size:
                    changed_output_records += 1
                    issues.append(
                        _issue(
                            "warning",
                            "changed_output_record",
                            f"{tr('历史输出已被修改：')}{path}",
                            deck_id=deck["id"],
                            output_id=record.get("id"),
                            output_kind=kind,
                            path=str(path),
                        )
                    )

    media_root = project.root / "media"
    media_files = {path.resolve() for path in media_root.rglob("*") if path.is_file()}
    untracked_files = sorted(media_files - tracked_paths)
    for path in untracked_files:
        issues.append(
            _issue(
                "warning",
                "untracked_media",
                f"{tr('media 目录中存在未登记文件：')}{path}",
                path=str(path),
            )
        )

    pending_cleanup_count = 0
    try:
        pending_cleanup_count = len(project.pending_cleanup())
        for message in project.cleanup_pending_issues():
            issues.append(_issue("error", "cleanup_index_issue", message))
    except Exception as exc:
        issues.append(
            _issue(
                "error",
                "cleanup_index_invalid",
                f"{tr('待清理索引不可用：')}{exc}",
            )
        )

    issue_counts = Counter(item["code"] for item in issues)
    severity_counts = Counter(item["severity"] for item in issues)
    family_count = len(project.families())
    unlinked_count = sum(
        reference_counts.get(family["id"], 0) == 0 for family in project.families()
    )
    multi_version_count = sum(
        len(family.get("variants", [])) > 1 for family in project.families()
    )
    return {
        "project": {
            "name": project.data.get("name", project.root.name),
            "root": str(project.root),
            "project_id": project.data.get("project_id", ""),
            "revision": project.data.get("revision", 0),
        },
        "mode": "full_hash" if verify_hashes else "fast",
        "ok": severity_counts["error"] == 0,
        "stats": {
            "families": family_count,
            "variants": variant_count,
            "decks": len(project.decks()),
            "references": sum(reference_counts.values()),
            "unlinked_families": unlinked_count,
            "multi_version_families": multi_version_count,
            "available_variants": available_count,
            "missing_variants": missing_count,
            "modified_variants": modified_count,
            "metadata_drift_variants": metadata_drift_count,
            "unreadable_variants": unreadable_count,
            "output_records": output_record_count,
            "stale_output_records": stale_output_records,
            "changed_output_records": changed_output_records,
            "missing_deck_sources": missing_deck_sources,
            "untracked_media_files": len(untracked_files),
            "pending_cleanup_files": pending_cleanup_count,
            "errors": severity_counts["error"],
            "warnings": severity_counts["warning"],
            "info": severity_counts["info"],
        },
        "issue_counts": dict(sorted(issue_counts.items())),
        "issues": issues,
    }


def prune_missing_output_records(project: VideoProject) -> int:
    """Remove only records whose output files no longer exist.

    Deck hashes, source aliases, video families, media entities, and shape
    anchors are untouched.
    """
    original_data = copy.deepcopy(project.data)
    removed = 0
    for deck in project.decks():
        for key in (
            "optimized_outputs",
            "detached_outputs",
            "restored_outputs",
        ):
            records = deck.get(key, [])
            retained = [
                record
                for record in records
                if project.resolve_path(record["path"]).is_file()
            ]
            removed += len(records) - len(retained)
            deck[key] = retained
    if not removed:
        return 0
    try:
        project.save()
    except Exception:
        project.data = original_data
        raise
    project.record("stale_output_records_pruned", count=removed)
    return removed
