from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import logging
import os
import posixpath
import shutil
import tempfile
import uuid
from collections.abc import Callable, Sequence
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from zipfile import BadZipFile, ZipFile

from PIL import Image, ImageOps, UnidentifiedImageError
from pypdf import PdfReader
from pypdf.errors import PdfReadError

from pptx_tools.project_lock import project_write_lock


SCHEMA_VERSION = 1
MANIFEST_NAME = "image-project.json"
BACKUP_MANIFEST_NAME = "image-project.json.bak"
LOCK_NAME = ".image-project.lock"
CLEANUP_DIR_NAME = "_cleanup"
CLEANUP_INDEX_NAME = "index.json"
CLEANUP_LOCK_NAME = ".image-cleanup.lock"
LOGGER = logging.getLogger("pptx_tools.image_manager")
SUPPORTED_IMAGE_SUFFIXES = {
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
OFFICE_MEDIA_PREFIXES = {
    ".pptx": "ppt/media/",
    ".docx": "word/media/",
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _image_metadata(data: bytes) -> dict[str, Any]:
    try:
        with Image.open(io.BytesIO(data)) as image:
            image.load()
            image_format = str(image.format or "").upper()
            image = ImageOps.exif_transpose(image)
            rgba = image.convert("RGBA")
            grayscale = image.convert("L").resize((9, 8), Image.Resampling.LANCZOS)
            pixels = list(grayscale.tobytes())
            bits = 0
            for row in range(8):
                offset = row * 9
                for column in range(8):
                    bits = (bits << 1) | int(
                        pixels[offset + column] > pixels[offset + column + 1]
                    )
            return {
                "width": int(image.width),
                "height": int(image.height),
                "format": image_format,
                "mode": str(image.mode),
                "dhash": f"{bits:016x}",
                "pixel_sha256": hashlib.sha256(
                    f"{image.width}x{image.height}:".encode() + rgba.tobytes()
                ).hexdigest(),
            }
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise ValueError(f"不支持或损坏的图片：{exc}") from exc


def _safe_suffix(name: str, image_format: str) -> str:
    suffix = Path(name).suffix.lower()
    if suffix in SUPPORTED_IMAGE_SUFFIXES:
        return ".jpg" if suffix == ".jpeg" else suffix
    return {
        "JPEG": ".jpg",
        "PNG": ".png",
        "GIF": ".gif",
        "BMP": ".bmp",
        "TIFF": ".tiff",
        "WEBP": ".webp",
    }.get(image_format.upper(), ".bin")


def _office_relationship_owner(rels_path: str) -> str | None:
    if rels_path == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rels_path or not rels_path.endswith(".rels"):
        return None
    prefix, filename = rels_path.rsplit(marker, 1)
    return f"{prefix}/{filename[:-5]}"


def _referenced_office_images(archive: ZipFile, media_prefix: str) -> list[str]:
    import xml.etree.ElementTree as ET

    members = set(archive.namelist())
    referenced: set[str] = set()
    for rels_path in members:
        owner = _office_relationship_owner(rels_path)
        if owner is None:
            continue
        try:
            root = ET.fromstring(archive.read(rels_path))
        except (ET.ParseError, KeyError):
            continue
        owner_dir = posixpath.dirname(owner)
        for relationship in root:
            if relationship.attrib.get("TargetMode") == "External":
                continue
            target = relationship.attrib.get("Target", "")
            rel_type = relationship.attrib.get("Type", "")
            if not target or not (
                rel_type.endswith("/image") or "media/" in target.replace("\\", "/")
            ):
                continue
            resolved = posixpath.normpath(
                posixpath.join(owner_dir, target.replace("\\", "/"))
            ).lstrip("/")
            if resolved.startswith(media_prefix) and resolved in members:
                referenced.add(resolved)
    return sorted(referenced)


def _hamming_hex(left: str, right: str) -> int:
    return (int(left, 16) ^ int(right, 16)).bit_count()


class ImageProject:
    def __init__(self, root: Path, data: dict[str, Any]) -> None:
        self.root = root.expanduser().resolve()
        self.data = data
        self._revision = int(data.get("revision", 0))
        self.recovered_from_backup = False
        self.recovery_detail = ""
        self._recover_pending_cleanup_moves()

    @property
    def manifest_path(self) -> Path:
        return self.root / MANIFEST_NAME

    @property
    def backup_manifest_path(self) -> Path:
        return self.root / BACKUP_MANIFEST_NAME

    @property
    def cleanup_dir(self) -> Path:
        return self.root / CLEANUP_DIR_NAME

    @property
    def cleanup_index_path(self) -> Path:
        return self.cleanup_dir / CLEANUP_INDEX_NAME

    @classmethod
    def create(cls, root: Path, name: str | None = None) -> ImageProject:
        root = root.expanduser().resolve()
        root.mkdir(parents=True, exist_ok=True)
        if (root / MANIFEST_NAME).exists():
            return cls.open(root)
        now = utc_now()
        project = cls(
            root,
            {
                "schema_version": SCHEMA_VERSION,
                "project_id": str(uuid.uuid4()),
                "revision": 0,
                "name": name or root.name,
                "created_at": now,
                "updated_at": now,
                "assets": [],
                "ignored_similar_pairs": [],
            },
        )
        (root / "images").mkdir(parents=True, exist_ok=True)
        project.save()
        return project

    @classmethod
    def open(cls, root: Path) -> ImageProject:
        root = root.expanduser().resolve()
        manifest = root if root.name == MANIFEST_NAME else root / MANIFEST_NAME
        if manifest.name == MANIFEST_NAME and manifest.is_file():
            root = manifest.parent
        try:
            data = cls._read_manifest(manifest)
        except Exception as primary_error:
            backup = root / BACKUP_MANIFEST_NAME
            if not backup.is_file():
                raise primary_error
            data = cls._read_manifest(backup)
            project = cls(root, data)
            project.save(
                recover_invalid_current=True,
                preserve_existing_backup=True,
            )
            project.recovered_from_backup = True
            project.recovery_detail = str(primary_error)
            return project
        return cls(root, data)

    @staticmethod
    def _read_manifest(path: Path) -> dict[str, Any]:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or data.get("schema_version") != SCHEMA_VERSION:
            raise ValueError("不支持的图片库清单格式。")
        if not isinstance(data.get("project_id"), str):
            raise ValueError("图片库缺少 project_id。")
        assets = data.get("assets")
        if not isinstance(assets, list):
            raise ValueError("图片库 assets 必须是列表。")
        ignored_pairs = data.setdefault("ignored_similar_pairs", [])
        if not isinstance(ignored_pairs, list) or not all(
            isinstance(item, str) and len(item.split(":")) == 2
            for item in ignored_pairs
        ):
            raise ValueError("图片库 ignored_similar_pairs 必须是哈希对列表。")
        ids: set[str] = set()
        digests: set[str] = set()
        for asset in assets:
            if not isinstance(asset, dict):
                raise ValueError("图片记录必须是对象。")
            for key in ("id", "sha256", "path", "name", "dhash"):
                if not isinstance(asset.get(key), str) or not asset[key]:
                    raise ValueError(f"图片记录缺少 {key}。")
            if asset["id"] in ids or asset["sha256"] in digests:
                raise ValueError("图片库存在重复 ID 或哈希。")
            if len(asset["sha256"]) != 64:
                raise ValueError("图片 SHA-256 无效。")
            stored = Path(asset["path"])
            if stored.is_absolute():
                raise ValueError("图片库不得保存绝对媒体路径。")
            if not asset.get("format"):
                asset["format"] = {
                    ".jpg": "JPEG",
                    ".jpeg": "JPEG",
                    ".png": "PNG",
                    ".gif": "GIF",
                    ".bmp": "BMP",
                    ".tif": "TIFF",
                    ".tiff": "TIFF",
                    ".webp": "WEBP",
                }.get(stored.suffix.lower(), "")
            try:
                (path.parent / stored).resolve().relative_to(path.parent.resolve())
            except ValueError:
                raise ValueError("图片路径越出图片库目录。") from None
            if not isinstance(asset.get("origins", []), list):
                raise ValueError("图片来源必须是列表。")
            if not isinstance(asset.get("tags", []), list):
                raise ValueError("图片标签必须是列表。")
            ids.add(asset["id"])
            digests.add(asset["sha256"])
        return data

    def save(
        self,
        *,
        recover_invalid_current: bool = False,
        preserve_existing_backup: bool = False,
    ) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        with project_write_lock(self.root, LOCK_NAME):
            current = None
            if self.manifest_path.is_file():
                try:
                    current = self._read_manifest(self.manifest_path)
                except Exception as exc:
                    if not recover_invalid_current:
                        raise RuntimeError(
                            "当前图片库清单已损坏或不可读；拒绝覆盖。"
                            "请重新打开图片库，让程序先从备份恢复。"
                        ) from exc
            if (
                current is not None
                and int(current.get("revision", 0)) != self._revision
            ):
                raise RuntimeError("图片库已在其他窗口修改，请重新打开后再保存。")
            payload = copy.deepcopy(self.data)
            payload["revision"] = self._revision + 1
            payload["updated_at"] = utc_now()
            fd, temp_name = tempfile.mkstemp(
                prefix=".image-project-", suffix=".json", dir=self.root
            )
            os.close(fd)
            temp_path = Path(temp_name)
            backup_temp = self.root / f".{BACKUP_MANIFEST_NAME}.tmp"
            try:
                temp_path.write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                self._read_manifest(temp_path)
                if current is not None and not preserve_existing_backup:
                    shutil.copyfile(self.manifest_path, backup_temp)
                    os.replace(backup_temp, self.backup_manifest_path)
                os.replace(temp_path, self.manifest_path)
                self._revision = int(payload["revision"])
                self.data["revision"] = self._revision
                self.data["updated_at"] = payload["updated_at"]
            finally:
                temp_path.unlink(missing_ok=True)
                backup_temp.unlink(missing_ok=True)

    def reload(self) -> None:
        refreshed = type(self).open(self.root)
        self.data = refreshed.data
        self._revision = refreshed._revision

    def assets(self) -> list[dict[str, Any]]:
        return self.data["assets"]

    def asset(self, asset_id: str) -> dict[str, Any]:
        for item in self.assets():
            if item["id"] == asset_id:
                return item
        raise KeyError(asset_id)

    def asset_path(self, asset: dict[str, Any]) -> Path:
        return (self.root / asset["path"]).resolve()

    def find_by_hash(self, digest: str) -> dict[str, Any] | None:
        return next((item for item in self.assets() if item["sha256"] == digest), None)

    def find_by_pixels(self, metadata: dict[str, Any]) -> dict[str, Any] | None:
        digest = metadata["pixel_sha256"]
        for asset in self.assets():
            if (
                asset["width"] != metadata["width"]
                or asset["height"] != metadata["height"]
                or asset["dhash"] != metadata["dhash"]
            ):
                continue
            if not asset.get("pixel_sha256"):
                try:
                    asset["pixel_sha256"] = _image_metadata(
                        self.asset_path(asset).read_bytes()
                    )["pixel_sha256"]
                except (OSError, ValueError) as exc:
                    LOGGER.warning(
                        "Unable to calculate pixel hash for %s: %s",
                        asset.get("path", asset.get("id", "unknown")),
                        exc,
                    )
                    continue
            if asset["pixel_sha256"] == digest:
                return asset
        return None

    @staticmethod
    def _similar_pair_key(left: dict[str, Any], right: dict[str, Any]) -> str:
        return ":".join(sorted((left["sha256"], right["sha256"])))

    def ignore_similar_pair(self, left_id: str, right_id: str) -> None:
        left = self.asset(left_id)
        right = self.asset(right_id)
        if left_id == right_id:
            raise ValueError("不能忽略图片自身。")
        original = copy.deepcopy(self.data)
        pairs = self.data.setdefault("ignored_similar_pairs", [])
        key = self._similar_pair_key(left, right)
        if key in pairs:
            return
        pairs.append(key)
        try:
            self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            raise

    def reset_ignored_similar_pairs(self) -> int:
        original = copy.deepcopy(self.data)
        pairs = self.data.setdefault("ignored_similar_pairs", [])
        count = len(pairs)
        if not count:
            return 0
        self.data["ignored_similar_pairs"] = []
        try:
            self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            raise
        return count

    def import_paths(
        self,
        paths: Iterable[Path],
        *,
        category: str = "",
        min_width: int = 0,
        min_height: int = 0,
        cancel_callback: Callable[[], bool] | None = None,
    ) -> dict[str, Any]:
        if min_width < 0 or min_height < 0:
            raise ValueError("图片最小宽度和高度不能小于 0。")
        added = 0
        reused = 0
        failed: list[dict[str, str]] = []
        skipped: list[dict[str, Any]] = []
        imported: list[str] = []
        cancelled = False
        original = copy.deepcopy(self.data)
        created_files: list[Path] = []
        try:
            for source in paths:
                if cancel_callback is not None and cancel_callback():
                    cancelled = True
                    break
                source = source.expanduser().resolve()
                try:
                    source_type, records = self.source_images(source)
                    for member_name, data in records:
                        if cancel_callback is not None and cancel_callback():
                            cancelled = True
                            break
                        try:
                            metadata = _image_metadata(data)
                            if (
                                metadata["width"] < min_width
                                or metadata["height"] < min_height
                            ):
                                skipped.append(
                                    {
                                        "source": str(source),
                                        "member": member_name,
                                        "width": metadata["width"],
                                        "height": metadata["height"],
                                        "reason": (
                                            f"低于最小尺寸 {min_width}×{min_height}"
                                        ),
                                    }
                                )
                                continue
                            asset, was_added, created_path = self._archive_bytes(
                                data,
                                member_name,
                                {
                                    "source_type": source_type,
                                    "source_path": str(source),
                                    "member_path": member_name,
                                    "imported_at": utc_now(),
                                },
                                metadata=metadata,
                            )
                            if created_path is not None:
                                created_files.append(created_path)
                        except ValueError as exc:
                            failed.append(
                                {
                                    "source": str(source),
                                    "member": member_name,
                                    "error": str(exc),
                                }
                            )
                            continue
                        added += int(was_added)
                        reused += int(not was_added)
                        if category and (was_added or not asset.get("category")):
                            asset["category"] = category.strip()
                            asset["updated_at"] = utc_now()
                        imported.append(asset["id"])
                    if cancelled:
                        break
                except (OSError, ValueError, BadZipFile, PdfReadError) as exc:
                    failed.append(
                        {"source": str(source), "member": "", "error": str(exc)}
                    )
            if added or reused:
                self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            for path in created_files:
                path.unlink(missing_ok=True)
            raise
        return {
            "added": added,
            "reused": reused,
            "failed": failed,
            "skipped": skipped,
            "asset_ids": list(dict.fromkeys(imported)),
            "cancelled": cancelled,
        }

    @classmethod
    def source_images(cls, source: Path) -> tuple[str, list[tuple[str, bytes]]]:
        source = source.expanduser().resolve()
        suffix = source.suffix.lower()
        if suffix in SUPPORTED_IMAGE_SUFFIXES:
            return "image", [(source.name, source.read_bytes())]
        if suffix in OFFICE_MEDIA_PREFIXES:
            return suffix.lstrip("."), cls._extract_office_images(source)
        if suffix == ".pdf":
            return "pdf", cls._extract_pdf_images(source)
        raise ValueError("仅支持图片、PPTX、DOCX 和 PDF。")

    @staticmethod
    def _extract_office_images(path: Path) -> list[tuple[str, bytes]]:
        prefix = OFFICE_MEDIA_PREFIXES[path.suffix.lower()]
        with ZipFile(path) as archive:
            return [
                (member, archive.read(member))
                for member in _referenced_office_images(archive, prefix)
                if Path(member).suffix.lower() in SUPPORTED_IMAGE_SUFFIXES
            ]

    @staticmethod
    def _extract_pdf_images(path: Path) -> list[tuple[str, bytes]]:
        reader = PdfReader(str(path))
        records: list[tuple[str, bytes]] = []
        for page_number, page in enumerate(reader.pages, start=1):
            for index, image_file in enumerate(page.images, start=1):
                data = bytes(image_file.data)
                name = str(
                    getattr(image_file, "name", "")
                    or f"page-{page_number:04d}-image-{index:03d}.bin"
                )
                records.append((f"page-{page_number:04d}/{name}", data))
        return records

    def _archive_bytes(
        self,
        data: bytes,
        source_name: str,
        origin: dict[str, Any],
        *,
        metadata: dict[str, Any] | None = None,
    ) -> tuple[dict[str, Any], bool, Path | None]:
        metadata = metadata or _image_metadata(data)
        digest = sha256_bytes(data)
        existing = self.find_by_hash(digest) or self.find_by_pixels(metadata)
        if existing is not None:
            origins = existing.setdefault("origins", [])
            identity = (
                origin.get("source_path", ""),
                origin.get("member_path", ""),
            )
            if not any(
                (item.get("source_path", ""), item.get("member_path", "")) == identity
                for item in origins
            ):
                origins.append(origin)
                existing["updated_at"] = utc_now()
            return existing, False, None
        suffix = _safe_suffix(source_name, metadata["format"])
        relative = Path("images") / digest[:2] / f"{digest}{suffix}"
        target = self.root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        created_path = None
        if not target.exists():
            fd, temp_name = tempfile.mkstemp(prefix=".image-", dir=target.parent)
            os.close(fd)
            temp_path = Path(temp_name)
            try:
                temp_path.write_bytes(data)
                os.replace(temp_path, target)
                created_path = target
            finally:
                temp_path.unlink(missing_ok=True)
        now = utc_now()
        asset = {
            "id": str(uuid.uuid4()),
            "sha256": digest,
            "path": relative.as_posix(),
            "name": Path(source_name).stem or digest[:12],
            "category": "",
            "tags": [],
            "summary": "",
            "width": metadata["width"],
            "height": metadata["height"],
            "format": metadata["format"],
            "mode": metadata["mode"],
            "size_bytes": len(data),
            "dhash": metadata["dhash"],
            "pixel_sha256": metadata["pixel_sha256"],
            "origins": [origin],
            "created_at": now,
            "updated_at": now,
        }
        self.assets().append(asset)
        return asset, True, created_path

    def _similar_matches(
        self, asset_id: str
    ) -> Iterable[tuple[dict[str, Any], int, float]]:
        source = self.asset(asset_id)
        source_ratio = source["width"] / max(source["height"], 1)
        ignored = set(self.data.get("ignored_similar_pairs") or [])
        for candidate in self.assets():
            if candidate["id"] == asset_id:
                continue
            if self._similar_pair_key(source, candidate) in ignored:
                continue
            ratio = candidate["width"] / max(candidate["height"], 1)
            ratio_diff = abs(source_ratio - ratio) / max(source_ratio, ratio, 0.001)
            if ratio_diff > 0.08:
                continue
            distance = _hamming_hex(source["dhash"], candidate["dhash"])
            if distance > 10:
                continue
            yield candidate, distance, ratio_diff

    def similar_candidates(
        self, asset_id: str, *, limit: int = 12
    ) -> list[dict[str, Any]]:
        candidates: list[dict[str, Any]] = []
        for candidate, distance, ratio_diff in self._similar_matches(asset_id):
            candidates.append(
                {
                    "asset_id": candidate["id"],
                    "name": candidate["name"],
                    "distance": distance,
                    "score": round(
                        max(0.0, 100.0 - distance * 8 - ratio_diff * 100), 1
                    ),
                    "width": candidate["width"],
                    "height": candidate["height"],
                }
            )
        candidates.sort(
            key=lambda item: (item["distance"], -item["score"], item["name"])
        )
        return candidates[: max(1, limit)]

    def asset_health(
        self,
        asset_id: str,
        *,
        min_width: int = 0,
        min_height: int = 0,
    ) -> dict[str, bool]:
        asset = self.asset(asset_id)
        return {
            "duplicate_origins": len(asset.get("origins") or []) > 1,
            "similar": next(iter(self._similar_matches(asset_id)), None) is not None,
            "undersized": bool(
                (min_width and asset["width"] < min_width)
                or (min_height and asset["height"] < min_height)
            ),
            "no_origin": not bool(asset.get("origins")),
        }

    def health_counts(
        self, *, min_width: int = 0, min_height: int = 0
    ) -> dict[str, int]:
        counts = {
            "all": len(self.assets()),
            "duplicate_origins": 0,
            "similar": 0,
            "undersized": 0,
            "no_origin": 0,
        }
        for asset in self.assets():
            for key, matched in self.asset_health(
                asset["id"], min_width=min_width, min_height=min_height
            ).items():
                counts[key] += int(matched)
        return counts

    def health_report(self, *, verify_hashes: bool = False) -> dict[str, Any]:
        missing: list[str] = []
        modified: list[str] = []
        for asset in self.assets():
            path = self.asset_path(asset)
            if not path.is_file():
                missing.append(asset["path"])
                continue
            if path.stat().st_size != asset["size_bytes"] or (
                verify_hashes and sha256_file(path) != asset["sha256"]
            ):
                modified.append(asset["path"])
        orphans = [
            path.relative_to(self.root).as_posix() for path in self.orphan_paths()
        ]
        cleanup_issues = self.cleanup_pending_issues()
        return {
            "ok": not (missing or modified or orphans or cleanup_issues),
            "asset_count": len(self.assets()),
            "missing_files": missing,
            "modified_files": modified,
            "orphan_files": orphans,
            "pending_cleanup_count": len(self.pending_cleanup()),
            "pending_cleanup_issues": cleanup_issues,
            "recovered_from_backup": self.recovered_from_backup,
            "hashes_verified": verify_hashes,
        }

    def update_metadata(
        self,
        asset_id: str,
        *,
        name: str | None = None,
        category: str | None = None,
        tags: Iterable[str] | None = None,
        summary: str | None = None,
    ) -> None:
        original = copy.deepcopy(self.data)
        asset = self.asset(asset_id)
        try:
            if name is not None:
                cleaned = name.strip()
                if not cleaned:
                    raise ValueError("图片名称不能为空。")
                asset["name"] = cleaned
            if category is not None:
                asset["category"] = category.strip()
            if tags is not None:
                asset["tags"] = list(
                    dict.fromkeys(tag.strip() for tag in tags if tag.strip())
                )
            if summary is not None:
                asset["summary"] = summary.strip()
            asset["updated_at"] = utc_now()
            self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            raise

    def merge_assets(
        self,
        primary_id: str,
        asset_ids: Iterable[str],
        *,
        confirmed_same_content: bool = False,
    ) -> dict[str, Any]:
        if not confirmed_same_content:
            raise ValueError("合并图片必须经过人工确认。")
        ids = list(dict.fromkeys(str(item) for item in asset_ids))
        if primary_id not in ids or len(ids) < 2:
            raise ValueError("图片合并候选或主图片无效。")
        original = copy.deepcopy(self.data)
        primary = self.asset(primary_id)
        losers = [self.asset(asset_id) for asset_id in ids if asset_id != primary_id]
        quarantined: list[dict[str, Any]] = []
        try:
            for loser in losers:
                entry = self._quarantine_file(
                    self.asset_path(loser),
                    reason="人工确认：合并重复图片",
                    asset=loser,
                )
                if entry is not None:
                    quarantined.append(entry)
            origins = primary.setdefault("origins", [])
            origin_keys = {
                (item.get("source_path", ""), item.get("member_path", ""))
                for item in origins
            }
            tags = list(primary.get("tags") or [])
            for loser in losers:
                for origin in loser.get("origins") or []:
                    key = (
                        origin.get("source_path", ""),
                        origin.get("member_path", ""),
                    )
                    if key not in origin_keys:
                        origins.append(origin)
                        origin_keys.add(key)
                tags.extend(loser.get("tags") or [])
                if not primary.get("category") and loser.get("category"):
                    primary["category"] = loser["category"]
                if not primary.get("summary") and loser.get("summary"):
                    primary["summary"] = loser["summary"]
            primary["tags"] = list(dict.fromkeys(tag for tag in tags if tag))
            primary["updated_at"] = utc_now()
            loser_ids = {asset["id"] for asset in losers}
            self.data["assets"] = [
                asset for asset in self.assets() if asset["id"] not in loser_ids
            ]
            self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            self._rollback_quarantine(quarantined)
            raise
        return {
            "primary_id": primary_id,
            "removed_ids": [asset["id"] for asset in losers],
        }

    def remove_asset(self, asset_id: str) -> Path:
        original = copy.deepcopy(self.data)
        asset = self.asset(asset_id)
        path = self.asset_path(asset)
        quarantined: list[dict[str, Any]] = []
        entry = self._quarantine_file(
            path,
            reason="人工确认：从图片库移除",
            asset=asset,
        )
        if entry is not None:
            quarantined.append(entry)
        self.data["assets"] = [item for item in self.assets() if item["id"] != asset_id]
        try:
            self.save()
        except Exception:
            self.data = original
            self._revision = int(original.get("revision", 0))
            self._rollback_quarantine(quarantined)
            raise
        return path

    def orphan_paths(self) -> list[Path]:
        referenced = {self.asset_path(item) for item in self.assets()}
        image_root = self.root / "images"
        if not image_root.is_dir():
            return []
        return sorted(
            candidate
            for candidate in image_root.rglob("*")
            if candidate.is_file() and candidate.resolve() not in referenced
        )

    def cleanup_orphans(self) -> list[Path]:
        image_root = self.root / "images"
        removed = self.orphan_paths()
        quarantined: list[dict[str, Any]] = []
        try:
            for candidate in removed:
                entry = self._quarantine_file(
                    candidate,
                    reason="人工确认：清理未引用文件",
                )
                if entry is not None:
                    quarantined.append(entry)
        except Exception:
            self._rollback_quarantine(quarantined)
            raise
        if image_root.is_dir():
            for directory in sorted(image_root.rglob("*"), reverse=True):
                if directory.is_dir() and not any(directory.iterdir()):
                    directory.rmdir()
        return removed

    def _read_cleanup_index(self) -> list[dict[str, Any]]:
        if not self.cleanup_index_path.is_file():
            return []
        try:
            data = json.loads(self.cleanup_index_path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise RuntimeError(
                f"待清理索引损坏，拒绝继续操作：{self.cleanup_index_path}"
            ) from exc
        if not isinstance(data, list) or not all(
            isinstance(item, dict) for item in data
        ):
            raise RuntimeError(
                f"待清理索引格式无效，拒绝继续操作：{self.cleanup_index_path}"
            )
        return data

    def _cleanup_entry_path(self, entry: dict[str, Any], key: str) -> Path:
        value = entry.get(key)
        if not isinstance(value, str) or not value:
            raise RuntimeError(f"待清理索引缺少 {key}")
        stored = Path(value).expanduser()
        path = (stored if stored.is_absolute() else self.root / stored).resolve()
        allowed_root = (
            self.cleanup_dir.resolve()
            if key == "quarantined_path"
            else (self.root / "images").resolve()
        )
        try:
            path.relative_to(allowed_root)
        except ValueError:
            raise RuntimeError(f"待清理索引包含越界路径，拒绝操作：{path}") from None
        return path

    def _write_cleanup_index(self, entries: list[dict[str, Any]]) -> None:
        self.cleanup_dir.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=".cleanup-index-", suffix=".json", dir=self.cleanup_dir
        )
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            temp_path.write_text(
                json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            os.replace(temp_path, self.cleanup_index_path)
        finally:
            temp_path.unlink(missing_ok=True)

    def _quarantine_file(
        self,
        source: Path,
        *,
        reason: str,
        asset: dict[str, Any] | None = None,
    ) -> dict[str, Any] | None:
        source = source.resolve()
        if not source.is_file():
            return None
        try:
            original_path = source.relative_to((self.root / "images").resolve())
        except ValueError:
            raise ValueError(f"拒绝隔离图片库外的文件：{source}") from None
        original_path = Path("images") / original_path
        token = uuid.uuid4().hex[:12]
        target = self.cleanup_dir / f"{token}_{source.name}"
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            self.cleanup_dir.mkdir(parents=True, exist_ok=True)
            entry = {
                "token": token,
                "asset": copy.deepcopy(asset),
                "original_path": original_path.as_posix(),
                "quarantined_path": target.relative_to(self.root).as_posix(),
                "sha256": str((asset or {}).get("sha256") or sha256_file(source)),
                "quarantined_at": utc_now(),
                "reason": reason,
                "state": "moving",
            }
            self._write_cleanup_index([*entries, entry])
            try:
                shutil.move(str(source), target)
                self._write_cleanup_index([*entries, {**entry, "state": "quarantined"}])
            except Exception:
                restored = source.is_file() and not target.exists()
                if target.is_file() and not source.exists():
                    try:
                        source.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(target), source)
                        restored = True
                    except OSError:
                        LOGGER.exception("Unable to roll back image quarantine move")
                if restored:
                    self._write_cleanup_index(entries)
                raise
        return entry

    def _recover_pending_cleanup_moves(self) -> None:
        if not self.cleanup_index_path.is_file():
            return
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            recovered: list[dict[str, Any]] = []
            changed = False
            for entry in entries:
                if entry.get("state") != "moving":
                    recovered.append(entry)
                    continue
                source = self._cleanup_entry_path(entry, "original_path")
                target = self._cleanup_entry_path(entry, "quarantined_path")
                if (
                    target.is_file()
                    and not source.exists()
                    and sha256_file(target) == entry.get("sha256")
                ):
                    recovered.append({**entry, "state": "quarantined"})
                    changed = True
                elif source.is_file() and not target.exists():
                    changed = True
                else:
                    recovered.append(entry)
            if changed:
                self._write_cleanup_index(recovered)

    def _rollback_quarantine(
        self,
        quarantined: list[dict[str, Any]],
    ) -> None:
        restored_tokens: set[str] = set()
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            for entry in reversed(quarantined):
                source = self._cleanup_entry_path(entry, "quarantined_path")
                target = self._cleanup_entry_path(entry, "original_path")
                if source.is_file() and not target.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(source), target)
                    restored_tokens.add(entry["token"])
                elif (
                    not source.exists()
                    and target.is_file()
                    and sha256_file(target) == entry.get("sha256")
                ):
                    restored_tokens.add(entry["token"])
            self._write_cleanup_index(
                [
                    entry
                    for entry in self._read_cleanup_index()
                    if entry.get("token") not in restored_tokens
                ]
            )

    def pending_cleanup(self) -> list[dict[str, Any]]:
        pending = []
        for entry in self._read_cleanup_index():
            path = self._cleanup_entry_path(entry, "quarantined_path")
            pending.append(
                {
                    **entry,
                    "quarantined_path": str(path),
                    "exists": path.is_file(),
                    "size_bytes": path.stat().st_size if path.is_file() else 0,
                }
            )
        return pending

    def cleanup_pending_issues(self) -> list[str]:
        referenced = {asset["id"] for asset in self.assets()}
        issues: list[str] = []
        for entry in self._read_cleanup_index():
            asset = entry.get("asset")
            if isinstance(asset, dict) and asset.get("id") in referenced:
                issues.append(f"图片记录仍被清单引用：{asset['id']}")
            if entry.get("state") == "moving":
                issues.append(f"图片隔离操作未完成：{entry.get('token', '?')}")
            path = self._cleanup_entry_path(entry, "quarantined_path")
            if not path.is_file():
                issues.append(f"待清理文件丢失：{path}")
        return issues

    def restore_cleanup_entry(self, token: str) -> Path:
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            entries = self._read_cleanup_index()
            entry = next(
                (item for item in entries if item.get("token") == token),
                None,
            )
            if entry is None:
                raise KeyError(f"Unknown cleanup entry: {token}")
            quarantined = self._cleanup_entry_path(entry, "quarantined_path")
            target = self._cleanup_entry_path(entry, "original_path")
            asset = copy.deepcopy(entry.get("asset"))
            existing = (
                next(
                    (
                        item
                        for item in self.assets()
                        if isinstance(asset, dict)
                        and (
                            item["id"] == asset["id"]
                            or item["sha256"] == asset["sha256"]
                        )
                    ),
                    None,
                )
                if isinstance(asset, dict)
                else None
            )
            if not quarantined.is_file():
                if existing is not None:
                    existing_path = self.asset_path(existing)
                    if existing_path.is_file() and sha256_file(
                        existing_path
                    ) == entry.get("sha256"):
                        self._write_cleanup_index(
                            [item for item in entries if item.get("token") != token]
                        )
                        return existing_path
                raise FileNotFoundError(quarantined)
            if sha256_file(quarantined) != entry.get("sha256"):
                raise ValueError("待清理文件内容已变化，拒绝还原")
            if target.exists():
                raise FileExistsError(f"原位置已有同名文件，拒绝覆盖：{target}")
            if existing is not None:
                if (
                    isinstance(asset, dict)
                    and existing.get("id") == asset.get("id")
                    and self.asset_path(existing) == target
                ):
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(quarantined), target)
                    self._write_cleanup_index(
                        [item for item in entries if item.get("token") != token]
                    )
                    return target
                raise RuntimeError("该图片记录或内容已在图片库中，拒绝重复还原")
            original_data = copy.deepcopy(self.data)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(quarantined), target)
            try:
                if isinstance(asset, dict):
                    asset["path"] = target.relative_to(self.root).as_posix()
                    asset["size_bytes"] = target.stat().st_size
                    asset["updated_at"] = utc_now()
                    self.assets().append(asset)
                    self.save()
            except Exception:
                self.data = original_data
                self._revision = int(original_data.get("revision", 0))
                if target.is_file() and not quarantined.exists():
                    shutil.move(str(target), quarantined)
                raise
            self._write_cleanup_index(
                [item for item in entries if item.get("token") != token]
            )
            return target

    def empty_cleanup(self) -> int:
        with project_write_lock(self.root, CLEANUP_LOCK_NAME):
            issues = self.cleanup_pending_issues()
            if issues:
                raise RuntimeError("待清理目录不满足清空条件：\n" + "\n".join(issues))
            entries = self._read_cleanup_index()
            removed = 0
            for entry in entries:
                path = self._cleanup_entry_path(entry, "quarantined_path")
                if path.is_file():
                    path.unlink()
                    removed += 1
            self._write_cleanup_index([])
            try:
                self.cleanup_index_path.unlink(missing_ok=True)
                self.cleanup_dir.rmdir()
            except OSError:
                pass
            return removed


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect a Doc Media image library")
    subparsers = parser.add_subparsers(dest="action", required=True)
    doctor = subparsers.add_parser("doctor")
    doctor.add_argument("project", type=Path)
    doctor.add_argument("--verify-hashes", action="store_true")
    doctor.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    project = ImageProject.open(args.project)
    report = project.health_report(verify_hashes=args.verify_hashes)
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        destination = args.report.expanduser().resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(rendered + "\n", encoding="utf-8")
        print(destination)
    else:
        print(rendered)
    return 0 if report["ok"] else 1
