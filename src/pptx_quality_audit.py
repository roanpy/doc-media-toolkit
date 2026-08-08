import concurrent.futures
import re
import shutil
import subprocess
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from PySide6.QtCore import QObject, Signal
from PIL import Image, ImageChops, UnidentifiedImageError
from pptx_output_watermark.process_utils import (
    finish_process,
    kill_process,
    start_process,
)
from pptx_video_compactor import (
    hidden_subprocess_kwargs,
    load_json_file,
    resolve_binary,
)


@dataclass(slots=True)
class AuditResult:
    media_path: str
    ssim: float | None
    is_video: bool
    status: str
    error: str | None = None


class QualityAuditWorker(QObject):
    progress = Signal(int, str)
    log = Signal(str)
    finished = Signal(list)  # list[AuditResult]

    def __init__(self, input_pptx: Path, output_pptx: Path, report_path: Path):
        super().__init__()
        self.input_pptx = input_pptx
        self.output_pptx = output_pptx
        self.report_path = report_path
        self.cancel_requested = False
        self._active_procs: set[subprocess.Popen] = set()
        self._process_lock = threading.Lock()

    @staticmethod
    def _is_standalone_report(report_data: dict) -> bool:
        return str(report_data.get("input_kind", "")).startswith("standalone_")

    @staticmethod
    def _copy_standalone_assets(
        report_data: dict,
        assets_to_check: list[dict],
        original_dir: Path,
        compressed_dir: Path,
    ) -> bool:
        input_path = Path(report_data.get("input_pptx") or "")
        output_path = Path(report_data.get("output_pptx") or "")
        if not input_path.exists() or not output_path.exists():
            return False
        for asset in assets_to_check:
            orig_target = QualityAuditWorker._asset_target(
                original_dir, asset["media_path"]
            )
            comp_target = QualityAuditWorker._asset_target(
                compressed_dir, asset["output_media_path"]
            )
            orig_target.parent.mkdir(parents=True, exist_ok=True)
            comp_target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(input_path, orig_target)
            shutil.copy2(output_path, comp_target)
        return True

    @staticmethod
    def _asset_target(root: Path, relative_path: str) -> Path:
        resolved_root = root.resolve()
        target = (resolved_root / relative_path).resolve()
        if not target.is_relative_to(resolved_root):
            raise ValueError(f"媒体路径超出评估目录: {relative_path}")
        return target

    @classmethod
    def _extract_asset(cls, archive: ZipFile, member: str, root: Path) -> None:
        target = cls._asset_target(root, member)
        target.parent.mkdir(parents=True, exist_ok=True)
        with archive.open(member, "r") as source, target.open("wb") as output:
            shutil.copyfileobj(source, output, length=1024 * 1024)

    def cancel(self) -> None:
        with self._process_lock:
            self.cancel_requested = True
            processes = list(self._active_procs)
        for proc in processes:
            kill_process(proc)

    @staticmethod
    def _stderr_tail(stderr: str | None, max_lines: int = 4) -> str:
        lines = [line.strip() for line in (stderr or "").splitlines() if line.strip()]
        return " | ".join(lines[-max_lines:])

    @staticmethod
    def _images_are_pixel_identical(first: Path, second: Path) -> bool:
        try:
            with Image.open(first) as first_image, Image.open(second) as second_image:
                if first_image.size != second_image.size:
                    return False
                if first_image.width * first_image.height > 25_000_000:
                    return False
                first_rgba = first_image.convert("RGBA")
                second_rgba = second_image.convert("RGBA")
                return ImageChops.difference(first_rgba, second_rgba).getbbox() is None
        except (UnidentifiedImageError, OSError, ValueError):
            return False

    @classmethod
    def _ffmpeg_audit_error(
        cls, asset: dict, stderr: str | None, returncode: int | None
    ) -> str:
        kind = "视频" if asset["is_video"] else "图片"
        suffix = Path(asset["media_path"]).suffix.lower()
        message = f"FFmpeg 无法计算{kind}质量评分"
        if asset["is_video"] and suffix and suffix not in {".mp4", ".m4v", ".mov"}:
            message += f"：源文件为 {suffix.upper().lstrip('.')}，可能是当前 FFmpeg 无法解码该编码"
        if returncode:
            message += f" (returncode={returncode})"
        detail = cls._stderr_tail(stderr)
        if detail:
            message += f"。详情：{detail}"
        return message

    def _evaluate_single(
        self, asset: dict, original_dir: Path, compressed_dir: Path
    ) -> AuditResult:
        if self.cancel_requested:
            return AuditResult(
                asset["media_path"], None, asset["is_video"], "error", "已取消"
            )

        orig_file = self._asset_target(original_dir, asset["media_path"])
        comp_file = self._asset_target(compressed_dir, asset["output_media_path"])

        if not orig_file.exists() or not comp_file.exists():
            return AuditResult(
                asset["media_path"], None, asset["is_video"], "error", "文件解压失败"
            )

        try:
            ffmpeg_binary = resolve_binary("ffmpeg")
            if ffmpeg_binary is None:
                return AuditResult(
                    asset["media_path"],
                    None,
                    asset["is_video"],
                    "error",
                    "未找到 ffmpeg",
                )
            if asset["is_video"]:
                filter_cmd = (
                    "[0:v]setpts=PTS-STARTPTS,fps=1[v0];"
                    "[1:v]setpts=PTS-STARTPTS,fps=1[v1];"
                    "[v0][v1]scale2ref[v0s][v1s];[v0s][v1s]ssim"
                )
            else:
                # Scale the compressed image back to the original canvas before SSIM.
                # This keeps low-size presets with image downsampling evaluable.
                filter_cmd = "[0:v][1:v]scale2ref[v0s][v1s];[v0s][v1s]ssim"

            cmd = [
                ffmpeg_binary,
                "-i",
                str(comp_file),
                "-i",
                str(orig_file),
                "-lavfi",
                filter_cmd,
                "-f",
                "null",
                "-",
            ]

            with self._process_lock:
                if self.cancel_requested:
                    return AuditResult(
                        asset["media_path"], None, asset["is_video"], "error", "已取消"
                    )
                proc = start_process(
                    cmd,
                    stderr=subprocess.PIPE,
                    stdout=subprocess.DEVNULL,
                    text=True,
                    **hidden_subprocess_kwargs(),
                )
                self._active_procs.add(proc)

            try:
                _, stderr = proc.communicate(timeout=300)
            except subprocess.TimeoutExpired:
                kill_process(proc)
                proc.communicate()
                return AuditResult(
                    asset["media_path"], None, asset["is_video"], "error", "计算超时"
                )
            finally:
                with self._process_lock:
                    self._active_procs.discard(proc)
                finish_process(proc)

            if self.cancel_requested:
                return AuditResult(
                    asset["media_path"], None, asset["is_video"], "error", "已取消"
                )

            if proc.returncode:
                return AuditResult(
                    asset["media_path"],
                    None,
                    asset["is_video"],
                    "error",
                    self._ffmpeg_audit_error(asset, stderr, proc.returncode),
                )

            match = re.search(r"All:([0-9.]+)", stderr)
            if match:
                score = float(match.group(1))
                if (
                    not asset["is_video"]
                    and score <= 0.0
                    and self._images_are_pixel_identical(orig_file, comp_file)
                ):
                    score = 1.0
                return AuditResult(
                    asset["media_path"], score, asset["is_video"], "success"
                )
            if not asset["is_video"] and self._images_are_pixel_identical(
                orig_file, comp_file
            ):
                return AuditResult(asset["media_path"], 1.0, False, "success")
            return AuditResult(
                asset["media_path"],
                None,
                asset["is_video"],
                "error",
                self._ffmpeg_audit_error(asset, stderr, proc.returncode),
            )
        except Exception as e:
            return AuditResult(
                asset["media_path"], None, asset["is_video"], "error", str(e)
            )

    def run(self) -> None:
        results: list[AuditResult] = []
        temp_dir = Path(tempfile.mkdtemp(prefix="pptx_audit_"))
        completed = False

        try:
            if not self.report_path.exists():
                self.log.emit("找不到对应的报告文件 (report.json)，无法进行跑分对比。")
                return

            try:
                report_data = load_json_file(
                    self.report_path, source="Quality audit report"
                )
                output_pptx_str = report_data.get("output_pptx")
                if not output_pptx_str or not Path(output_pptx_str).exists():
                    self.log.emit("报告中指定的输出 PPTX 不存在，无法对比。")
                    return
                # Only fallback if self.output_pptx is a dummy
                if str(self.output_pptx) == "dummy" or not self.output_pptx.exists():
                    self.output_pptx = Path(output_pptx_str)
            except Exception as e:
                self.log.emit(f"读取报告文件失败: {e}")
                return

            assets_to_check = []
            for video in report_data.get("videos", []):
                if video.get("status") not in {"encoded", "encoded_gpu", "copied"}:
                    continue
                assets_to_check.append(
                    {
                        "media_path": video["media_path"],
                        "output_media_path": video.get("output_media_path")
                        or video["media_path"],
                        "is_video": True,
                        "zip_size": video.get(
                            "zip_size_bytes", video.get("zip_size", 1000)
                        ),
                    }
                )

            for image in report_data.get("images", []):
                if image.get("status") not in {"encoded", "optimize", "copied"}:
                    continue
                assets_to_check.append(
                    {
                        "media_path": image["media_path"],
                        "output_media_path": image["media_path"],
                        "is_video": False,
                        "zip_size": image.get(
                            "zip_size_bytes", image.get("zip_size", 1000)
                        ),
                    }
                )

            if not assets_to_check:
                self.log.emit("没有发现被重压的媒体文件，跳过评估。")
                return

            original_dir = temp_dir / "original"
            compressed_dir = temp_dir / "compressed"
            original_dir.mkdir()
            compressed_dir.mkdir()

            try:
                for asset in assets_to_check:
                    self._asset_target(original_dir, asset["media_path"])
                    self._asset_target(compressed_dir, asset["output_media_path"])
            except ValueError as exc:
                self.log.emit(f"报告包含不安全的媒体路径，已停止评估: {exc}")
                return

            self.progress.emit(10, "解压素材进行对比准备...")

            if self._is_standalone_report(report_data):
                if not self._copy_standalone_assets(
                    report_data, assets_to_check, original_dir, compressed_dir
                ):
                    self.log.emit("报告中指定的独立媒体输入或输出不存在，无法对比。")
                    return
            else:
                with ZipFile(self.input_pptx, "r") as zin:
                    for asset in assets_to_check:
                        try:
                            self._extract_asset(zin, asset["media_path"], original_dir)
                        except KeyError:
                            pass

                with ZipFile(self.output_pptx, "r") as zout:
                    for asset in assets_to_check:
                        try:
                            self._extract_asset(
                                zout, asset["output_media_path"], compressed_dir
                            )
                        except KeyError:
                            pass

            total_size = sum(asset["zip_size"] for asset in assets_to_check)
            processed_size = 0

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                future_to_asset = {
                    executor.submit(
                        self._evaluate_single, asset, original_dir, compressed_dir
                    ): asset
                    for asset in assets_to_check
                }

                for future in concurrent.futures.as_completed(future_to_asset):
                    if self.cancel_requested:
                        break
                    asset = future_to_asset[future]
                    try:
                        result = future.result()
                        results.append(result)
                    except Exception as e:
                        results.append(
                            AuditResult(
                                asset["media_path"],
                                None,
                                asset["is_video"],
                                "error",
                                str(e),
                            )
                        )

                    processed_size += asset["zip_size"]
                    percent = 10 + int(80 * (processed_size / max(1, total_size)))
                    self.progress.emit(percent, f"正在并发评估... {percent}%")
            completed = True

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            if completed and not self.cancel_requested:
                self.progress.emit(100, "画质评估完成")
            self.finished.emit(results)
