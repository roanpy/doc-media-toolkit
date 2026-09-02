from __future__ import annotations

import base64
import io
import json
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.parse import urlparse

from PIL import Image, ImageOps
from pptx_tools import __version__


MAX_CONTEXT_CHARS = 12_000
MAX_RESPONSE_BYTES = 5 * 1024 * 1024
TRANSIENT_RETRIES = 2


@dataclass(frozen=True)
class AIConfig:
    base_url: str
    model: str
    api_key: str = ""
    vision_enabled: bool = True
    timeout_seconds: int = 120
    context: str = ""


class AIClientError(RuntimeError):
    pass


def privacy_scope(config: AIConfig) -> str:
    parsed = urlparse(config.base_url.strip())
    host = (parsed.hostname or "invalid-host").lower()
    try:
        port = parsed.port
    except ValueError:
        port = None
    authority = f"{host}:{port}" if port is not None else host
    mode = "vision" if config.vision_enabled else "text"
    return f"{authority}|{normalize_model_name(config.model)}|{mode}"


def normalize_model_name(value: str) -> str:
    return value.strip().translate(
        str.maketrans({"–": "-", "—": "-", "−": "-", "‑": "-"})
    )


def _redact_secret(value: object, secret: str) -> str:
    text = str(value)
    return text.replace(secret, "***") if secret else text


def _retry_delay(headers: object, retry_index: int) -> float:
    fallback = 0.5 * (2**retry_index)
    value = headers.get("Retry-After") if hasattr(headers, "get") else None
    if value:
        try:
            delay = float(value)
        except (TypeError, ValueError):
            try:
                retry_at = parsedate_to_datetime(str(value))
                if retry_at.tzinfo is None:
                    retry_at = retry_at.replace(tzinfo=timezone.utc)
                delay = (retry_at - datetime.now(timezone.utc)).total_seconds()
            except (TypeError, ValueError, OverflowError):
                delay = fallback
        return max(0.5, min(10.0, delay))
    return fallback


def _wait_before_retry(
    delay: float, cancel_callback: Callable[[], bool] | None
) -> None:
    remaining = delay
    while remaining > 0:
        if cancel_callback and cancel_callback():
            raise AIClientError("AI 操作已取消。")
        interval = min(0.1, remaining)
        time.sleep(interval)
        remaining -= interval


def _read_response(response: Any) -> bytes:
    body = response.read(MAX_RESPONSE_BYTES + 1)
    if len(body) > MAX_RESPONSE_BYTES:
        raise AIClientError(
            f"AI 响应超过 {MAX_RESPONSE_BYTES // (1024 * 1024)} MiB，拒绝继续读取。"
        )
    return body


def _thumbnail_data_url(path: Path, *, max_side: int = 1024) -> str:
    try:
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")
            image.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
            output = io.BytesIO()
            image.save(output, format="JPEG", quality=82, optimize=True)
    except (OSError, ValueError) as exc:
        raise AIClientError(f"无法读取 AI 预览图片：{path.name}：{exc}") from exc
    encoded = base64.b64encode(output.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def _vision_probe_data_url() -> str:
    image = Image.new("RGB", (64, 32), "red")
    image.paste("blue", (32, 0, 64, 32))
    output = io.BytesIO()
    image.save(output, format="PNG")
    return "data:image/png;base64," + base64.b64encode(output.getvalue()).decode(
        "ascii"
    )


def _validate_suggestion(payload: Any, item_ids: set[str]) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise AIClientError("AI 返回结果必须是 JSON 对象。")
    result = {
        "suggested_name": str(payload.get("suggested_name") or "").strip()[:160],
        "category": str(payload.get("category") or "").strip()[:240],
        "tags": [],
        "summary": str(payload.get("summary") or "").strip()[:1000],
        "merge_groups": [],
        "warnings": [],
    }
    tags = payload.get("tags", [])
    if not isinstance(tags, list):
        raise AIClientError("AI 返回的 tags 必须是列表。")
    result["tags"] = list(
        dict.fromkeys(str(item).strip()[:80] for item in tags if str(item).strip())
    )[:12]
    warnings = payload.get("warnings", [])
    if not isinstance(warnings, list):
        raise AIClientError("AI 返回的 warnings 必须是列表。")
    result["warnings"] = [
        str(item).strip()[:300] for item in warnings if str(item).strip()
    ][:12]
    groups = payload.get("merge_groups", [])
    if not isinstance(groups, list):
        raise AIClientError("AI 返回的 merge_groups 必须是列表。")
    for group in groups:
        if not isinstance(group, dict):
            raise AIClientError("AI 合并建议必须是对象。")
        group_ids = group.get("item_ids", [])
        primary_id = str(group.get("primary_id") or "")
        if not isinstance(group_ids, list):
            raise AIClientError("AI 合并建议 item_ids 必须是列表。")
        normalized_ids = list(
            dict.fromkeys(str(item) for item in group_ids if str(item) in item_ids)
        )
        if len(normalized_ids) < 2 or primary_id not in normalized_ids:
            continue
        try:
            confidence = float(group.get("confidence", 0))
        except (TypeError, ValueError):
            raise AIClientError("AI 合并建议置信度无效。") from None
        result["merge_groups"].append(
            {
                "item_ids": normalized_ids,
                "primary_id": primary_id,
                "confidence": max(0.0, min(1.0, confidence)),
                "reason": str(group.get("reason") or "").strip()[:500],
            }
        )
    return result


class OpenAICompatibleClient:
    def __init__(self, config: AIConfig) -> None:
        base_url = config.base_url.strip().rstrip("/")
        model = normalize_model_name(config.model)
        context = config.context.strip()
        if not base_url or not model:
            raise AIClientError("请先填写 AI Base URL 和模型名称。")
        try:
            parsed_url = urlparse(base_url)
            hostname = parsed_url.hostname
            _port = parsed_url.port
        except ValueError as exc:
            raise AIClientError("AI Base URL 格式无效。") from exc
        if parsed_url.scheme not in {"http", "https"} or not hostname:
            raise AIClientError("AI Base URL 必须是有效的 HTTP 或 HTTPS 地址。")
        if parsed_url.username or parsed_url.password:
            raise AIClientError("AI Base URL 不得包含用户名或密码。")
        if parsed_url.params or parsed_url.query or parsed_url.fragment:
            raise AIClientError("AI Base URL 不得包含参数、查询字符串或片段。")
        if len(context) > MAX_CONTEXT_CHARS:
            raise AIClientError(
                f"AI 业务上下文过长；请控制在 {MAX_CONTEXT_CHARS} 个字符以内。"
            )
        self.config = AIConfig(
            base_url=base_url,
            model=model,
            api_key=config.api_key.strip(),
            vision_enabled=bool(config.vision_enabled),
            timeout_seconds=max(5, int(config.timeout_seconds)),
            context=context,
        )

    def test_connection(self) -> str:
        response = self._chat(
            [
                {
                    "role": "user",
                    "content": "Reply briefly to confirm the connection works.",
                }
            ],
            json_mode=False,
        )
        confirmation = " ".join(str(response).split())
        if not confirmation:
            raise AIClientError("AI 接口返回了空内容。")
        return confirmation[:80]

    def detect_vision_support(self) -> bool:
        supported, _detail = self.probe_vision_support()
        return supported is True

    def probe_vision_support(self) -> tuple[bool | None, str]:
        try:
            response = self._chat(
                [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": (
                                    "Inspect the image. Reply with exactly the two dominant "
                                    "colors from left to right, lowercase, joined by a hyphen."
                                ),
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": _vision_probe_data_url(),
                                    "detail": "low",
                                },
                            },
                        ],
                    }
                ],
                json_mode=False,
                max_tokens=30,
            )
        except AIClientError as exc:
            return None, str(exc)
        normalized = str(response).strip().lower().replace(" ", "")
        supported = "red-blue" in normalized
        return supported, (
            "视觉输入可用" if supported else "接口已响应，但未正确识别测试图片"
        )

    def organize_media(
        self,
        items: Iterable[dict[str, Any]],
        *,
        media_kind: str,
        cancel_callback: Callable[[], bool] | None = None,
    ) -> dict[str, Any]:
        normalized = list(items)[:6]
        if not normalized:
            raise AIClientError("没有可供 AI 分析的资源。")
        ids = {str(item["id"]) for item in normalized}
        text_items = [
            {
                "id": str(item["id"]),
                "name": str(item.get("name") or ""),
                "width": int(item.get("width") or 0),
                "height": int(item.get("height") or 0),
                "format": str(item.get("format") or ""),
                "size_bytes": int(item.get("size_bytes") or 0),
                "duration_sec": float(item.get("duration_sec") or 0),
                "bitrate_kbps": int(item.get("bitrate_kbps") or 0),
                "health": str(item.get("health") or "正常"),
                "code_similarity": item.get("code_similarity"),
                "category": str(item.get("category") or "")[:240],
                "tags": [str(tag)[:80] for tag in (item.get("tags") or [])[:12]],
                "summary": str(item.get("summary") or "")[:500],
            }
            for item in normalized
        ]
        selected_id = text_items[0]["id"]
        visual_context = (
            "提供了可用预览时才可引用画面内容。"
            if self.config.vision_enabled
            else "未提供任何画面；不得声称看过图片或视频内容。"
        )
        instructions = (
            "你是本地文档媒体资产整理顾问。列表第一项是当前选中资源，其余仅为"
            "代码规则筛出的候选。suggested_name、category、tags 和 summary 只针对"
            f"当前资源 ID {selected_id}。请基于可靠证据判断哪些条目属于同一内容，"
            "并推荐其中最适合作为主资源的条目。主资源优先完整、清晰、健康、"
            "兼容且分辨率合理的版本。代码相似度和精确哈希规则优先于视觉印象。"
            "仅题材、客户或场景相似不得合并；证据不足时 merge_groups 必须为空。"
            f"{visual_context}仅返回 JSON，结构为："
            '{"suggested_name":"","category":"","tags":[],"summary":"",'
            '"merge_groups":[{"item_ids":[],"primary_id":"","confidence":0.0,'
            '"reason":""}],"warnings":[]}。'
            f"\n资源类型：{media_kind}\n条目："
            + json.dumps(text_items, ensure_ascii=False)
        )
        if self.config.context:
            instructions += f"\n业务上下文：{self.config.context}"
        content: list[dict[str, Any]] = [{"type": "text", "text": instructions}]
        for item in normalized:
            if not self.config.vision_enabled:
                break
            preview_path = item.get("preview_path")
            if not preview_path:
                continue
            path = Path(str(preview_path))
            if not path.is_file():
                continue
            content.append(
                {
                    "type": "text",
                    "text": f"下面预览对应条目 ID：{item['id']}",
                }
            )
            content.append(
                {
                    "type": "image_url",
                    "image_url": {"url": _thumbnail_data_url(path), "detail": "low"},
                }
            )
        payload = self._chat(
            [{"role": "user", "content": content}],
            max_tokens=1200,
            cancel_callback=cancel_callback,
        )
        return _validate_suggestion(payload, ids)

    def _chat(
        self,
        messages: list[dict[str, Any]],
        *,
        json_mode: bool = True,
        max_tokens: int | None = None,
        cancel_callback: Callable[[], bool] | None = None,
    ) -> Any:
        payload: dict[str, Any] = {
            "model": self.config.model,
            "messages": messages,
            "temperature": 0.1,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens
        response_body = self._request(payload, cancel_callback=cancel_callback)
        try:
            envelope = json.loads(response_body.decode("utf-8"))
            message = envelope["choices"][0]["message"]
            content = message.get("content")
            if isinstance(content, list):
                content = "".join(
                    str(item.get("text") or "")
                    for item in content
                    if isinstance(item, dict)
                )
            if not content:
                content = message.get("reasoning_content")
            text = str(content or "").strip()
            if not text:
                raise AIClientError("AI 接口返回了空内容。")
            if not json_mode:
                return text
            if text.startswith("```"):
                text = text.removeprefix("```json").removeprefix("```JSON")
                text = text.removeprefix("```").removesuffix("```").strip()
            return json.loads(text)
        except AIClientError:
            raise
        except (
            KeyError,
            IndexError,
            TypeError,
            ValueError,
            json.JSONDecodeError,
        ) as exc:
            raise AIClientError("AI 响应不是有效的 OpenAI 兼容 JSON 结果。") from exc

    def _request(
        self,
        payload: dict[str, Any],
        *,
        cancel_callback: Callable[[], bool] | None = None,
    ) -> bytes:
        allow_json_fallback = "response_format" in payload
        transient_retries = 0
        while True:
            if cancel_callback and cancel_callback():
                raise AIClientError("AI 操作已取消。")
            request = urllib.request.Request(
                f"{self.config.base_url}/chat/completions",
                data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                method="POST",
                headers={
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "User-Agent": f"Doc-Media-Toolkit/{__version__}",
                    **(
                        {"Authorization": f"Bearer {self.config.api_key}"}
                        if self.config.api_key
                        else {}
                    ),
                },
            )
            try:
                # The constructor rejects every scheme except HTTP(S) before this call.
                with urllib.request.urlopen(  # nosec B310
                    request, timeout=self.config.timeout_seconds
                ) as response:
                    return _read_response(response)
            except urllib.error.HTTPError as exc:
                detail = exc.read(8192).decode("utf-8", errors="replace")[:500]
                exc.close()
                detail = _redact_secret(detail, self.config.api_key)
                lowered = detail.lower()
                if (
                    allow_json_fallback
                    and exc.code in {400, 422}
                    and any(
                        marker in lowered
                        for marker in ("response_format", "json mode", "json_object")
                    )
                ):
                    payload = {
                        key: value
                        for key, value in payload.items()
                        if key != "response_format"
                    }
                    allow_json_fallback = False
                    continue
                if (
                    exc.code in {429, 502, 503, 504}
                    and transient_retries < TRANSIENT_RETRIES
                ):
                    delay = _retry_delay(exc.headers, transient_retries)
                    transient_retries += 1
                    _wait_before_retry(delay, cancel_callback)
                    continue
                if exc.code == 401:
                    message = "AI 认证失败：请在设置中重新填写本次运行使用的 API Key。"
                elif exc.code == 403:
                    message = (
                        "AI 服务拒绝访问：请检查 API Key 权限、Base URL 和服务套餐。"
                    )
                elif exc.code == 429:
                    message = "AI 请求过于频繁或额度不足，请稍后重试。"
                else:
                    message = f"AI 请求失败（HTTP {exc.code}）：{detail}"
                raise AIClientError(message) from exc
            except (TimeoutError, urllib.error.URLError, OSError) as exc:
                if transient_retries < TRANSIENT_RETRIES:
                    delay = _retry_delay(None, transient_retries)
                    transient_retries += 1
                    _wait_before_retry(delay, cancel_callback)
                    continue
                if isinstance(exc, TimeoutError):
                    message = (
                        f"AI 响应超时（{self.config.timeout_seconds} 秒），请稍后重试。"
                    )
                else:
                    reason = getattr(exc, "reason", exc)
                    message = (
                        "无法连接 AI 服务："
                        f"{_redact_secret(reason, self.config.api_key)}"
                    )
                raise AIClientError(message) from exc
