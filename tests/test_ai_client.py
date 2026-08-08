from __future__ import annotations

import json
import tempfile
import unittest
import urllib.error
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from pptx_tools import __version__
from pptx_tools.ai_client import (
    AIClientError,
    AIConfig,
    MAX_CONTEXT_CHARS,
    MAX_RESPONSE_BYTES,
    OpenAICompatibleClient,
    privacy_scope,
)


class FakeResponse:
    def __init__(self, payload: dict) -> None:
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *_args) -> None:
        return None

    def read(self, size: int = -1) -> bytes:
        body = json.dumps(self.payload).encode("utf-8")
        return body if size < 0 else body[:size]


def completion(content: object) -> dict:
    return {"choices": [{"message": {"content": json.dumps(content)}}]}


def text_completion(content: str) -> dict:
    return {"choices": [{"message": {"content": content}}]}


class AIClientTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = OpenAICompatibleClient(
            AIConfig("https://example.test/v1", "vision-model", "secret")
        )

    def test_test_connection_accepts_openai_compatible_json(self) -> None:
        with patch(
            "urllib.request.urlopen",
            return_value=FakeResponse(text_completion("OK，连接正常。")),
        ) as request:
            self.assertEqual(self.client.test_connection(), "OK，连接正常。")
        sent = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertNotIn("response_format", sent)
        self.assertNotIn("max_tokens", sent)
        self.assertEqual(
            request.call_args.args[0].headers["User-agent"],
            f"Doc-Media-Toolkit/{__version__}",
        )

    def test_cancelled_media_request_does_not_open_network_connection(self) -> None:
        with patch("urllib.request.urlopen") as request:
            with self.assertRaisesRegex(AIClientError, "已取消"):
                self.client.organize_media(
                    [{"id": "a", "name": "one"}],
                    media_kind="图片",
                    cancel_callback=lambda: True,
                )
        request.assert_not_called()

    def test_detect_vision_support_uses_image_content(self) -> None:
        with patch(
            "urllib.request.urlopen",
            return_value=FakeResponse(text_completion("red-blue")),
        ) as request:
            self.assertTrue(self.client.detect_vision_support())
        sent = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertIn("data:image/png;base64,", json.dumps(sent))
        self.assertNotIn("response_format", sent)

    def test_organize_media_validates_merge_and_primary(self) -> None:
        payload = {
            "suggested_name": "工厂产线",
            "category": "智能制造",
            "tags": ["工厂", "产线", "工厂"],
            "summary": "同一场景的两个版本。",
            "merge_groups": [
                {
                    "item_ids": ["a", "b"],
                    "primary_id": "b",
                    "confidence": 0.91,
                    "reason": "b 更清晰",
                },
                {
                    "item_ids": ["a", "missing"],
                    "primary_id": "missing",
                    "confidence": 1,
                    "reason": "无效",
                },
            ],
            "warnings": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            preview = Path(temp_dir) / "preview.png"
            Image.new("RGB", (30, 20), (10, 20, 30)).save(preview)
            with patch(
                "urllib.request.urlopen",
                return_value=FakeResponse(completion(payload)),
            ) as request:
                result = self.client.organize_media(
                    [
                        {
                            "id": "a",
                            "name": "one",
                            "width": 30,
                            "height": 20,
                            "preview_path": preview,
                        },
                        {"id": "b", "name": "two", "width": 60, "height": 40},
                    ],
                    media_kind="图片",
                )

        self.assertEqual(result["tags"], ["工厂", "产线"])
        self.assertEqual(len(result["merge_groups"]), 1)
        self.assertEqual(result["merge_groups"][0]["primary_id"], "b")
        sent = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertEqual(sent["model"], "vision-model")
        self.assertIn("image_url", json.dumps(sent))

    def test_invalid_response_fails_without_exposing_api_key(self) -> None:
        with patch(
            "urllib.request.urlopen",
            return_value=FakeResponse({"choices": [{"message": {"content": "  "}}]}),
        ):
            with self.assertRaises(AIClientError) as raised:
                self.client.test_connection()
        self.assertNotIn("secret", str(raised.exception))

    def test_unauthorized_response_has_actionable_message(self) -> None:
        error = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            401,
            "Unauthorized",
            {},
            BytesIO(b'{"error":"Missing API key"}'),
        )
        with patch("urllib.request.urlopen", side_effect=error):
            with self.assertRaises(AIClientError) as raised:
                self.client.test_connection()
        self.assertIn("重新填写", str(raised.exception))
        self.assertNotIn("Traceback", str(raised.exception))

    def test_text_only_model_uses_metadata_without_sending_preview(self) -> None:
        client = OpenAICompatibleClient(
            AIConfig(
                "https://example.test/v1",
                "text-model",
                "secret",
                vision_enabled=False,
            )
        )
        with patch(
            "urllib.request.urlopen",
            return_value=FakeResponse(completion({"merge_groups": []})),
        ) as request:
            client.organize_media(
                [{"id": "a", "name": "one", "preview_path": "/not/read.jpg"}],
                media_kind="视频",
            )
        sent = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertNotIn("image_url", json.dumps(sent))
        prompt = json.dumps(sent, ensure_ascii=False)
        self.assertIn("第一项是当前选中资源", prompt)
        self.assertIn("不得声称看过图片或视频内容", prompt)

    def test_normalizes_model_dashes_and_includes_context(self) -> None:
        client = OpenAICompatibleClient(
            AIConfig(
                "https://opencode.ai/zen/go/v1",
                "deepseek–v4–flash",
                context="按客户和年份分类",
            )
        )
        with patch(
            "urllib.request.urlopen",
            return_value=FakeResponse(completion({"merge_groups": []})),
        ) as request:
            client.organize_media([{"id": "a", "name": "one"}], media_kind="图片")
        sent = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertEqual(sent["model"], "deepseek-v4-flash")
        self.assertIn("按客户和年份分类", json.dumps(sent, ensure_ascii=False))

    def test_json_mode_falls_back_and_accepts_fenced_json(self) -> None:
        unsupported = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            400,
            "Bad Request",
            {},
            BytesIO(b'{"error":"response_format is not supported"}'),
        )
        response = FakeResponse(
            text_completion(
                '```json\n{"merge_groups": [], "suggested_name": "new"}\n```'
            )
        )
        with patch(
            "urllib.request.urlopen", side_effect=[unsupported, response]
        ) as request:
            result = self.client.organize_media(
                [{"id": "a", "name": "one"}], media_kind="图片"
            )
        self.assertEqual(result["suggested_name"], "new")
        first = json.loads(request.call_args_list[0].args[0].data.decode("utf-8"))
        second = json.loads(request.call_args_list[1].args[0].data.decode("utf-8"))
        self.assertIn("response_format", first)
        self.assertNotIn("response_format", second)

    def test_transient_http_error_retries_once(self) -> None:
        busy = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            503,
            "Unavailable",
            {},
            BytesIO(b"busy"),
        )
        with (
            patch(
                "urllib.request.urlopen",
                side_effect=[busy, FakeResponse(text_completion("OK"))],
            ) as request,
            patch("pptx_tools.ai_client.time.sleep"),
        ):
            self.assertEqual(self.client.test_connection(), "OK")
        self.assertEqual(request.call_count, 2)

    def test_retry_after_header_controls_transient_delay(self) -> None:
        busy = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            429,
            "Busy",
            {"Retry-After": "3"},
            BytesIO(b"busy"),
        )
        with (
            patch(
                "urllib.request.urlopen",
                side_effect=[busy, FakeResponse(text_completion("OK"))],
            ),
            patch("pptx_tools.ai_client.time.sleep") as sleep,
        ):
            self.assertEqual(self.client.test_connection(), "OK")
        self.assertAlmostEqual(sum(call.args[0] for call in sleep.call_args_list), 3.0)

    def test_network_error_redacts_api_key(self) -> None:
        with (
            patch(
                "urllib.request.urlopen",
                side_effect=urllib.error.URLError(
                    "proxy rejected https://secret@example.test"
                ),
            ),
            patch("pptx_tools.ai_client.time.sleep"),
            self.assertRaises(AIClientError) as raised,
        ):
            self.client.test_connection()
        self.assertNotIn("secret", str(raised.exception))
        self.assertIn("***", str(raised.exception))

    def test_json_fallback_keeps_transient_retry_available(self) -> None:
        unsupported = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            400,
            "Bad Request",
            {},
            BytesIO(b'{"error":"response_format is not supported"}'),
        )
        busy = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            503,
            "Unavailable",
            {},
            BytesIO(b"busy"),
        )
        response = FakeResponse(completion({"merge_groups": []}))
        with (
            patch(
                "urllib.request.urlopen",
                side_effect=[unsupported, busy, response],
            ) as request,
            patch("pptx_tools.ai_client.time.sleep"),
        ):
            self.client.organize_media([{"id": "a", "name": "one"}], media_kind="图片")
        self.assertEqual(request.call_count, 3)

    def test_context_length_is_bounded_before_request(self) -> None:
        with self.assertRaisesRegex(AIClientError, "业务上下文过长"):
            OpenAICompatibleClient(
                AIConfig(
                    "https://example.test/v1",
                    "model",
                    context="x" * (MAX_CONTEXT_CHARS + 1),
                )
            )

    def test_base_url_rejects_unsafe_or_ambiguous_urls(self) -> None:
        for value in (
            "file:///tmp/api",
            "example.test/v1",
            "https://user:secret@example.test/v1",
            "https://example.test/v1?tenant=secret",
        ):
            with self.subTest(value=value), self.assertRaises(AIClientError):
                OpenAICompatibleClient(AIConfig(value, "model"))

    def test_response_size_is_bounded(self) -> None:
        response = FakeResponse(text_completion("x" * MAX_RESPONSE_BYTES))
        with patch("urllib.request.urlopen", return_value=response):
            with self.assertRaisesRegex(AIClientError, "响应超过"):
                self.client.test_connection()

    def test_retry_wait_can_be_cancelled(self) -> None:
        busy = urllib.error.HTTPError(
            "https://example.test/v1/chat/completions",
            503,
            "Unavailable",
            {"Retry-After": "3"},
            BytesIO(b"busy"),
        )
        checks = iter([False, True])
        with (
            patch("urllib.request.urlopen", side_effect=busy),
            patch("pptx_tools.ai_client.time.sleep"),
            self.assertRaisesRegex(AIClientError, "已取消"),
        ):
            self.client._request({}, cancel_callback=lambda: next(checks))

    def test_vision_probe_keeps_unknown_separate_from_unsupported(self) -> None:
        with (
            patch(
                "urllib.request.urlopen",
                side_effect=urllib.error.URLError("offline"),
            ),
            patch("pptx_tools.ai_client.time.sleep"),
        ):
            supported, detail = self.client.probe_vision_support()
        self.assertIsNone(supported)
        self.assertIn("无法连接", detail)

    def test_privacy_scope_changes_with_visual_mode(self) -> None:
        text = privacy_scope(
            AIConfig("https://example.test/v1", "model", vision_enabled=False)
        )
        vision = privacy_scope(
            AIConfig("https://example.test/v1", "model", vision_enabled=True)
        )
        self.assertNotEqual(text, vision)

    def test_privacy_scope_never_contains_url_credentials(self) -> None:
        scope = privacy_scope(
            AIConfig("https://user:secret@example.test:8443/v1", "model")
        )
        self.assertEqual(scope, "example.test:8443|model|vision")
        self.assertNotIn("secret", scope)


if __name__ == "__main__":
    unittest.main()
