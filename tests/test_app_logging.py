from __future__ import annotations

import logging
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx_tools.app_logging import (
    LOGGER_NAMES,
    app_dir_name,
    configure_app_logging,
    write_ai_audit_event,
)


class AppLoggingTest(unittest.TestCase):
    def test_experimental_logs_use_isolated_application_directory(self) -> None:
        with patch.dict(os.environ, {"PPTX_TOOLS_EXPERIMENTAL": "1"}):
            self.assertEqual(app_dir_name(), "Doc Media Toolkit Experimental")

    def test_ai_audit_records_action_without_credentials(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            with patch("pptx_tools.app_logging.log_directory", return_value=directory):
                write_ai_audit_event(
                    media_kind="image",
                    target_id="asset-1",
                    provider="https://example.test/v1",
                    model="model",
                    vision_enabled=False,
                    applied_fields=["名称"],
                    merge_group_count=0,
                )
            record = json.loads(
                (directory / "ai-audit.jsonl").read_text(encoding="utf-8")
            )
            self.assertEqual(record["target_id"], "asset-1")
            self.assertNotIn("api_key", record)

    def test_all_tool_logs_are_written_to_rotating_log(self) -> None:
        snapshots = {
            name: (
                list(logging.getLogger(name).handlers),
                logging.getLogger(name).level,
                logging.getLogger(name).propagate,
            )
            for name in LOGGER_NAMES
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            try:
                for name in LOGGER_NAMES:
                    logging.getLogger(name).handlers.clear()
                with patch(
                    "pptx_tools.app_logging.log_directory", return_value=directory
                ):
                    log_path = configure_app_logging()
                    logging.getLogger("pptx_output_watermark.gui").info("watermark-log")
                    logging.getLogger("pptx_video_compactor_gui").info("compactor-log")
                    for handler in logging.getLogger(LOGGER_NAMES[0]).handlers:
                        handler.flush()
                content = log_path.read_text(encoding="utf-8")
                self.assertIn("watermark-log", content)
                self.assertIn("compactor-log", content)
            finally:
                current = {
                    handler
                    for name in LOGGER_NAMES
                    for handler in logging.getLogger(name).handlers
                    if handler not in snapshots[name][0]
                }
                for name, (handlers, level, propagate) in snapshots.items():
                    logger = logging.getLogger(name)
                    logger.handlers[:] = handlers
                    logger.setLevel(level)
                    logger.propagate = propagate
                for handler in current:
                    handler.close()

    def test_ai_audit_log_rotates_before_append(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            path = directory / "ai-audit.jsonl"
            path.write_text("old-record\n", encoding="utf-8")
            with (
                patch(
                    "pptx_tools.app_logging.log_directory",
                    return_value=directory,
                ),
                patch("pptx_tools.app_logging.AI_AUDIT_MAX_BYTES", 1),
            ):
                write_ai_audit_event(
                    media_kind="video",
                    target_id="target",
                    provider="https://example.test/v1",
                    model="model",
                    vision_enabled=False,
                    applied_fields=[],
                    merge_group_count=0,
                )
            self.assertEqual(
                (directory / "ai-audit.jsonl.1").read_text(encoding="utf-8"),
                "old-record\n",
            )
            self.assertIn('"target_id": "target"', path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
