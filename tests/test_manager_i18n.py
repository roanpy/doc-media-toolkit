"""Smoke tests for the video/image library i18n layer."""

from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pptx_tools import manager_i18n


class ManagerI18nTest(unittest.TestCase):
    def tearDown(self) -> None:
        manager_i18n.set_language("zh")

    def test_chinese_passthrough(self) -> None:
        manager_i18n.set_language("zh")
        self.assertEqual(manager_i18n.tr("取消"), "取消")

    def test_english_lookup_and_fallback(self) -> None:
        manager_i18n.set_language("en")
        self.assertEqual(manager_i18n.tr("取消"), "Cancel")
        # Untranslated keys fall back to the Chinese source unchanged.
        self.assertEqual(manager_i18n.tr("某条没有翻译的文案"), "某条没有翻译的文案")

    def test_empty_string(self) -> None:
        manager_i18n.set_language("en")
        self.assertEqual(manager_i18n.tr(""), "")

    def test_windows_smoke_instantiate_both_languages(self) -> None:
        from PySide6.QtWidgets import QApplication

        app = QApplication.instance() or QApplication([])

        from pptx_tools.image_manager_gui import MainWindow as ImageWindow
        from pptx_tools.video_manager_gui import MainWindow as VideoWindow

        for lang in ("zh", "en"):
            manager_i18n.set_language(lang)
            for cls in (VideoWindow, ImageWindow):
                window = cls()
                window.show()
                app.processEvents()
                window.close()
                window.deleteLater()


if __name__ == "__main__":
    unittest.main()
