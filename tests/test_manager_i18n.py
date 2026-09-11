"""Smoke tests for the video/image library i18n layer."""

from __future__ import annotations

import ast
import os
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pptx_tools import manager_i18n


ROOT = Path(__file__).resolve().parents[1]
GUI_MODULES = (
    ROOT / "src" / "pptx_tools" / "video_manager_gui.py",
    ROOT / "src" / "pptx_tools" / "image_manager_gui.py",
)


def _literal_tr_keys(path: Path) -> list[tuple[int, str]]:
    """Return every literal Chinese string passed directly to ``tr()``."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found: list[tuple[int, str]] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not node.args:
            continue
        func = node.func
        if not isinstance(func, ast.Name) or func.id != "tr":
            continue
        argument = node.args[0]
        if not isinstance(argument, ast.Constant) or not isinstance(
            argument.value, str
        ):
            continue
        if any("\u4e00" <= char <= "\u9fff" for char in argument.value):
            found.append((node.lineno, argument.value))
    return found


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

    def test_every_library_literal_has_an_english_translation(self) -> None:
        untranslated: list[str] = []
        for path in GUI_MODULES:
            for line, value in _literal_tr_keys(path):
                if value not in manager_i18n.TRANSLATIONS:
                    untranslated.append(f"{path.name}:{line}: {value!r}")
        self.assertEqual(
            untranslated,
            [],
            "Library UI literals need an English entry in manager_i18n.TRANSLATIONS.",
        )

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
