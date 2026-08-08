from __future__ import annotations

import ast
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def imported_modules(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    modules: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            modules.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            modules.add(node.module)
    return modules


class ArchitectureBoundaryTests(unittest.TestCase):
    def test_asset_manager_guis_depend_on_shared_ui_not_each_other(self) -> None:
        image_imports = imported_modules(ROOT / "src/pptx_tools/image_manager_gui.py")
        video_imports = imported_modules(ROOT / "src/pptx_tools/video_manager_gui.py")

        self.assertIn("pptx_tools.media_manager_ui", image_imports)
        self.assertIn("pptx_tools.media_manager_ui", video_imports)
        self.assertIn("pptx_tools.ui_theme", image_imports)
        self.assertIn("pptx_tools.ui_theme", video_imports)
        self.assertNotIn("pptx_tools.video_manager_gui", image_imports)
        self.assertNotIn("pptx_tools.image_manager_gui", video_imports)
        self.assertNotIn("pptx_video_compactor_gui", image_imports)
        self.assertNotIn("pptx_video_compactor_gui", video_imports)


if __name__ == "__main__":
    unittest.main()
