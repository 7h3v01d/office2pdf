from __future__ import annotations

import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ICON_PATH = PROJECT_ROOT / "assets" / "office2pdf.ico"
BUILD_SCRIPT = PROJECT_ROOT / "build_exe.bat"


class BrandingTests(unittest.TestCase):
    def test_icon_is_valid_multi_resolution_windows_resource(self) -> None:
        data = ICON_PATH.read_bytes()
        self.assertGreaterEqual(len(data), 6)
        self.assertEqual(int.from_bytes(data[0:2], "little"), 0)
        self.assertEqual(int.from_bytes(data[2:4], "little"), 1)
        self.assertGreaterEqual(int.from_bytes(data[4:6], "little"), 4)

    def test_windows_build_embeds_and_bundles_icon(self) -> None:
        script = BUILD_SCRIPT.read_text(encoding="utf-8")
        self.assertIn('--icon "%CD%\\assets\\office2pdf.ico"', script)
        self.assertIn(
            '--add-data "%CD%\\assets\\office2pdf.ico;assets"',
            script,
        )
        self.assertIn("Create Office2PDF Shortcuts.bat", script)
        self.assertIn('--version-file "%CD%\\windows_version_info.txt"', script)


if __name__ == "__main__":
    unittest.main()
