from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import package_source
import version_info


PROJECT_ROOT = Path(__file__).resolve().parents[1]
GUI_PATH = PROJECT_ROOT / "gui.py"
BUILD_SCRIPT = PROJECT_ROOT / "build_exe.bat"
SHORTCUT_SCRIPT = PROJECT_ROOT / "create_shortcuts.bat"
LICENSE_PATH = PROJECT_ROOT / "LICENSE.txt"
NOTICES_PATH = PROJECT_ROOT / "THIRD_PARTY_NOTICES.txt"
VERSION_RESOURCE = PROJECT_ROOT / "windows_version_info.txt"


class ProfessionalReleaseTests(unittest.TestCase):
    def test_release_metadata_is_consistent(self) -> None:
        self.assertEqual(version_info.APP_VERSION, "1.0.0")
        self.assertEqual(version_info.APP_VERSION_TUPLE, (1, 0, 0, 0))
        self.assertEqual(version_info.SPLASH_DURATION_MS, 5000)
        self.assertEqual(version_info.APP_LICENSE_ID, "GPL-3.0-or-later")

        version_resource = VERSION_RESOURCE.read_text(encoding="utf-8")
        self.assertIn("filevers=(1, 0, 0, 0)", version_resource)
        self.assertIn("ProductVersion', u'1.0.0'", version_resource)

    def test_gui_contains_five_second_splash_and_about_tab(self) -> None:
        source = GUI_PATH.read_text(encoding="utf-8")
        self.assertIn("class StartupSplash(QWidget)", source)
        self.assertIn("QTimer.singleShot(SPLASH_DURATION_MS", source)
        self.assertIn('self.main_tabs.addTab(self._build_about_page(), "About")', source)
        self.assertIn('info_tabs.addTab(licence, "GPL Licence")', source)
        self.assertIn('info_tabs.addTab(notices, "Third-Party Notices")', source)

    def test_licence_and_notices_are_present(self) -> None:
        licence = LICENSE_PATH.read_text(encoding="utf-8")
        notices = NOTICES_PATH.read_text(encoding="utf-8")
        self.assertIn("GNU GENERAL PUBLIC LICENSE", licence)
        self.assertIn("Version 3, 29 June 2007", licence)
        self.assertIn("PyQt6", notices)
        self.assertIn("LibreOffice", notices)
        self.assertIn("Microsoft Office", notices)

    def test_build_bundles_professional_release_materials(self) -> None:
        script = BUILD_SCRIPT.read_text(encoding="utf-8")
        self.assertIn('--version-file "%CD%\\windows_version_info.txt"', script)
        self.assertIn('--add-data "%CD%\\LICENSE.txt;."', script)
        self.assertIn('Office2PDF-1.0.0-source.zip', script)
        self.assertIn('Create Office2PDF Shortcuts.bat', script)

    def test_shortcut_creator_avoids_powershell_and_wrong_build_exe(self) -> None:
        script = SHORTCUT_SCRIPT.read_text(encoding="utf-8")
        self.assertNotIn("powershell", script.casefold())
        self.assertIn("_internal\\python311.dll", script)
        self.assertIn("Do not use build\\Office2PDF\\Office2PDF.exe", script)
        self.assertIn("%SystemRoot%\\System32\\cscript.exe", script)

    def test_source_archive_contains_corresponding_source(self) -> None:
        with tempfile.TemporaryDirectory() as tempdir:
            output = Path(tempdir) / "source.zip"
            package_source.create_source_archive(output)
            with ZipFile(output) as archive:
                names = set(archive.namelist())

        root = "Office2PDF-1.0.0-source/"
        self.assertIn(root + "gui.py", names)
        self.assertIn(root + "office2pdf.py", names)
        self.assertIn(root + "native_office_worker.py", names)
        self.assertIn(root + "LICENSE.txt", names)
        self.assertIn(root + "tests/test_office2pdf.py", names)


if __name__ == "__main__":
    unittest.main()
