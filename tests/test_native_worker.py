from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import native_office_worker


class _WordApplicationWithoutHwnd:
    def __init__(self) -> None:
        self.Visible = False
        self.WindowState = 0


class _StaticWindow:
    Hwnd = 12345


class _StaticDocument:
    ActiveWindow = _StaticWindow()


class _DynamicWindow:
    def __init__(self, app: _WordApplicationWithoutHwnd) -> None:
        self._app = app

    @property
    def Hwnd(self) -> int:
        return 24680 if self._app.Visible else 0


class _DynamicDocument:
    def __init__(self, app: _WordApplicationWithoutHwnd) -> None:
        self.ActiveWindow = _DynamicWindow(app)


class NativeWordPidTests(unittest.TestCase):
    def test_word_uses_document_window_when_application_has_no_hwnd(self) -> None:
        with tempfile.TemporaryDirectory() as tempdir:
            pid_file = Path(tempdir) / "pid.txt"
            word = _WordApplicationWithoutHwnd()
            document = _StaticDocument()

            with mock.patch.object(native_office_worker, "_pid_from_hwnd", return_value=4321):
                native_office_worker._record_word_pid(word, document, pid_file)

            self.assertEqual(pid_file.read_text(encoding="ascii"), "4321")
            self.assertFalse(word.Visible)

    def test_word_visibility_probe_is_bounded_and_restores_hidden_state(self) -> None:
        with tempfile.TemporaryDirectory() as tempdir:
            pid_file = Path(tempdir) / "pid.txt"
            word = _WordApplicationWithoutHwnd()
            document = _DynamicDocument(word)

            with (
                mock.patch.object(native_office_worker, "_pid_from_hwnd", return_value=8765),
                mock.patch.object(native_office_worker, "_hide_window") as hide_window,
            ):
                native_office_worker._record_word_pid(word, document, pid_file)

            self.assertEqual(pid_file.read_text(encoding="ascii"), "8765")
            self.assertFalse(word.Visible)
            hide_window.assert_called_once_with(24680)


if __name__ == "__main__":
    unittest.main()
