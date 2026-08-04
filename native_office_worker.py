#!/usr/bin/env python3
# Copyright (C) 2026 Leon Priest
# SPDX-License-Identifier: GPL-3.0-or-later

"""Isolated Microsoft Office PDF export worker.

This program is launched as a separate process by ``office2pdf.py``.  Keeping
COM automation out of the GUI process gives the parent a real timeout boundary
and lets it terminate only the dedicated Office instance that this worker
created.  It is not intended to be launched directly by users.
"""

from __future__ import annotations

import ctypes
from ctypes import wintypes
import sys
from pathlib import Path


BACKENDS = {"word", "excel", "powerpoint"}


def _window_handle(obj) -> int | None:
    """Return an Office application/window handle when the object exposes one."""
    for attribute in ("Hwnd", "HWND"):
        try:
            hwnd = int(getattr(obj, attribute))
        except Exception:
            continue
        if hwnd:
            return hwnd
    return None


def _pid_from_hwnd(hwnd: int) -> int:
    """Return the process ID that owns ``hwnd``."""
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    get_window_thread_process_id = user32.GetWindowThreadProcessId
    get_window_thread_process_id.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    get_window_thread_process_id.restype = wintypes.DWORD

    pid = wintypes.DWORD(0)
    thread_id = get_window_thread_process_id(wintypes.HWND(hwnd), ctypes.byref(pid))
    if not thread_id or not pid.value:
        raise RuntimeError("Could not determine the dedicated Office process ID")
    return int(pid.value)


def _record_pid_from_objects(pid_file: Path, *objects) -> int:
    """Record the PID owned by the first Office object exposing a valid HWND."""
    for obj in objects:
        if obj is None:
            continue
        hwnd = _window_handle(obj)
        if not hwnd:
            continue
        pid = _pid_from_hwnd(hwnd)
        pid_file.write_text(str(pid), encoding="ascii")
        return hwnd
    raise RuntimeError("Office did not expose a usable application or document window handle")


def _hide_window(hwnd: int) -> None:
    """Best-effort immediate hide for Word's bounded visibility probe."""
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        show_window = user32.ShowWindow
        show_window.argtypes = [wintypes.HWND, ctypes.c_int]
        show_window.restype = wintypes.BOOL
        show_window(wintypes.HWND(hwnd), 0)  # SW_HIDE
    except Exception:
        pass


def _record_word_pid(word, document, pid_file: Path) -> None:
    """Record the dedicated Word PID without requiring a permanently visible UI.

    Word does not consistently expose ``Application.Hwnd``.  Its document
    ``Window.Hwnd`` is the reliable handle, but some Office builds withhold UI
    properties while the automation application is hidden.  First try the
    hidden document window.  Only if that fails, make the dedicated DispatchEx
    instance visible briefly, capture its document-window handle, and hide it
    immediately.  No unrelated Word instance is inspected or terminated.
    """
    try:
        active_window = document.ActiveWindow
    except Exception:
        active_window = None

    try:
        _record_pid_from_objects(pid_file, word, active_window)
        return
    except RuntimeError:
        pass

    try:
        original_visible = bool(word.Visible)
    except Exception:
        original_visible = False

    hwnd = None
    try:
        # Minimise first where supported to reduce the chance of a visible flash.
        try:
            word.WindowState = 2  # wdWindowStateMinimize
        except Exception:
            pass
        word.Visible = True

        # Word can publish the window asynchronously. Keep this probe short and
        # bounded; the parent process still owns the overall conversion deadline.
        import time

        deadline = time.monotonic() + 2.0
        while time.monotonic() < deadline:
            try:
                active_window = document.ActiveWindow
            except Exception:
                active_window = None
            try:
                hwnd = _record_pid_from_objects(pid_file, word, active_window)
                break
            except RuntimeError:
                time.sleep(0.05)

        if not hwnd:
            raise RuntimeError(
                "Word did not expose a usable document window handle, even during "
                "the bounded visibility probe"
            )
    finally:
        if hwnd:
            _hide_window(hwnd)
        try:
            word.Visible = original_visible
        except Exception:
            try:
                word.Visible = False
            except Exception:
                pass


def _force_disable_macros(app) -> None:
    """Require Office to disable macros before opening any input document."""
    try:
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception as exc:
        raise RuntimeError("Could not force-disable Office macros") from exc


def _convert_word(src: Path, dest_pdf: Path, pid_file: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        _force_disable_macros(word)
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.Options.SaveNormalPrompt = False
        except Exception:
            pass
        document = word.Documents.Open(
            str(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        _record_word_pid(word, document, pid_file)
        document.ExportAsFixedFormat(str(dest_pdf), 17)  # wdExportFormatPDF
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_excel(src: Path, dest_pdf: Path, pid_file: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        _record_pid_from_objects(pid_file, excel)
        _force_disable_macros(excel)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(src),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        workbook.ExportAsFixedFormat(0, str(dest_pdf))  # xlTypePDF
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_powerpoint(src: Path, dest_pdf: Path, pid_file: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        _record_pid_from_objects(pid_file, powerpoint)
        _force_disable_macros(powerpoint)
        presentation = powerpoint.Presentations.Open(
            str(src), ReadOnly=True, Untitled=False, WithWindow=False
        )
        presentation.SaveAs(str(dest_pdf), 32)  # ppSaveAsPDF
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) != 4 or args[0] not in BACKENDS:
        print(
            "usage: native_office_worker.py <word|excel|powerpoint> "
            "<source> <destination.pdf> <pid-file>",
            file=sys.stderr,
        )
        return 2

    backend, source_raw, destination_raw, pid_raw = args
    source = Path(source_raw).resolve()
    destination = Path(destination_raw).resolve()
    pid_file = Path(pid_raw).resolve()

    if not source.is_file():
        print(f"source does not exist: {source}", file=sys.stderr)
        return 2

    converters = {
        "word": _convert_word,
        "excel": _convert_excel,
        "powerpoint": _convert_powerpoint,
    }

    try:
        converters[backend](source, destination, pid_file)
    except Exception as exc:
        print(f"{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
