#!/usr/bin/env python3
# Copyright (C) 2026 Leon
# SPDX-License-Identifier: GPL-3.0-or-later

"""
office2pdf GUI
==============

Industrial-dark PyQt6 front end for office2pdf.py.

Drag files or folders onto the queue, set your options, hit Convert.
The heavy lifting (soffice invocation, retries, timeouts, collision
detection, success verification) is all done by the office2pdf module
sitting next to this file -- this GUI is a thin, honest wrapper around
it. Nothing here silently swallows a failure; every file gets an
explicit OK / SKIP / FAIL line in the log, same philosophy as the CLI.

Run:
    python gui.py

Requires:
    PyQt6
    office2pdf.py in the same directory
    LibreOffice (soffice) reachable on PATH (or default Windows install path)
"""

from __future__ import annotations

import os
import sys
import subprocess
import concurrent.futures
import threading
import time
from collections import deque
from pathlib import Path
from dataclasses import dataclass

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMimeData, QTimer
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QStackedWidget,
    QGroupBox, QCheckBox, QSpinBox, QLineEdit, QFileDialog, QPlainTextEdit,
    QProgressBar, QSplitter, QAbstractItemView, QFrame,
    QMessageBox, QStyle, QTabWidget, QTextBrowser,
)

import office2pdf as backend
from version_info import (
    APP_AUTHOR,
    APP_BUILD_DATE,
    APP_COPYRIGHT,
    APP_DESCRIPTION,
    APP_LICENSE_ID,
    APP_LICENSE_NAME,
    APP_NAME,
    APP_USER_MODEL_ID,
    APP_VERSION,
    SPLASH_DURATION_MS,
)


ICON_RELATIVE_PATH = Path("assets") / "office2pdf.ico"
LICENSE_RELATIVE_PATH = Path("LICENSE.txt")
NOTICES_RELATIVE_PATH = Path("THIRD_PARTY_NOTICES.txt")


def resource_path(relative_path: Path | str) -> Path:
    """Return a source-tree or PyInstaller-safe path to a bundled resource."""
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return bundle_root / Path(relative_path)


def load_application_icon() -> QIcon:
    """Load the branded icon without making startup depend on the asset."""
    icon_path = resource_path(ICON_RELATIVE_PATH)
    return QIcon(str(icon_path)) if icon_path.is_file() else QIcon()


def load_text_resource(relative_path: Path, fallback: str) -> str:
    """Read a bundled text resource without making startup fragile."""
    try:
        return resource_path(relative_path).read_text(encoding="utf-8")
    except (OSError, UnicodeError):
        return fallback


def set_windows_app_identity() -> None:
    """Give Windows a stable taskbar identity so the packaged icon is used."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            APP_USER_MODEL_ID
        )
    except Exception:
        # Cosmetic integration must never prevent the converter from starting.
        pass


# ==========================================================================
# Theme
# ==========================================================================

INDUSTRIAL_DARK_QSS = """
QMainWindow, QWidget {
    background-color: #1c1e21;
    color: #d7dadd;
    font-family: "Segoe UI", "Consolas", sans-serif;
    font-size: 10pt;
}

QLabel#HeaderTitle {
    color: #f2a900;
    font-size: 16pt;
    font-weight: 700;
    letter-spacing: 2px;
}

QLabel#HeaderSubtitle {
    color: #7d8489;
    font-size: 8pt;
}

QGroupBox {
    border: 1px solid #34383d;
    border-radius: 3px;
    margin-top: 14px;
    padding-top: 10px;
    background-color: #202327;
    font-weight: 600;
    color: #b9bec3;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 6px;
    color: #f2a900;
    text-transform: uppercase;
    font-size: 8pt;
    letter-spacing: 1px;
}

QListWidget {
    background-color: #17181a;
    border: 1px solid #34383d;
    border-radius: 3px;
    color: #d7dadd;
    outline: none;
}

QListWidget::item {
    padding: 5px 6px;
    border-bottom: 1px solid #232629;
}

QListWidget::item:selected {
    background-color: #3a3020;
    color: #f2a900;
}

QListWidget::item:hover {
    background-color: #24272b;
}

QPlainTextEdit#LogConsole {
    background-color: #101113;
    color: #c7cbcf;
    border: 1px solid #34383d;
    border-radius: 3px;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
}

QPushButton {
    background-color: #2b2f34;
    border: 1px solid #40454b;
    border-radius: 3px;
    padding: 6px 14px;
    color: #d7dadd;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #34393f;
    border: 1px solid #f2a900;
}

QPushButton:pressed {
    background-color: #1c1e21;
}

QPushButton:disabled {
    color: #55595d;
    border: 1px solid #2c2f33;
    background-color: #202327;
}

QPushButton#PrimaryButton {
    background-color: #f2a900;
    color: #1c1e21;
    border: 1px solid #f2a900;
}

QPushButton#PrimaryButton:hover {
    background-color: #ffbb2e;
}

QPushButton#PrimaryButton:disabled {
    background-color: #4a4020;
    color: #7d7460;
    border: 1px solid #4a4020;
}

QPushButton#DangerButton {
    background-color: #3a2020;
    border: 1px solid #7a3030;
    color: #e07070;
}

QPushButton#DangerButton:hover {
    background-color: #4a2626;
    border: 1px solid #c04040;
}

QPushButton#DangerButton:disabled {
    background-color: #202327;
    border: 1px solid #2c2f33;
    color: #55595d;
}

QLineEdit, QSpinBox {
    background-color: #17181a;
    border: 1px solid #34383d;
    border-radius: 3px;
    padding: 4px 6px;
    color: #d7dadd;
    selection-background-color: #f2a900;
    selection-color: #1c1e21;
}

QLineEdit:disabled, QSpinBox:disabled {
    color: #55595d;
    background-color: #1a1c1e;
}

QLineEdit:focus, QSpinBox:focus {
    border: 1px solid #f2a900;
}

QCheckBox {
    spacing: 8px;
    padding: 2px 0;
}

QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #40454b;
    border-radius: 2px;
    background-color: #17181a;
}

QCheckBox::indicator:checked {
    background-color: #f2a900;
    border: 1px solid #f2a900;
}

QCheckBox:disabled {
    color: #55595d;
}

QProgressBar {
    border: 1px solid #34383d;
    border-radius: 3px;
    background-color: #17181a;
    text-align: center;
    color: #d7dadd;
    height: 18px;
}

QProgressBar::chunk {
    background-color: #f2a900;
    border-radius: 2px;
}

QSplitter::handle {
    background-color: #1c1e21;
    width: 4px;
}

QScrollBar:vertical {
    background: #17181a;
    width: 12px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #34383d;
    min-height: 24px;
    border-radius: 3px;
}
QScrollBar::handle:vertical:hover {
    background: #4a4f55;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QLabel#DropHint {
    color: #55595d;
    border: 2px dashed #34383d;
    border-radius: 4px;
}

QLabel#DropHint[dragActive="true"] {
    color: #f2a900;
    border: 2px dashed #f2a900;
}

QLabel#StatusLabel {
    color: #7d8489;
    font-size: 8pt;
}

QFrame#Divider {
    background-color: #34383d;
    max-height: 1px;
}

QTabWidget::pane {
    border: 1px solid #34383d;
    background-color: #1c1e21;
}

QTabBar::tab {
    background-color: #24272b;
    color: #9ca2a8;
    border: 1px solid #34383d;
    border-bottom: none;
    padding: 8px 18px;
    min-width: 90px;
}

QTabBar::tab:selected {
    background-color: #1c1e21;
    color: #f2a900;
    border-top: 2px solid #f2a900;
}

QTabBar::tab:hover:!selected {
    color: #d7dadd;
    background-color: #2b2f34;
}

QLabel#AboutTitle {
    color: #f2a900;
    font-size: 24pt;
    font-weight: 700;
}

QLabel#AboutVersion {
    color: #9ca2a8;
    font-size: 10pt;
}

QTextBrowser#AboutOverview {
    background-color: #17181a;
    color: #d7dadd;
    border: 1px solid #34383d;
    padding: 12px;
}
"""


def _mime_local_paths(mime: QMimeData) -> list[str]:
    """Extract local filesystem paths from dropped mime data."""
    paths: list[str] = []
    if mime.hasUrls():
        for url in mime.urls():
            if url.isLocalFile():
                paths.append(url.toLocalFile())
    return paths


# ==========================================================================
# Startup splash
# ==========================================================================

class StartupSplash(QWidget):
    """Five-second branded startup card modelled on the supplied artwork."""

    def __init__(self, icon: QIcon) -> None:
        flags = (
            Qt.WindowType.SplashScreen
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        super().__init__(None, flags)
        self.setObjectName("StartupSplash")
        self.setFixedSize(620, 620)
        self._started_at = time.monotonic()

        self.setStyleSheet(
            """
            QWidget#StartupSplash {
                background-color: #07172a;
                border: 1px solid #17314d;
            }
            QLabel {
                background: transparent;
                color: #dfe8f1;
                font-family: "Segoe UI";
            }
            QLabel#SplashTitle {
                font-size: 34pt;
                font-weight: 700;
            }
            QLabel#SplashSubtitle {
                color: #b7c3ce;
                font-size: 12pt;
            }
            QLabel#SplashVersion {
                color: #8da0b2;
                font-size: 9pt;
            }
            QProgressBar {
                border: 1px solid #24445f;
                border-radius: 2px;
                background-color: #0b2035;
                height: 5px;
                text-align: center;
                color: transparent;
            }
            QProgressBar::chunk {
                background-color: #f2a900;
                border-radius: 2px;
            }
            """
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(72, 48, 72, 42)
        layout.setSpacing(10)

        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if not icon.isNull():
            icon_label.setPixmap(icon.pixmap(150, 150))
        layout.addWidget(icon_label)

        title = QLabel(
            '<span style="color:#edf3f8;">Office</span>'
            '<span style="color:#f2a900;">2PDF</span>'
        )
        title.setObjectName("SplashTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setStyleSheet("background:#24445f; max-height:1px;")
        layout.addWidget(divider)

        subtitle = QLabel("Office file conversion made simple")
        subtitle.setObjectName("SplashSubtitle")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)

        version = QLabel(f"Version {APP_VERSION}")
        version.setObjectName("SplashVersion")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version)
        layout.addStretch(1)

        loading = QLabel("Loading components…")
        loading.setObjectName("SplashVersion")
        loading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(loading)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        layout.addWidget(self.progress)

        credit = QLabel("A PROJECT BY LEON")
        credit.setObjectName("SplashVersion")
        credit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(credit)

        self._timer = QTimer(self)
        self._timer.setInterval(50)
        self._timer.timeout.connect(self._advance_progress)

    def showEvent(self, event) -> None:  # noqa: N802 - Qt API name
        super().showEvent(event)
        screen = QApplication.primaryScreen()
        if screen is not None:
            frame = self.frameGeometry()
            frame.moveCenter(screen.availableGeometry().center())
            self.move(frame.topLeft())
        self._started_at = time.monotonic()
        self._timer.start()

    def _advance_progress(self) -> None:
        elapsed_ms = (time.monotonic() - self._started_at) * 1000.0
        self.progress.setValue(min(99, int((elapsed_ms / SPLASH_DURATION_MS) * 100)))

    def reveal(self, window: QMainWindow) -> None:
        self._timer.stop()
        self.progress.setValue(100)
        self.close()
        window.show()
        window.raise_()
        window.activateWindow()


# ==========================================================================
# Drop-enabled queue widgets
# ==========================================================================

class DropHintLabel(QLabel):
    """Empty-state placeholder shown when the queue has nothing in it yet."""

    filesDropped = pyqtSignal(list)

    def __init__(self) -> None:
        super().__init__(
            "DRAG & DROP\nfiles or folders here\n\n— or use Add Files / Add Folder below —"
        )
        self.setObjectName("DropHint")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        font = self.font()
        font.setPointSize(10)
        font.setBold(True)
        self.setFont(font)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if _mime_local_paths(event.mimeData()):
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event: QDropEvent) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        paths = _mime_local_paths(event.mimeData())
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()


class DropListWidget(QListWidget):
    """The populated queue view -- still accepts further drops."""

    filesDropped = pyqtSignal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setAlternatingRowColors(False)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if _mime_local_paths(event.mimeData()):
            event.acceptProposedAction()

    def dragMoveEvent(self, event) -> None:
        event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = _mime_local_paths(event.mimeData())
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()


class FileQueue(QWidget):
    """Combines the empty-state hint and the populated list into one panel."""

    countChanged = pyqtSignal(int)

    def __init__(self) -> None:
        super().__init__()
        self._paths: list[str] = []

        self.stack = QStackedWidget()
        self.hint = DropHintLabel()
        self.list_widget = DropListWidget()
        self.stack.addWidget(self.hint)         # index 0: empty
        self.stack.addWidget(self.list_widget)  # index 1: populated

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stack)

        self.hint.filesDropped.connect(self.add_paths)
        self.list_widget.filesDropped.connect(self.add_paths)

    # -- public API ---------------------------------------------------

    def add_paths(self, raw_paths: list[str]) -> None:
        style = self.style()
        added = 0
        for raw in raw_paths:
            p = Path(raw)
            resolved = str(p.resolve())
            if resolved in self._paths:
                continue
            if not p.exists():
                continue
            self._paths.append(resolved)
            item = QListWidgetItem(resolved)
            if p.is_dir():
                item.setIcon(style.standardIcon(QStyle.StandardPixmap.SP_DirIcon))
            else:
                item.setIcon(style.standardIcon(QStyle.StandardPixmap.SP_FileIcon))
            item.setToolTip(resolved)
            self.list_widget.addItem(item)
            added += 1
        if added:
            self._sync_stack()
            self.countChanged.emit(len(self._paths))

    def remove_selected(self) -> None:
        for item in self.list_widget.selectedItems():
            path = item.text()
            if path in self._paths:
                self._paths.remove(path)
            self.list_widget.takeItem(self.list_widget.row(item))
        self._sync_stack()
        self.countChanged.emit(len(self._paths))

    def clear(self) -> None:
        self._paths.clear()
        self.list_widget.clear()
        self._sync_stack()
        self.countChanged.emit(0)

    def paths(self) -> list[str]:
        return list(self._paths)

    def _sync_stack(self) -> None:
        self.stack.setCurrentIndex(1 if self._paths else 0)


# ==========================================================================
# Conversion worker
# ==========================================================================

@dataclass
class ConversionOptions:
    output_dir: str | None
    recursive: bool
    overwrite: bool
    jobs: int
    timeout: int
    retries: int
    prefer_native: bool = False


class ConversionWorker(QThread):
    log = pyqtSignal(str, str)                 # message, level
    progress = pyqtSignal(int, int)            # completed, total
    finishedRun = pyqtSignal(int, int, int, int)  # ok, fail, skip, cancelled
    outputProduced = pyqtSignal(str)            # output folder
    fatalError = pyqtSignal(str)

    def __init__(
        self,
        queued_paths: list[str],
        options: ConversionOptions,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.queued_paths = queued_paths
        self.options = options
        self._cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self._cancel_event.set()

    def run(self) -> None:
        try:
            self._run_impl()
        except Exception as exc:  # final containment boundary for the GUI thread
            self.fatalError.emit(
                f"Unexpected conversion worker failure: {type(exc).__name__}: {exc}"
            )

    def _run_impl(self) -> None:
        soffice_bin: str | None
        try:
            soffice_bin = backend.find_soffice()
            self.log.emit(f"Using LibreOffice binary: {soffice_bin}", "info")
        except backend.LibreOfficeNotFoundError as exc:
            soffice_bin = None
            # Not automatically fatal: a native Office backend might still
            # handle every queued file on Windows. convert_one() reports a
            # clear per-file failure for anything that actually needed
            # LibreOffice and didn't have it.
            self.log.emit(str(exc), "warn")
            if not (sys.platform.startswith("win") and self.options.prefer_native):
                self.fatalError.emit(str(exc))
                return

        inputs = backend.discover_inputs(self.queued_paths, self.options.recursive)
        if not inputs:
            self.fatalError.emit(
                "No convertible files found among the queued items "
                f"(supported: {', '.join(sorted(backend.SUPPORTED_EXTENSIONS))})."
            )
            return

        collisions = backend.find_destination_collisions(inputs, self.options.output_dir)
        if collisions:
            lines = [
                "Refusing to run: output filename collisions detected.",
                "Two different source files would overwrite the same PDF:",
            ]
            for dest, srcs in collisions.items():
                lines.append(f"  {dest}  <-  {', '.join(str(s) for s in srcs)}")
            lines.append(
                "Convert the conflicting files separately, choose different output "
                "folders, or rename one of the source files."
            )
            self.fatalError.emit("\n".join(lines))
            return

        total = len(inputs)
        self.log.emit(f"Found {total} file(s) to convert.", "info")
        self.progress.emit(0, total)

        def run_one(src: Path) -> backend.ConversionResult:
            out_dir = (
                Path(self.options.output_dir)
                if self.options.output_dir
                else src.resolve().parent
            )
            return backend.convert_one(
                src=src,
                out_dir=out_dir,
                soffice_bin=soffice_bin,
                timeout=self.options.timeout,
                retries=self.options.retries,
                overwrite=self.options.overwrite,
                prefer_native=self.options.prefer_native,
            )

        completed = 0
        ok_count = fail_count = skip_count = cancelled_count = 0
        cancel_logged = False

        if self.options.jobs > 1:
            pending = deque(inputs)
            futures: dict[concurrent.futures.Future, Path] = {}

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=self.options.jobs
            ) as pool:
                while pending and len(futures) < self.options.jobs and not self._cancel_event.is_set():
                    src = pending.popleft()
                    futures[pool.submit(run_one, src)] = src

                while futures:
                    finished, _ = concurrent.futures.wait(
                        futures,
                        return_when=concurrent.futures.FIRST_COMPLETED,
                    )
                    for future in finished:
                        src = futures.pop(future)
                        try:
                            result = future.result()
                        except Exception as exc:  # defensive: backend should contain these
                            result = backend.ConversionResult(
                                source=src,
                                output=None,
                                status=backend.ConversionStatus.FAILED,
                                attempts=0,
                                message=(
                                    "unexpected worker error: "
                                    f"{type(exc).__name__}: {exc}"
                                ),
                                seconds=0.0,
                            )

                        completed += 1
                        ok_count, fail_count, skip_count = self._handle_result(
                            result, ok_count, fail_count, skip_count
                        )
                        self.progress.emit(completed, total)

                    if self._cancel_event.is_set() and not cancel_logged:
                        cancel_logged = True
                        self.log.emit(
                            "Cancel requested — active conversions will finish safely; "
                            "queued files will not start.",
                            "warn",
                        )

                    while (
                        pending
                        and len(futures) < self.options.jobs
                        and not self._cancel_event.is_set()
                    ):
                        src = pending.popleft()
                        futures[pool.submit(run_one, src)] = src

            if self._cancel_event.is_set():
                cancelled_count = len(pending)
        else:
            for index, src in enumerate(inputs):
                if self._cancel_event.is_set():
                    cancelled_count = total - index
                    self.log.emit(
                        "Cancel requested — remaining queued files were not started.",
                        "warn",
                    )
                    break

                result = run_one(src)
                completed += 1
                ok_count, fail_count, skip_count = self._handle_result(
                    result, ok_count, fail_count, skip_count
                )
                self.progress.emit(completed, total)

        self.finishedRun.emit(ok_count, fail_count, skip_count, cancelled_count)

    def _handle_result(self, result, ok_count, fail_count, skip_count):
        if result.status is backend.ConversionStatus.SKIPPED:
            skip_count += 1
            self.log.emit(f"SKIP  {result.source.name} — {result.message}", "skip")
            if result.output:
                self.outputProduced.emit(str(result.output.parent))
        elif result.status is backend.ConversionStatus.SUCCESS:
            ok_count += 1
            self.log.emit(
                f"OK    [{result.backend}] {result.source.name}  →  {result.output.name}  "
                f"({result.seconds:.1f}s, attempt {result.attempts})",
                "ok",
            )
            if result.output:
                self.outputProduced.emit(str(result.output.parent))
        else:
            fail_count += 1
            self.log.emit(f"FAIL  {result.source.name} — {result.message}", "fail")
        return ok_count, fail_count, skip_count


# ==========================================================================
# Main window
# ==========================================================================

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(1040, 680)
        self.worker: ConversionWorker | None = None
        self.last_output_dir: str | None = None
        self._output_folders: set[str] = set()
        self._close_when_worker_finishes = False

        self._build_ui()

    # -- UI construction ------------------------------------------------

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)

        self.main_tabs = QTabWidget()
        self.main_tabs.setDocumentMode(True)
        outer.addWidget(self.main_tabs)

        converter_page = QWidget()
        self.main_tabs.addTab(converter_page, "Converter")
        root = QVBoxLayout(converter_page)
        root.setContentsMargins(16, 14, 16, 14)
        root.setSpacing(10)

        # Header ----------------------------------------------------
        header = QVBoxLayout()
        header.setSpacing(0)
        title = QLabel("OFFICE  →  PDF")
        title.setObjectName("HeaderTitle")
        subtitle = QLabel("Reliable DOCX / XLSX / PPTX conversion — drag & drop enabled")
        subtitle.setObjectName("HeaderSubtitle")
        header.addWidget(title)
        header.addWidget(subtitle)
        root.addLayout(header)

        divider = QFrame()
        divider.setObjectName("Divider")
        divider.setFrameShape(QFrame.Shape.HLine)
        root.addWidget(divider)

        # Main split: queue (left) / options (right) -----------------
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        # Queue panel
        queue_box = QGroupBox("Conversion Queue")
        queue_layout = QVBoxLayout(queue_box)

        queue_toolbar = QHBoxLayout()
        self.btn_add_files = QPushButton("Add Files…")
        self.btn_add_folder = QPushButton("Add Folder…")
        self.btn_remove = QPushButton("Remove Selected")
        self.btn_clear = QPushButton("Clear All")
        self.btn_clear.setObjectName("DangerButton")
        for b in (self.btn_add_files, self.btn_add_folder, self.btn_remove, self.btn_clear):
            queue_toolbar.addWidget(b)
        queue_toolbar.addStretch(1)
        queue_layout.addLayout(queue_toolbar)

        self.queue = FileQueue()
        queue_layout.addWidget(self.queue, stretch=1)

        self.queue_count_label = QLabel("0 items queued")
        self.queue_count_label.setObjectName("StatusLabel")
        queue_layout.addWidget(self.queue_count_label)

        splitter.addWidget(queue_box)

        # Options panel
        options_box = QGroupBox("Options")
        options_layout = QVBoxLayout(options_box)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setSpacing(8)

        self.chk_same_folder = QCheckBox("Output next to each source file")
        self.chk_same_folder.setChecked(True)
        out_row = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("Choose an output folder…")
        self.output_dir_edit.setEnabled(False)
        self.btn_browse_out = QPushButton("Browse…")
        self.btn_browse_out.setEnabled(False)
        out_row.addWidget(self.output_dir_edit)
        out_row.addWidget(self.btn_browse_out)
        out_row_widget = QWidget()
        out_row_widget.setLayout(out_row)

        form.addRow(self.chk_same_folder)
        form.addRow("Output folder", out_row_widget)

        self.chk_recursive = QCheckBox("Recurse into subfolders")
        self.chk_overwrite = QCheckBox("Overwrite existing PDFs")
        self.chk_prefer_native = QCheckBox("Use Microsoft Office for best layout fidelity (experimental)")
        self.chk_prefer_native.setChecked(False)
        self.chk_prefer_native.setToolTip(
            "Windows only. Runs the matching Microsoft Office application in an "
            "isolated helper process for output closest to 'Save as PDF'. This "
            "remains opt-in until it has been verified on this PC. LibreOffice "
            "is used automatically if native export fails or is unavailable."
        )
        form.addRow(self.chk_recursive)
        form.addRow(self.chk_overwrite)
        form.addRow(self.chk_prefer_native)

        self.spin_jobs = QSpinBox()
        self.spin_jobs.setRange(1, 16)
        self.spin_jobs.setValue(1)
        form.addRow("Parallel jobs", self.spin_jobs)

        self.spin_timeout = QSpinBox()
        self.spin_timeout.setRange(10, 3600)
        self.spin_timeout.setSingleStep(10)
        self.spin_timeout.setValue(backend.DEFAULT_TIMEOUT)
        self.spin_timeout.setSuffix(" s")
        form.addRow("Timeout per file", self.spin_timeout)

        self.spin_retries = QSpinBox()
        self.spin_retries.setRange(0, 10)
        self.spin_retries.setValue(backend.DEFAULT_RETRIES)
        form.addRow("Retries on failure", self.spin_retries)

        options_layout.addLayout(form)
        options_layout.addStretch(1)

        supported = QLabel("Supported: " + ", ".join(sorted(backend.SUPPORTED_EXTENSIONS)))
        supported.setObjectName("StatusLabel")
        supported.setWordWrap(True)
        options_layout.addWidget(supported)

        splitter.addWidget(options_box)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        root.addWidget(splitter, stretch=1)

        # Progress row -------------------------------------------------
        progress_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Idle")
        progress_row.addWidget(self.progress_bar, stretch=1)
        root.addLayout(progress_row)

        # Log console ----------------------------------------------------
        log_box = QGroupBox("Log")
        log_layout = QVBoxLayout(log_box)
        self.log_console = QPlainTextEdit()
        self.log_console.setObjectName("LogConsole")
        self.log_console.setReadOnly(True)
        self.log_console.setMinimumHeight(160)
        log_layout.addWidget(self.log_console)
        root.addWidget(log_box)

        # Action row -------------------------------------------------
        action_row = QHBoxLayout()
        self.btn_open_output = QPushButton("Open Output Folder")
        self.btn_open_output.setEnabled(False)
        self.btn_clear_log = QPushButton("Clear Log")
        action_row.addWidget(self.btn_open_output)
        action_row.addWidget(self.btn_clear_log)
        action_row.addStretch(1)
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setObjectName("DangerButton")
        self.btn_cancel.setEnabled(False)
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.setObjectName("PrimaryButton")
        action_row.addWidget(self.btn_cancel)
        action_row.addWidget(self.btn_convert)
        root.addLayout(action_row)

        # -- wire up ------------------------------------------------
        self.btn_add_files.clicked.connect(self._on_add_files)
        self.btn_add_folder.clicked.connect(self._on_add_folder)
        self.btn_remove.clicked.connect(self.queue.remove_selected)
        self.btn_clear.clicked.connect(self.queue.clear)
        self.queue.countChanged.connect(self._on_queue_count_changed)

        self.chk_same_folder.toggled.connect(self._on_same_folder_toggled)
        self.btn_browse_out.clicked.connect(self._on_browse_output)

        self.btn_convert.clicked.connect(self._on_convert_clicked)
        self.btn_cancel.clicked.connect(self._on_cancel_clicked)
        self.btn_clear_log.clicked.connect(self.log_console.clear)
        self.btn_open_output.clicked.connect(self._on_open_output)

        self.main_tabs.addTab(self._build_about_page(), "About")

    def _build_about_page(self) -> QWidget:
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(14)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(18)
        icon_label = QLabel()
        icon_label.setFixedSize(112, 112)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon = load_application_icon()
        if not icon.isNull():
            icon_label.setPixmap(icon.pixmap(104, 104))
        brand_row.addWidget(icon_label)

        brand_text = QVBoxLayout()
        brand_text.setSpacing(3)
        title = QLabel(APP_NAME)
        title.setObjectName("AboutTitle")
        version = QLabel(f"Version {APP_VERSION}  •  Built {APP_BUILD_DATE}")
        version.setObjectName("AboutVersion")
        description = QLabel(APP_DESCRIPTION)
        description.setWordWrap(True)
        brand_text.addWidget(title)
        brand_text.addWidget(version)
        brand_text.addWidget(description)
        brand_text.addStretch(1)
        brand_row.addLayout(brand_text, stretch=1)
        root.addLayout(brand_row)

        details = QGroupBox("Program Information")
        details_form = QFormLayout(details)
        details_form.addRow("Author", QLabel(APP_AUTHOR))
        details_form.addRow("Copyright", QLabel(APP_COPYRIGHT))
        details_form.addRow("Licence", QLabel(f"{APP_LICENSE_NAME} ({APP_LICENSE_ID})"))
        details_form.addRow("Default backend", QLabel("LibreOffice (separately installed)"))
        details_form.addRow("Optional backend", QLabel("Microsoft Word / Excel / PowerPoint on Windows"))
        details_form.addRow("Source code", QLabel(f"Office2PDF-{APP_VERSION}-source.zip is supplied with the release"))
        root.addWidget(details)

        info_tabs = QTabWidget()

        overview = QTextBrowser()
        overview.setObjectName("AboutOverview")
        overview.setOpenExternalLinks(False)
        overview.setHtml(
            f"""
            <h3 style='color:#f2a900;'>Reliable conversion without silent failure</h3>
            <p>Office2PDF converts Word, Excel and PowerPoint documents to PDF using
            LibreOffice by default. On Windows, native Microsoft Office export can be
            enabled for maximum layout fidelity.</p>
            <p>Every conversion uses a private staging directory. A new PDF is checked
            before it atomically replaces the destination, preventing stale-output false
            success and protecting an existing valid PDF when conversion fails.</p>
            <p><b>Supported formats:</b> {', '.join(sorted(backend.SUPPORTED_EXTENSIONS))}</p>
            <p><b>Release:</b> {APP_VERSION}<br/>
            <b>Build date:</b> {APP_BUILD_DATE}<br/>
            <b>Licence identifier:</b> {APP_LICENSE_ID}</p>
            """
        )
        info_tabs.addTab(overview, "Overview")

        licence = QPlainTextEdit()
        licence.setReadOnly(True)
        licence.setPlainText(
            load_text_resource(
                LICENSE_RELATIVE_PATH,
                "GNU GPL v3 licence text was not found in this installation.",
            )
        )
        info_tabs.addTab(licence, "GPL Licence")

        notices = QPlainTextEdit()
        notices.setReadOnly(True)
        notices.setPlainText(
            load_text_resource(
                NOTICES_RELATIVE_PATH,
                "Third-party notices were not found in this installation.",
            )
        )
        info_tabs.addTab(notices, "Third-Party Notices")

        root.addWidget(info_tabs, stretch=1)
        return page

    # -- queue handlers ---------------------------------------------

    def _on_add_files(self) -> None:
        exts = " ".join(f"*{e}" for e in sorted(backend.SUPPORTED_EXTENSIONS))
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add files", "", f"Office documents ({exts});;All files (*)"
        )
        if files:
            self.queue.add_paths(files)

    def _on_add_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Add folder")
        if folder:
            self.queue.add_paths([folder])

    def _on_queue_count_changed(self, count: int) -> None:
        noun = "item" if count == 1 else "items"
        self.queue_count_label.setText(f"{count} {noun} queued")

    # -- options handlers ---------------------------------------------

    def _on_same_folder_toggled(self, checked: bool) -> None:
        self.output_dir_edit.setEnabled(not checked)
        self.btn_browse_out.setEnabled(not checked)

    def _on_browse_output(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder")
        if folder:
            self.output_dir_edit.setText(folder)

    # -- conversion lifecycle ---------------------------------------------

    def _on_convert_clicked(self) -> None:
        paths = self.queue.paths()
        if not paths:
            QMessageBox.warning(self, "Nothing queued", "Add files or folders to the queue first.")
            return

        output_dir = None
        if not self.chk_same_folder.isChecked():
            output_dir = self.output_dir_edit.text().strip()
            if not output_dir:
                QMessageBox.warning(
                    self, "Output folder required",
                    "Choose an output folder, or check 'Output next to each source file'.",
                )
                return

        options = ConversionOptions(
            output_dir=output_dir,
            recursive=self.chk_recursive.isChecked(),
            overwrite=self.chk_overwrite.isChecked(),
            jobs=self.spin_jobs.value(),
            timeout=self.spin_timeout.value(),
            retries=self.spin_retries.value(),
            prefer_native=self.chk_prefer_native.isChecked(),
        )
        self.last_output_dir = output_dir
        self._output_folders.clear()
        self.btn_open_output.setEnabled(False)

        self.log_console.clear()
        self._append_log(f"Starting conversion of {len(paths)} queued item(s)…", "info")
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Scanning…")

        self._set_running_state(True)

        self.worker = ConversionWorker(paths, options, self)
        self.worker.log.connect(self._append_log)
        self.worker.progress.connect(self._on_progress)
        self.worker.finishedRun.connect(self._on_finished)
        self.worker.outputProduced.connect(self._on_output_produced)
        self.worker.fatalError.connect(self._on_fatal_error)
        self.worker.finished.connect(self._on_worker_thread_finished)
        self.worker.start()

    def _on_cancel_clicked(self) -> None:
        if self.worker:
            self.worker.request_cancel()
            self.btn_cancel.setEnabled(False)

    def _set_running_state(self, running: bool) -> None:
        self.btn_convert.setEnabled(not running)
        self.btn_cancel.setEnabled(running)
        for w in (
            self.btn_add_files, self.btn_add_folder, self.btn_remove, self.btn_clear,
            self.chk_same_folder, self.chk_recursive, self.chk_overwrite, self.chk_prefer_native,
            self.spin_jobs, self.spin_timeout, self.spin_retries,
        ):
            w.setEnabled(not running)
        if running:
            self.btn_browse_out.setEnabled(False)
            self.output_dir_edit.setEnabled(False)
        else:
            self._on_same_folder_toggled(self.chk_same_folder.isChecked())

    def _on_progress(self, done: int, total: int) -> None:
        pct = int((done / total) * 100) if total else 0
        self.progress_bar.setValue(pct)
        self.progress_bar.setFormat(f"{done} / {total}  ({pct}%)")

    def _on_finished(self, ok: int, fail: int, skip: int, cancelled: int) -> None:
        if cancelled:
            self.progress_bar.setFormat(
                f"Cancelled — {ok} ok, {skip} skipped, {fail} failed, "
                f"{cancelled} not started"
            )
            level = "warn" if not fail else "fail"
            self._append_log(
                f"Cancelled: {ok} succeeded, {skip} skipped, {fail} failed, "
                f"{cancelled} not started.",
                level,
            )
        else:
            self.progress_bar.setFormat(
                f"Done — {ok} ok, {skip} skipped, {fail} failed"
            )
            level = "fail" if fail else "ok"
            self._append_log(
                f"Finished: {ok} succeeded, {skip} skipped, {fail} failed.",
                level,
            )
        self.btn_open_output.setEnabled(bool(self._output_folders))

    def _on_output_produced(self, folder: str) -> None:
        self._output_folders.add(str(Path(folder).resolve()))

    def _on_fatal_error(self, message: str) -> None:
        self.progress_bar.setFormat("Failed")
        self._append_log(message, "fail")
        if not self._close_when_worker_finishes:
            QMessageBox.critical(self, "Conversion could not start", message)

    def _on_worker_thread_finished(self) -> None:
        self._set_running_state(False)
        worker = self.worker
        self.worker = None
        if worker is not None:
            worker.deleteLater()
        if self._close_when_worker_finishes:
            self._close_when_worker_finishes = False
            self.close()

    def _on_open_output(self) -> None:
        if self._output_folders:
            folder = sorted(self._output_folders, key=str.casefold)[0]
        else:
            target = self.last_output_dir or (
                self.queue.paths()[0] if self.queue.paths() else None
            )
            if not target:
                return
            folder = (
                target
                if Path(target).is_dir()
                else str(Path(target).resolve().parent)
            )
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", folder], check=False)
            else:
                subprocess.run(["xdg-open", folder], check=False)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Couldn't open folder", str(exc))

    # -- log formatting ---------------------------------------------

    def _append_log(self, message: str, level: str = "info") -> None:
        colors = {
            "info": "#7d8489",
            "ok": "#5fbf6a",
            "fail": "#e0605f",
            "skip": "#d9a441",
            "warn": "#d9a441",
        }
        color = colors.get(level, "#c7cbcf")
        safe = message.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        safe = safe.replace("\n", "<br/>")
        self.log_console.appendHtml(f'<span style="color:{color};">{safe}</span>')

    # -- window lifecycle ---------------------------------------------

    def closeEvent(self, event) -> None:
        if self.worker and self.worker.isRunning():
            if not self._close_when_worker_finishes:
                self._close_when_worker_finishes = True
                self.worker.request_cancel()
                self.btn_cancel.setEnabled(False)
                self.progress_bar.setFormat("Closing after active conversion finishes safely…")
                self._append_log(
                    "Close requested — queued work was cancelled; waiting for active "
                    "conversion(s) to finish safely.",
                    "warn",
                )
            event.ignore()
            return
        event.accept()


def main() -> int:
    set_windows_app_identity()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_NAME)

    icon = load_application_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)

    app.setApplicationVersion(APP_VERSION)
    app.setStyleSheet(INDUSTRIAL_DARK_QSS)

    splash = StartupSplash(icon)
    splash.show()
    app.processEvents()

    window = MainWindow()
    if not icon.isNull():
        window.setWindowIcon(icon)

    QTimer.singleShot(SPLASH_DURATION_MS, lambda: splash.reveal(window))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
