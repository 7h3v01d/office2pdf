#!/usr/bin/env python3
# Copyright (C) 2026 Leon Priest
# SPDX-License-Identifier: GPL-3.0-or-later

"""Reliable Office document to PDF conversion with safe backend isolation.

The backend is deliberately conservative:

* every conversion uses an isolated LibreOffice profile;
* each attempt writes to a private staging directory;
* staged PDFs are validated before they atomically replace the destination;
* subprocess timeouts terminate the LibreOffice process tree;
* expected operating-system failures become structured results, not crashes;
* every input receives an explicit success, skipped, failed, or cancelled state.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import dataclasses
import enum
import functools
import logging
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Iterable

from version_info import APP_VERSION

LOG = logging.getLogger("office2pdf")

SUPPORTED_EXTENSIONS = frozenset(
    {
        ".docx", ".doc", ".dotx", ".dot", ".odt", ".rtf",
        ".xlsx", ".xls", ".xltx", ".xlt", ".ods", ".csv",
        ".pptx", ".ppt", ".odp",
    }
)

# Extensions each native Office application can export directly, used for
# backend selection when Word/Excel/PowerPoint are actually installed (see
# "native Microsoft Office backend" section below).
WORD_EXTENSIONS = frozenset({".doc", ".docx", ".dot", ".dotx", ".rtf", ".odt"})
EXCEL_EXTENSIONS = frozenset({".xls", ".xlsx", ".xlt", ".xltx", ".ods", ".csv"})
POWERPOINT_EXTENSIONS = frozenset({".ppt", ".pptx", ".odp"})

DEFAULT_TIMEOUT = 120
DEFAULT_RETRIES = 2
MIN_PDF_SIZE = 128


class Office2PDFError(RuntimeError):
    """Base exception for startup/configuration errors."""


class LibreOfficeNotFoundError(Office2PDFError):
    """Raised when no usable LibreOffice executable can be located."""


class ConversionStatus(str, enum.Enum):
    SUCCESS = "success"
    SKIPPED = "skipped"
    FAILED = "failed"
    CANCELLED = "cancelled"


class Backend(str, enum.Enum):
    LIBREOFFICE = "libreoffice"
    MS_WORD = "word"
    MS_EXCEL = "excel"
    MS_POWERPOINT = "powerpoint"


@dataclasses.dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path | None
    status: ConversionStatus
    attempts: int
    message: str
    seconds: float
    backend: str = Backend.LIBREOFFICE.value

    @property
    def ok(self) -> bool:
        """Compatibility property: skipped outputs count as non-failures."""
        return self.status in {ConversionStatus.SUCCESS, ConversionStatus.SKIPPED}

    @property
    def skipped(self) -> bool:
        return self.status is ConversionStatus.SKIPPED

    @property
    def cancelled(self) -> bool:
        return self.status is ConversionStatus.CANCELLED


def find_soffice() -> str:
    """Locate LibreOffice without terminating the importing process."""
    for name in ("soffice", "libreoffice", "soffice.bin"):
        path = shutil.which(name)
        if path:
            return path

    candidates: list[Path] = []
    for env_name in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        root = os.environ.get(env_name)
        if root:
            candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")

    # Keep literal fallbacks for environments where the variables are absent.
    candidates.extend(
        [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    raise LibreOfficeNotFoundError(
        "LibreOffice ('soffice') was not found. Install LibreOffice, then retry. "
        "The Windows installer normally places it under Program Files\\LibreOffice."
    )


# --------------------------------------------------------------------------
# native Microsoft Office backend (Windows only, isolated worker process)
# --------------------------------------------------------------------------
#
# Native Office export gives the closest match to Word/Excel/PowerPoint's own
# "Save as PDF" output.  COM automation is deliberately kept in a separate
# helper process.  A Python thread cannot safely enforce a hard COM timeout,
# and image-wide taskkill commands could destroy an unrelated unsaved Office
# document.  The helper records the PID of the dedicated DispatchEx instance;
# timeout cleanup verifies that exact process image before terminating it.

@functools.lru_cache(maxsize=1)
def _native_office_available() -> dict[str, bool]:
    """Best-effort detection of installed Word/Excel/PowerPoint."""
    if not sys.platform.startswith("win"):
        return {}

    # Source runs require pywin32.  Frozen builds use the separately packaged
    # worker, which contains pywin32 even though the GUI executable may not.
    if not getattr(sys, "frozen", False):
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            return {}

    import winreg

    availability: dict[str, bool] = {}
    for prog_id, key in (
        ("Word.Application", "word"),
        ("Excel.Application", "excel"),
        ("PowerPoint.Application", "powerpoint"),
    ):
        try:
            winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, prog_id)
            availability[key] = True
        except OSError:
            availability[key] = False
    return availability


def choose_backend(src: Path, prefer_native: bool = False) -> Backend:
    """Pick the preferred converter for ``src``.

    LibreOffice remains the safe default until the isolated native worker has
    been exercised on the target Windows + Office installation.
    """
    ext = src.suffix.casefold()
    if prefer_native and sys.platform.startswith("win"):
        available = _native_office_available()
        if ext in WORD_EXTENSIONS and available.get("word"):
            return Backend.MS_WORD
        if ext in EXCEL_EXTENSIONS and available.get("excel"):
            return Backend.MS_EXCEL
        if ext in POWERPOINT_EXTENSIONS and available.get("powerpoint"):
            return Backend.MS_POWERPOINT
    return Backend.LIBREOFFICE


_EXPECTED_OFFICE_IMAGES = {
    Backend.MS_WORD: "WINWORD.EXE",
    Backend.MS_EXCEL: "EXCEL.EXE",
    Backend.MS_POWERPOINT: "POWERPNT.EXE",
}


def _native_worker_command(
    backend_kind: Backend, src: Path, dest_pdf: Path, pid_file: Path
) -> list[str]:
    """Build the command for the isolated native Office helper."""
    if getattr(sys, "frozen", False):
        worker = Path(sys.executable).with_name("Office2PDFNativeWorker.exe")
        if not worker.is_file():
            raise Office2PDFError(
                "native Office worker is missing from the application folder"
            )
        prefix = [str(worker)]
    else:
        worker = Path(__file__).resolve().with_name("native_office_worker.py")
        if not worker.is_file():
            raise Office2PDFError(f"native Office worker was not found: {worker}")
        prefix = [sys.executable, str(worker)]

    return prefix + [
        backend_kind.value,
        str(src.resolve()),
        str(dest_pdf.resolve()),
        str(pid_file.resolve()),
    ]


def _read_recorded_pid(pid_file: Path) -> int | None:
    try:
        value = int(pid_file.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        return None
    return value if value > 0 else None


def _terminate_recorded_office_process(backend_kind: Backend, pid: int | None) -> bool:
    """Terminate only the dedicated Office PID using a verified Win32 handle.

    ``taskkill`` is intentionally not used here.  Starting that utility can block
    for several seconds on Windows, which would turn a one-second conversion
    deadline into a much longer wait.  The image is queried and terminated via
    the same process handle, avoiding both the delay and the verify-then-kill PID
    reuse race.
    """
    if not pid or not sys.platform.startswith("win"):
        return False

    expected = _EXPECTED_OFFICE_IMAGES.get(backend_kind)
    if not expected:
        return False

    try:
        import ctypes
        from ctypes import wintypes

        process_terminate = 0x0001
        process_query_limited_information = 0x1000
        synchronize = 0x00100000
        wait_timeout_ms = 500

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        open_process = kernel32.OpenProcess
        open_process.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        open_process.restype = wintypes.HANDLE

        query_image = kernel32.QueryFullProcessImageNameW
        query_image.argtypes = [
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.LPWSTR,
            ctypes.POINTER(wintypes.DWORD),
        ]
        query_image.restype = wintypes.BOOL

        terminate_process = kernel32.TerminateProcess
        terminate_process.argtypes = [wintypes.HANDLE, wintypes.UINT]
        terminate_process.restype = wintypes.BOOL

        wait_for_single_object = kernel32.WaitForSingleObject
        wait_for_single_object.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        wait_for_single_object.restype = wintypes.DWORD

        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        close_handle.restype = wintypes.BOOL

        access = process_terminate | process_query_limited_information | synchronize
        handle = open_process(access, False, pid)
        if not handle:
            return False
        try:
            buffer = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buffer))
            if not query_image(handle, 0, buffer, ctypes.byref(size)):
                actual = None
            else:
                actual = Path(buffer.value).name.upper()

            if actual != expected:
                LOG.error(
                    "Refusing native timeout cleanup for PID %s: expected %s, observed %s",
                    pid,
                    expected,
                    actual or "(unverifiable)",
                )
                return False

            if not terminate_process(handle, 1):
                return False
            # Bounded best-effort reap; never extend the public deadline by
            # multiple seconds merely to wait for process bookkeeping.
            wait_for_single_object(handle, wait_timeout_ms)
            return True
        finally:
            close_handle(handle)
    except Exception:
        return False


def _terminate_native_worker(process: subprocess.Popen[bytes]) -> None:
    """Stop the isolated Python/EXE helper without a slow shell utility."""
    if process.poll() is None:
        try:
            if sys.platform.startswith("win"):
                # Popen.kill() maps directly to TerminateProcess on Windows.
                process.kill()
            else:
                os.killpg(os.getpgid(process.pid), signal.SIGKILL)
        except OSError:
            try:
                process.kill()
            except OSError:
                pass

    # Drain/close the pipe, but keep cleanup strictly bounded so a conversion
    # timeout remains a real deadline even if Windows process teardown misbehaves.
    try:
        process.communicate(timeout=0.5)
    except (OSError, subprocess.SubprocessError):
        if process.stdout is not None:
            try:
                process.stdout.close()
            except OSError:
                pass


def _run_native_with_timeout(
    backend_kind: Backend, src: Path, dest_pdf: Path, timeout: int
) -> None:
    """Run native Office export behind a real process timeout boundary."""
    pid_file = dest_pdf.parent / ".office2pdf_native_pid"
    try:
        pid_file.unlink(missing_ok=True)
    except OSError:
        pass

    command = _native_worker_command(backend_kind, src, dest_pdf, pid_file)
    popen_kwargs: dict[str, object] = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.STDOUT,
    }
    if sys.platform.startswith("win"):
        popen_kwargs["creationflags"] = (
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    else:
        popen_kwargs["start_new_session"] = True

    process = subprocess.Popen(command, **popen_kwargs)  # type: ignore[arg-type]
    office_pid: int | None = None
    try:
        output, _ = process.communicate(timeout=timeout)
        office_pid = _read_recorded_pid(pid_file)
    except subprocess.TimeoutExpired:
        office_pid = _read_recorded_pid(pid_file)
        _terminate_recorded_office_process(backend_kind, office_pid)
        _terminate_native_worker(process)
        raise TimeoutError(
            f"{backend_kind.value} did not finish within {timeout} seconds"
        ) from None
    finally:
        try:
            pid_file.unlink(missing_ok=True)
        except OSError:
            pass

    text = (output or b"").decode(errors="replace").strip()
    if process.returncode != 0:
        # The worker normally quits Office in its finally block.  If it crashed,
        # clean up only the PID it recorded and only after image verification.
        _terminate_recorded_office_process(backend_kind, office_pid)
        raise RuntimeError(
            f"native worker exited {process.returncode}: "
            f"{text or '(no diagnostic output)'}"
        )


def validate_pdf(path: Path) -> tuple[bool, str]:
    """Perform dependency-free structural sanity checks on a PDF."""
    try:
        size = path.stat().st_size
        if size < MIN_PDF_SIZE:
            return False, f"file is too small to be a valid PDF ({size} bytes)"

        with path.open("rb") as handle:
            header = handle.read(8)
            if not header.startswith(b"%PDF-"):
                return False, "missing %PDF header"

            handle.seek(max(0, size - 4096))
            tail = handle.read()
            if b"%%EOF" not in tail:
                return False, "missing PDF end marker"
    except OSError as exc:
        return False, f"could not inspect PDF: {exc}"

    return True, "ok"


def _terminate_process_tree(process: subprocess.Popen[bytes]) -> None:
    """Best-effort termination of LibreOffice and any child processes."""
    if process.poll() is not None:
        return

    try:
        if sys.platform.startswith("win"):
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
                check=False,
                creationflags=creationflags,
            )
        else:
            os.killpg(os.getpgid(process.pid), signal.SIGKILL)
    except (OSError, subprocess.SubprocessError):
        try:
            process.kill()
        except OSError:
            pass

    try:
        process.communicate(timeout=5)
    except (OSError, subprocess.SubprocessError):
        pass


def _run_soffice(
    args: list[str], env: dict[str, str], timeout: int
) -> tuple[int | None, str, bool]:
    """Run LibreOffice and return (returncode, combined output, timed_out)."""
    popen_kwargs: dict[str, object] = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.STDOUT,
        "env": env,
    }
    if sys.platform.startswith("win"):
        popen_kwargs["creationflags"] = (
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    else:
        popen_kwargs["start_new_session"] = True

    process = subprocess.Popen(args, **popen_kwargs)  # type: ignore[arg-type]
    try:
        output, _ = process.communicate(timeout=timeout)
        text = (output or b"").decode(errors="replace").strip()
        return process.returncode, text, False
    except subprocess.TimeoutExpired:
        _terminate_process_tree(process)
        return None, "", True


def _failure(
    src: Path,
    message: str,
    started: float,
    attempts: int = 0,
    backend: str = Backend.LIBREOFFICE.value,
) -> ConversionResult:
    return ConversionResult(
        source=src,
        output=None,
        status=ConversionStatus.FAILED,
        attempts=attempts,
        message=message,
        seconds=max(0.0, time.monotonic() - started),
        backend=backend,
    )


def convert_one(
    src: Path,
    out_dir: Path,
    soffice_bin: str | None,
    timeout: int,
    retries: int,
    overwrite: bool,
    prefer_native: bool = False,
) -> ConversionResult:
    """Convert one source file using staged, atomic destination replacement.

    `soffice_bin` may be None if LibreOffice isn't installed -- that's only
    fatal if the chosen backend actually needs it (i.e. no native Office
    application applies to this file type).
    """
    src = Path(src)
    out_dir = Path(out_dir)
    started = time.monotonic()

    if timeout < 1:
        return _failure(src, "timeout must be at least 1 second", started)
    if retries < 0:
        return _failure(src, "retries cannot be negative", started)
    if not src.is_file():
        return _failure(src, "source file does not exist or is not a file", started)
    if src.name.startswith("~$"):
        return _failure(src, "Microsoft Office temporary/lock file was ignored", started)
    if src.suffix.casefold() not in SUPPORTED_EXTENSIONS:
        return _failure(src, f"unsupported file type: {src.suffix or '(none)'}", started)

    backend_kind = choose_backend(src, prefer_native)
    if backend_kind == Backend.LIBREOFFICE and not soffice_bin:
        return _failure(
            src,
            "no LibreOffice binary available and no native Office application "
            "applies to this file type",
            started,
            backend=backend_kind.value,
        )

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return _failure(src, f"could not create output directory: {exc}", started, backend=backend_kind.value)

    dest = out_dir / f"{src.stem}.pdf"

    if dest.exists() and not overwrite:
        valid, reason = validate_pdf(dest)
        if valid:
            return ConversionResult(
                src, dest, ConversionStatus.SKIPPED, 0,
                "skipped (valid output already exists)",
                time.monotonic() - started,
                backend_kind.value,
            )
        return _failure(
            src,
            f"existing output is not a valid PDF ({reason}); enable overwrite or remove it",
            started,
            backend=backend_kind.value,
        )

    last_error = "unknown failure"
    max_attempts = retries + 1

    for attempt in range(1, max_attempts + 1):
        try:
            # Staging inside out_dir keeps os.replace() on the same filesystem,
            # which gives atomic replacement semantics on Windows and POSIX.
            with tempfile.TemporaryDirectory(
                prefix=".office2pdf_stage_",
                dir=out_dir,
                ignore_cleanup_errors=True,
            ) as stage_dir, tempfile.TemporaryDirectory(
                prefix="office2pdf_profile_",
                ignore_cleanup_errors=True,
            ) as profile_dir:
                stage_path = Path(stage_dir)
                staged_pdf = stage_path / dest.name

                if backend_kind == Backend.LIBREOFFICE:
                    args = [
                        soffice_bin,
                        f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
                        "--headless",
                        "--nologo",
                        "--nodefault",
                        "--nofirststartwizard",
                        "--norestore",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(stage_path),
                        str(src.resolve()),
                    ]
                    env = os.environ.copy()
                    env["SAL_USE_VCLPLUGIN"] = "svp"

                    returncode, output_text, timed_out = _run_soffice(args, env, timeout)
                    if timed_out:
                        last_error = f"LibreOffice timed out after {timeout} seconds"
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue

                    if returncode != 0:
                        last_error = (
                            f"LibreOffice exited {returncode}: "
                            f"{output_text or '(no diagnostic output)'}"
                        )
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue

                    valid, reason = validate_pdf(staged_pdf)
                    if not valid:
                        last_error = (
                            "LibreOffice reported success but the staged PDF was invalid: "
                            f"{reason}. Output: {output_text or '(none)'}"
                        )
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue
                else:
                    try:
                        _run_native_with_timeout(backend_kind, src, staged_pdf, timeout)
                    except TimeoutError as exc:
                        last_error = str(exc)
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue
                    except Exception as exc:  # COM automation failure of any kind
                        last_error = f"{backend_kind.value} export failed: {type(exc).__name__}: {exc}"
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue

                    valid, reason = validate_pdf(staged_pdf)
                    if not valid:
                        last_error = (
                            f"{backend_kind.value} reported success but the staged PDF "
                            f"was invalid: {reason}"
                        )
                        LOG.warning(
                            "Attempt %d/%d for %s: %s",
                            attempt, max_attempts, src.name, last_error,
                        )
                        continue

                # Re-check overwrite policy immediately before committing to
                # prevent a race with another process creating the destination.
                if dest.exists() and not overwrite:
                    valid_existing, existing_reason = validate_pdf(dest)
                    if valid_existing:
                        return ConversionResult(
                            src, dest, ConversionStatus.SKIPPED, attempt,
                            "skipped (valid output appeared during conversion)",
                            time.monotonic() - started,
                            backend_kind.value,
                        )
                    last_error = (
                        "destination appeared during conversion but is invalid "
                        f"({existing_reason}); refusing to replace without overwrite"
                    )
                    continue

                os.replace(staged_pdf, dest)
                final_valid, final_reason = validate_pdf(dest)
                if not final_valid:
                    last_error = f"committed PDF failed validation: {final_reason}"
                    LOG.warning(
                        "Attempt %d/%d for %s: %s",
                        attempt, max_attempts, src.name, last_error,
                    )
                    continue

                return ConversionResult(
                    src, dest, ConversionStatus.SUCCESS, attempt, "ok",
                    time.monotonic() - started,
                    backend_kind.value,
                )

        except (OSError, subprocess.SubprocessError) as exc:
            last_error = f"operating-system error: {exc}"
            LOG.warning(
                "Attempt %d/%d for %s: %s",
                attempt, max_attempts, src.name, last_error,
            )
        except Exception as exc:  # defensive terminal-state guarantee
            last_error = f"unexpected conversion error: {type(exc).__name__}: {exc}"
            LOG.exception("Unexpected conversion failure for %s", src)

    if backend_kind is not Backend.LIBREOFFICE and soffice_bin:
        LOG.warning(
            "Native %s export failed for %s; trying LibreOffice fallback: %s",
            backend_kind.value,
            src.name,
            last_error,
        )
        fallback = convert_one(
            src,
            out_dir,
            soffice_bin,
            timeout,
            retries,
            overwrite,
            prefer_native=False,
        )
        combined_message = (
            f"native {backend_kind.value} failed ({last_error}); "
            f"LibreOffice fallback: {fallback.message}"
        )
        return dataclasses.replace(
            fallback,
            attempts=max_attempts + fallback.attempts,
            message=combined_message,
            seconds=max(0.0, time.monotonic() - started),
        )

    return _failure(src, last_error, started, max_attempts, backend=backend_kind.value)


def _iter_directory_files(directory: Path, recursive: bool) -> Iterable[Path]:
    iterator = directory.rglob("*") if recursive else directory.iterdir()
    for candidate in iterator:
        try:
            if not candidate.is_file():
                continue
        except OSError:
            continue
        if candidate.name.startswith(("~$", ".")):
            continue
        if candidate.suffix.casefold() in SUPPORTED_EXTENSIONS:
            yield candidate


def discover_inputs(paths: list[str], recursive: bool) -> list[Path]:
    """Discover supported files case-insensitively and deterministically."""
    found: list[Path] = []
    for raw in paths:
        path = Path(raw)
        try:
            if path.is_dir():
                found.extend(_iter_directory_files(path, recursive))
            elif path.is_file():
                if path.name.startswith("~$"):
                    LOG.warning("Skipping Microsoft Office temporary file: %s", path)
                elif path.suffix.casefold() in SUPPORTED_EXTENSIONS:
                    found.append(path)
                else:
                    LOG.warning("Skipping unsupported file type: %s", path)
            else:
                LOG.error("Path does not exist: %s", path)
        except OSError as exc:
            LOG.error("Could not inspect path %s: %s", path, exc)

    unique: dict[str, Path] = {}
    for candidate in found:
        try:
            key = os.path.normcase(str(candidate.resolve()))
        except OSError:
            key = os.path.normcase(str(candidate.absolute()))
        unique.setdefault(key, candidate)

    return sorted(unique.values(), key=lambda p: str(p).casefold())


def find_destination_collisions(
    inputs: list[Path], output_dir: str | None
) -> dict[Path, list[Path]]:
    planned: dict[Path, list[Path]] = {}
    for src in inputs:
        out_dir = Path(output_dir) if output_dir else src.resolve().parent
        dest = (out_dir / f"{src.stem}.pdf").resolve()
        planned.setdefault(dest, []).append(src)
    return {dest: sources for dest, sources in planned.items() if len(sources) > 1}


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return parsed


def _nonnegative_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("cannot be negative")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="office2pdf",
        description="Convert Office documents to PDF safely.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {APP_VERSION}",
    )
    parser.add_argument("paths", nargs="+", help="Files and/or directories to convert.")
    parser.add_argument(
        "-o", "--output-dir", default=None,
        help="Directory for PDFs (default: beside each source file).",
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="Search queued directories recursively.",
    )
    parser.add_argument(
        "-j", "--jobs", type=_positive_int, default=1,
        help="Parallel conversions (default: 1).",
    )
    parser.add_argument(
        "--timeout", type=_positive_int, default=DEFAULT_TIMEOUT,
        help=f"Per-file timeout in seconds (default: {DEFAULT_TIMEOUT}).",
    )
    parser.add_argument(
        "--retries", type=_nonnegative_int, default=DEFAULT_RETRIES,
        help=f"Additional attempts after the first failure (default: {DEFAULT_RETRIES}).",
    )
    parser.add_argument(
        "--overwrite", action="store_true",
        help="Atomically replace an existing valid PDF.",
    )
    native_group = parser.add_mutually_exclusive_group()
    native_group.add_argument(
        "--native-office",
        dest="prefer_native",
        action="store_true",
        help=(
            "Use the installed Word/Excel/PowerPoint application when available "
            "(Windows only; opt-in until verified on the target PC)."
        ),
    )
    native_group.add_argument(
        "--no-native-office",
        dest="prefer_native",
        action="store_false",
        help=argparse.SUPPRESS,
    )
    parser.set_defaults(prefer_native=False)
    parser.add_argument("-v", "--verbose", action="store_true", help="Debug logging.")
    return parser


def _safe_convert(
    src: Path,
    output_dir: str | None,
    soffice_bin: str | None,
    timeout: int,
    retries: int,
    overwrite: bool,
    prefer_native: bool = False,
) -> ConversionResult:
    started = time.monotonic()
    try:
        out_dir = Path(output_dir) if output_dir else src.resolve().parent
        return convert_one(src, out_dir, soffice_bin, timeout, retries, overwrite, prefer_native)
    except Exception as exc:  # final containment boundary for batch workers
        LOG.exception("Unhandled batch worker error for %s", src)
        return _failure(
            src,
            f"unexpected worker error: {type(exc).__name__}: {exc}",
            started,
        )


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    prefer_native = args.prefer_native

    soffice_bin: str | None
    try:
        soffice_bin = find_soffice()
        LOG.debug("Using LibreOffice binary: %s", soffice_bin)
    except LibreOfficeNotFoundError as exc:
        soffice_bin = None
        # Not automatically fatal: a native Office backend might still
        # handle every queued file on Windows. convert_one() reports a
        # clear per-file failure for anything that actually needed
        # LibreOffice and didn't have it.
        LOG.warning("%s", exc)
        if not (sys.platform.startswith("win") and prefer_native):
            return 3

    inputs = discover_inputs(args.paths, args.recursive)
    if not inputs:
        LOG.error("No convertible files found among the given paths.")
        return 2

    collisions = find_destination_collisions(inputs, args.output_dir)
    if collisions:
        LOG.error(
            "Refusing to run: %d output filename collision(s) detected.",
            len(collisions),
        )
        for dest, sources in collisions.items():
            LOG.error("  %s <- %s", dest, ", ".join(str(source) for source in sources))
        return 2

    LOG.info("Found %d file(s) to convert.", len(inputs))

    run = lambda source: _safe_convert(  # noqa: E731 - compact pool callback
        source,
        args.output_dir,
        soffice_bin,
        args.timeout,
        args.retries,
        args.overwrite,
        prefer_native,
    )

    results: list[ConversionResult] = []
    if args.jobs > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.jobs) as pool:
            for result in pool.map(run, inputs):
                results.append(result)
                _log_result(result)
    else:
        for source in inputs:
            result = run(source)
            results.append(result)
            _log_result(result)

    succeeded = sum(result.status is ConversionStatus.SUCCESS for result in results)
    skipped = sum(result.status is ConversionStatus.SKIPPED for result in results)
    failed = sum(result.status is ConversionStatus.FAILED for result in results)
    LOG.info("Done: %d converted, %d skipped, %d failed.", succeeded, skipped, failed)

    if failed:
        LOG.error("Failed conversions:")
        for result in results:
            if result.status is ConversionStatus.FAILED:
                LOG.error("  %s -- %s", result.source, result.message)
        return 1
    return 0


def _log_result(result: ConversionResult) -> None:
    if result.status is ConversionStatus.SUCCESS:
        LOG.info(
            "OK   [%s] %s -> %s (%.1fs, attempt %d)",
            result.backend, result.source, result.output, result.seconds, result.attempts,
        )
    elif result.status is ConversionStatus.SKIPPED:
        LOG.info("SKIP %s -> %s (%s)", result.source, result.output, result.message)
    elif result.status is ConversionStatus.CANCELLED:
        LOG.warning("CANCEL %s -- %s", result.source, result.message)
    else:
        LOG.error(
            "FAIL %s -- %s (%.1fs, %d attempt(s))",
            result.source, result.message, result.seconds, result.attempts,
        )


if __name__ == "__main__":
    sys.exit(main())
