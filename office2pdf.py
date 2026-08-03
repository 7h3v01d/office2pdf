#!/usr/bin/env python3
"""Reliable Office document to PDF conversion through LibreOffice.

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

LOG = logging.getLogger("office2pdf")

SUPPORTED_EXTENSIONS = frozenset(
    {
        ".docx", ".doc", ".dotx", ".dot", ".odt", ".rtf",
        ".xlsx", ".xls", ".xltx", ".xlt", ".ods", ".csv",
        ".pptx", ".ppt", ".odp",
    }
)

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


@dataclasses.dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path | None
    status: ConversionStatus
    attempts: int
    message: str
    seconds: float

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
) -> ConversionResult:
    return ConversionResult(
        source=src,
        output=None,
        status=ConversionStatus.FAILED,
        attempts=attempts,
        message=message,
        seconds=max(0.0, time.monotonic() - started),
    )


def convert_one(
    src: Path,
    out_dir: Path,
    soffice_bin: str,
    timeout: int,
    retries: int,
    overwrite: bool,
) -> ConversionResult:
    """Convert one source file using staged, atomic destination replacement."""
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

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return _failure(src, f"could not create output directory: {exc}", started)

    dest = out_dir / f"{src.stem}.pdf"

    if dest.exists() and not overwrite:
        valid, reason = validate_pdf(dest)
        if valid:
            return ConversionResult(
                src, dest, ConversionStatus.SKIPPED, 0,
                "skipped (valid output already exists)",
                time.monotonic() - started,
            )
        return _failure(
            src,
            f"existing output is not a valid PDF ({reason}); enable overwrite or remove it",
            started,
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

                # Re-check overwrite policy immediately before committing to
                # prevent a race with another process creating the destination.
                if dest.exists() and not overwrite:
                    valid_existing, existing_reason = validate_pdf(dest)
                    if valid_existing:
                        return ConversionResult(
                            src, dest, ConversionStatus.SKIPPED, attempt,
                            "skipped (valid output appeared during conversion)",
                            time.monotonic() - started,
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

    return _failure(src, last_error, started, max_attempts)


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
        description="Convert Office documents to PDF through LibreOffice.",
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
    parser.add_argument("-v", "--verbose", action="store_true", help="Debug logging.")
    return parser


def _safe_convert(
    src: Path,
    output_dir: str | None,
    soffice_bin: str,
    timeout: int,
    retries: int,
    overwrite: bool,
) -> ConversionResult:
    started = time.monotonic()
    try:
        out_dir = Path(output_dir) if output_dir else src.resolve().parent
        return convert_one(src, out_dir, soffice_bin, timeout, retries, overwrite)
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

    try:
        soffice_bin = find_soffice()
    except LibreOfficeNotFoundError as exc:
        LOG.error("%s", exc)
        return 3

    LOG.debug("Using LibreOffice binary: %s", soffice_bin)
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
            "OK   %s -> %s (%.1fs, attempt %d)",
            result.source, result.output, result.seconds, result.attempts,
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
