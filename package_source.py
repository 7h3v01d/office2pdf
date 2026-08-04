#!/usr/bin/env python3
# Copyright (C) 2026 Leon Priest
# SPDX-License-Identifier: GPL-3.0-or-later

"""Create the corresponding-source archive shipped with Office2PDF releases."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from version_info import APP_VERSION

PROJECT_ROOT = Path(__file__).resolve().parent

# Root-level source, build, legal, and user-facing files that are required
# for a complete corresponding-source release.
SOURCE_FILES = (
    ".gitignore",
    "gui.py",
    "office2pdf.py",
    "native_office_worker.py",
    "version_info.py",
    "package_source.py",
    "requirements.txt",
    "requirements-build.txt",
    "build_exe.bat",
    "setup.bat",
    "run.bat",
    "test.bat",
    "create_shortcuts.bat",
    "windows_version_info.txt",
    "windows_worker_version_info.txt",
    "LICENSE.txt",
    "THIRD_PARTY_NOTICES.txt",
    "SOURCE_OFFER.txt",
    "README.md",
    "CHANGELOG.md",
)

# These directories are intentionally recursive so reorganising documentation
# or adding tests/assets does not require updating a brittle file manifest.
SOURCE_DIRECTORIES = (
    "assets",
    "docs",
    "tests",
)


def _iter_source_files() -> list[Path]:
    files: list[Path] = []

    for relative in SOURCE_FILES:
        source = PROJECT_ROOT / relative
        if not source.is_file():
            raise FileNotFoundError(f"required source file missing: {relative}")
        files.append(source)

    for relative in SOURCE_DIRECTORIES:
        directory = PROJECT_ROOT / relative
        if not directory.is_dir():
            raise FileNotFoundError(f"required source directory missing: {relative}")
        files.extend(path for path in directory.rglob("*") if path.is_file())

    return sorted(
        set(files),
        key=lambda path: path.relative_to(PROJECT_ROOT).as_posix().casefold(),
    )


def create_source_archive(output: Path) -> Path:
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    root_name = f"Office2PDF-{APP_VERSION}-source"

    with ZipFile(output, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for source in _iter_source_files():
            relative = source.relative_to(PROJECT_ROOT).as_posix()
            archive.write(source, f"{root_name}/{relative}")

    return output


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / f"Office2PDF-{APP_VERSION}-source.zip",
    )
    args = parser.parse_args()
    result = create_source_archive(args.output)
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
