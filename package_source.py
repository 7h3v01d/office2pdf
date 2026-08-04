#!/usr/bin/env python3
# Copyright (C) 2026 Leon
# SPDX-License-Identifier: GPL-3.0-or-later

"""Create the corresponding-source archive shipped with Office2PDF releases."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from version_info import APP_VERSION

PROJECT_ROOT = Path(__file__).resolve().parent
SOURCE_PATHS = (
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
    "README_cli.md",
    "CHANGELOG.md",
    "PROFESSIONAL_RELEASE.md",
    "assets/office2pdf.ico",
)


def create_source_archive(output: Path) -> Path:
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    root_name = f"Office2PDF-{APP_VERSION}-source"

    with ZipFile(output, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for relative in SOURCE_PATHS:
            source = PROJECT_ROOT / relative
            if not source.is_file():
                raise FileNotFoundError(f"required source file missing: {relative}")
            archive.write(source, f"{root_name}/{relative}")

        tests_dir = PROJECT_ROOT / "tests"
        for source in sorted(tests_dir.glob("test_*.py"), key=lambda p: p.name.casefold()):
            archive.write(source, f"{root_name}/tests/{source.name}")

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
