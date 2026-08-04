# Office2PDF 1.0.0 — Professional Release Pass

Build date: 4 August 2026

## New user-facing features

- Five-second branded startup splash matching the supplied dark navy and amber artwork.
- Animated loading indicator with the Office2PDF icon, release version and Leon Priest credit.
- New top-level `About` tab.
- Overview, full GPL licence and third-party notices are readable inside the application.
- Version 1.0.0 and build date are displayed consistently.

## Windows release integration

- Windows Explorer file metadata now reports product name, version, description and copyright.
- The main executable and isolated Office worker both receive version resources.
- The final release folder contains `LICENSE.txt`, `THIRD_PARTY_NOTICES.txt` and `SOURCE_OFFER.txt`.
- `Office2PDF-1.0.0-source.zip` is generated automatically during the build.
- The shortcut creator no longer requires PowerShell and refuses PyInstaller's incomplete `build` executable.

## Licence decision

Office2PDF is released under `GPL-3.0-or-later`. The packaged application uses the GPL edition of PyQt6, so the release includes the complete GPL text and corresponding source code.

## Build

Run:

```bat
build_exe.bat
```

The final application is created at:

```text
dist\Office2PDF\Office2PDF.exe
```

Copy the complete `dist\Office2PDF` directory.
