# Changelog

## 1.0.0 — 4 August 2026

- Corrected all author, copyright and Windows metadata attribution to the full name **Leon Priest**.
- Added a five-second branded Office2PDF splash screen.
- Added a professional About tab with overview, version, author, build date, licence and third-party notices.
- Added Windows file-version resources for both executables.
- Adopted GPL-3.0-or-later for the distributable PyQt6 build.
- Added bundled licence notices and an automatically generated corresponding-source archive.
- Replaced the shortcut creator with the no-PowerShell, release-folder-validating implementation.
- Expanded the suite to 37 tests.
- Consolidated technical documentation under `docs/` and documentation artwork under `docs/images/`.
- Made corresponding-source packaging recursive for `assets/`, `docs/`, and `tests/` so future file organisation changes do not require a brittle manifest update.

## Branded icon integration — 2026-08-04

### Added

- Added the supplied multi-resolution `office2pdf.ico` unchanged under `assets/`.
- Embedded the icon into both PyInstaller executables.
- Added PyInstaller-safe runtime resource lookup for source and packaged execution.
- Applied the icon to the QApplication, main title bar and Windows taskbar identity.
- Added a release-folder shortcut creator for branded Desktop and Start Menu shortcuts.
- Added two branding regression tests covering ICO structure and PyInstaller build flags.

### Build

- `build_exe.bat` now refuses to build if the icon asset is missing.
- The icon is bundled as runtime data and copied shortcut tooling is verified in the release folder.

## Windows deadline hotfix — 2026-08-04

### Fixed

- Removed the slow `taskkill` subprocess from native Office timeout cleanup.
- Native Office PID verification and termination now use one Win32 process handle, preventing a verify-then-kill PID-reuse race.
- The isolated native worker is terminated directly and its pipe is reaped with a 0.5-second upper bound.
- A one-second native timeout now returns at approximately one second in the controlled sleeper regression instead of taking nine seconds on Windows.

### Verification

- Full local suite: 27 tests, 19 passed, 8 platform-specific skipped, 0 failed.
- Controlled five-second sleeper with a one-second deadline: `1.005s`.
- On Windows without Microsoft Office, the expected result is 18 passed, 9 Office/platform-specific skipped, 0 failed.

## Native safety pass — 2026-08-04

### Fixed

- Replaced the native `ThreadPoolExecutor` timeout, which still waited for the timed-out COM thread, with a real companion-process deadline.
- Removed image-wide Office termination that could close unrelated unsaved Word, Excel or PowerPoint sessions.
- Added PID-scoped timeout cleanup with executable-image verification.
- Corrected 64-bit Win32 handle declarations for HWND/PID and process-image queries.
- Added automatic LibreOffice fallback when a detected native Office application fails to export.
- Changed native Office from enabled-by-default to explicit opt-in until real Windows validation passes.
- Made native source and destination paths absolute inside the worker.
- Updated build scripts to package and smoke-test the companion worker.
- Standardised development setup on 64-bit Python 3.11.

### Security

- Native Office macros are force-disabled before opening a source document.
- Native automation is isolated from the GUI and batch worker process.
- Timeout cleanup can target only the dedicated `DispatchEx` instance recorded by the worker.

### Tests

- Added a regression proving the native timeout returns near its deadline.
- Added native failure-to-LibreOffice fallback coverage.
- Added a default-opt-in policy test.
- Added real Windows-only Excel backend selection, export and overwrite tests.
- Local Linux result for that pass: 27 tests, 19 passed, 8 Windows/Office-only skipped, 0 failed.

### Verified

- Real DOCX and XLSX samples converted concurrently through actual LibreOffice.
- Both generated PDFs passed structural and parser validation and rendered as one-page documents.

## Second pass — 2026-08-03

- Added native Word/Excel/PowerPoint export, real LibreOffice end-to-end tests and native Word tests.
- Corrected the LibreOffice download reference.

## Reliability pass — 2026-08-03

- Added staged atomic output, PDF validation, exception containment, process-tree timeout cleanup, cancellation safety, deterministic discovery, collision detection, Windows scripts and the initial regression suite.
