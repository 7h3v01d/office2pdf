# Changelog

## Reliability pass — 2026-08-03

### Fixed

- Eliminated the stale-output overwrite false positive. Every conversion now writes to a private staging directory and only commits a validated new PDF with `os.replace()`.
- Removed the Linux `LD_PRELOAD` / dynamically compiled socket shim from the product code.
- Converted LibreOffice launch errors, file-system errors, invalid destinations, and unexpected worker errors into structured failures rather than uncaught exceptions.
- Added process-tree termination when LibreOffice exceeds its timeout.
- Reworked parallel cancellation so queued work is not submitted after cancellation and cancelled futures cannot crash the worker.
- Made window closing safe during conversion. The application cancels queued work and closes only after active conversions finish.
- Prevented premature destruction of a still-running `QThread`.
- Added case-insensitive, deterministic file discovery.
- Excluded Microsoft Office lock files such as `~$newsletter.docx` and hidden dot-files.
- Added positive-range validation for jobs, timeout, and retry CLI arguments.
- Replaced backend `sys.exit()` calls with `LibreOfficeNotFoundError`.
- Enabled **Open Output Folder** only when a valid output exists.

### Added

- Dependency-free PDF validation using file size, `%PDF-` header, and `%%EOF` marker checks.
- Nine backend regression tests covering overwrite safety, stale outputs, launch failures, invalid existing PDFs, discovery, collision detection, and argument validation.
- Windows setup, test, and PyInstaller one-folder build scripts.
- Pinned runtime and build requirement files.

### Verified

- Real DOCX and XLSX samples converted successfully through LibreOffice.
- Both generated PDFs passed structural validation and opened with a PDF parser.
