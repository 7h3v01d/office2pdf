# Office2PDF reliability pass

Date: 3 August 2026

## Work completed

The first stabilisation pass is implemented in the supplied project.

1. **Safe overwrite:** conversion happens in a private staging folder. A validated new PDF is atomically moved into place. An older destination can no longer be mistaken for a successful new conversion.
2. **Failure containment:** LibreOffice launch problems, permission errors, invalid outputs, and unexpected worker errors now return explicit failure results.
3. **Timeout cleanup:** a timed-out LibreOffice process tree is terminated before a retry begins.
4. **Safe cancellation:** the parallel worker submits only a bounded number of jobs. Cancellation stops queued files from starting and cannot raise `CancelledError` through the GUI thread.
5. **Safe window closing:** closing the GUI during a conversion cancels queued work and waits for active conversions to finish rather than destroying a running `QThread`.
6. **Input hygiene:** discovery is deterministic and case-insensitive. Microsoft Office `~$` lock files and hidden dot-files are excluded.
7. **PDF validation:** output must have a plausible size, a `%PDF-` header, and a `%%EOF` marker.
8. **Build path:** setup, test, and PyInstaller one-folder build scripts are included.

## Verification performed

- Python syntax compilation passed for `office2pdf.py` and `gui.py`.
- Nine regression tests passed.
- A real DOCX newsletter and XLSX treasurer report were converted simultaneously through LibreOffice.
- Both generated PDFs were structurally valid and readable through a PDF parser.
- A valid existing PDF was correctly skipped.
- An overwrite replaced a controlled old PDF with the new converted PDF.
- A simulated LibreOffice success that produced no file correctly failed while preserving the old PDF.
- A simulated launch `OSError` became a structured failure.
- A deliberately hung child process was terminated at the configured timeout.

## What remains before Dad receives it

1. Run `setup.bat` and `test.bat` on a Windows 11 machine.
2. Run `build_exe.bat` on Windows to create `dist\Office2PDF\Office2PDF.exe`.
3. Test several of Dad's real newsletters and spreadsheets, especially documents using uncommon fonts, tables, text boxes, print areas, or page scaling.
4. Confirm LibreOffice is installed on Dad's machine, or add a native Microsoft Word/Excel export backend.
5. After field testing, wrap the one-folder build in an installer and add an application icon/version metadata.

## Honest limitation

This environment can verify the backend with LibreOffice, but it cannot produce or execute a trustworthy Windows-native PyInstaller build. PyInstaller builds for Windows must be created and tested on Windows. The project is now prepared for that step; an actual Windows EXE is not included in this archive.
