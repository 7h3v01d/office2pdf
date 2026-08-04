# Office2PDF 1.0.0

Office2PDF is a Windows desktop converter for Word, Excel and PowerPoint documents. Version 1.0.0 adds a branded five-second splash screen, a full About tab, embedded Windows version information and release licence notices.

A Windows-friendly desktop converter for DOC/DOCX, XLS/XLSX, PPT/PPTX and other office-document formats.

The primary target is a dependable Windows 11 application for ordinary users: add a file, choose an output location, and create a PDF without risking an existing good copy.

<img width="1042" height="836" alt="screenshot" src="https://github.com/user-attachments/assets/9d208186-b0c1-4b33-8fcf-08a4fa7f51c1" />

---

## Current safety model

**LibreOffice is the default backend.** Each conversion uses an isolated profile and a private staging directory. The staged result must pass PDF sanity checks before `os.replace()` atomically commits it to the final destination.

An optional Microsoft Office backend is available on Windows for closer Word/Excel/PowerPoint layout fidelity. It is deliberately **off by default until tested on the target PC**. When enabled:

- COM automation runs in `Office2PDFNativeWorker`, not in the GUI process.
- A real process timeout can terminate the helper instead of waiting on an unkillable Python thread.
- The helper records the PID of its dedicated `DispatchEx` Office instance.
- Timeout cleanup verifies and terminates the dedicated Office PID through one Win32 handle; the native path launches no `taskkill` command and cannot target unrelated Office sessions.
- Office macro automation is force-disabled before a document is opened.
- A failed native export falls back to LibreOffice when LibreOffice is available.

## Reliability guarantees

- Existing PDFs are never used as evidence that a new conversion succeeded.
- A failed overwrite attempt leaves the previous PDF untouched.
- Output is validated before atomic commit.
- LibreOffice and native-worker timeouts have process-level cleanup.
- Every input ends as converted, skipped, failed, or not started because the run was cancelled.
- Office lock files such as `~$newsletter.docx` and hidden dot-files are ignored.
- Discovery is deterministic and case-insensitive.
- Same-name destination collisions are refused before conversion begins.
- Unexpected filesystem, process and worker errors become structured failures rather than crashing the batch.

## Windows development setup

1. Install 64-bit Python 3.11.
2. Install LibreOffice. `Libre_Office_Link.txt` points to the official download page.
3. Double-click `setup.bat`.
4. Double-click `test.bat`.
5. Double-click `run.bat`.

Microsoft Word and Excel are optional. They are required only to exercise the opt-in native backend.

## Build the Windows application

Run `build_exe.bat` on Windows 11. It compiles the sources, runs the regression suite, builds the isolated Office worker, builds the GUI, copies both into the release folder, and smoke-tests the worker executable.

The supplied `assets/office2pdf.ico` is embedded into the executable and also loaded at runtime for the title bar and Windows taskbar.

The release is:

```text
dist\Office2PDF\
    Office2PDF.exe
    Office2PDFNativeWorker.exe
    Create Office2PDF Shortcuts.bat
    ...supporting runtime files...
```

Copy or install the **entire** `dist\Office2PDF` folder. Do not copy only `Office2PDF.exe`. After placing the folder in its permanent location, run `Create Office2PDF Shortcuts.bat` to create branded Desktop and Start Menu shortcuts.

The first release intentionally uses PyInstaller one-folder mode because it is easier to diagnose and more dependable than jumping directly to one-file packaging.

## Native Office validation

Leave **Use Microsoft Office for best layout fidelity (experimental)** unchecked for Dad's first LibreOffice-only test.

After that path is proven, follow `WINDOWS_VALIDATION_CHECKLIST.md` to test native Word and Excel export. `test.bat` contains real end-to-end Word and Excel tests that automatically run when Windows, pywin32, LibreOffice and the matching Office application are present.

## Verification completed in this environment

- Python compilation passed for the backend, GUI, native worker and tests.
- Local suite: 29 tests, 21 passed and 8 platform-specific skipped. On Windows without Office, expect 20 passed and 9 Office/platform-specific skips.
- The native timeout regression returned at the configured process deadline rather than waiting for the worker to finish.
- Native failure-to-LibreOffice fallback passed.
- Real DOCX and XLSX files converted concurrently through actual LibreOffice.
- Both PDFs passed structural validation, opened with a PDF parser and contained one page.

The remaining boundary is genuine Windows validation of COM automation and the PyInstaller executables.

## Project files

- `gui.py` — PyQt6 desktop interface
- `office2pdf.py` — conversion backend and CLI
- `native_office_worker.py` — isolated Microsoft Office COM worker
- `tests/test_office2pdf.py` — mocked and real end-to-end regression tests
- `setup.bat`, `run.bat`, `test.bat` — Windows development workflow
- `build_exe.bat` — test-gated Windows build with embedded application icon
- `assets/office2pdf.ico` — title-bar, taskbar, executable and shortcut icon
- `create_shortcuts.bat` — creates Desktop and Start Menu shortcuts for the built release
- `WINDOWS_VALIDATION_CHECKLIST.md` — target-PC proving steps
- `REVIEW_FINDINGS.md` — findings from the incoming native-backend update
- `RELIABILITY_REPORT.md` — consolidated technical status

## Licence

Office2PDF is free software released under the GNU General Public License version 3 or later (`GPL-3.0-or-later`). The complete licence is in `LICENSE.txt`; third-party component notices are in `THIRD_PARTY_NOTICES.txt`. The Windows build also contains `Office2PDF-1.0.0-source.zip`.
