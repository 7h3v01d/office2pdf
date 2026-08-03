# Office2PDF

A Windows-friendly desktop converter for Word, Excel, PowerPoint, and other LibreOffice-readable documents.

The intended workflow is simple: add files or folders, choose where the PDFs should go, and click **Convert**. The backend is deliberately conservative so an old PDF is never mistaken for a newly converted file.

<img width="1042" height="723" alt="screenshot" src="https://github.com/user-attachments/assets/d62332db-6a47-4e88-8ab5-15ec9afe68c3" />

---

## Reliability guarantees

- LibreOffice writes each result into a private staging directory.
- A staged result must have a plausible PDF structure before it is accepted.
- The final destination is replaced atomically only after validation.
- A timeout terminates the LibreOffice process tree.
- Every source ends as converted, skipped, failed, or not started because the run was cancelled.
- Microsoft Office temporary files such as `~$newsletter.docx` are ignored.
- Same-name destination collisions are refused before conversion begins.

## Windows development setup

1. Install Python 3.11 or newer.
2. Install LibreOffice. The normal Windows installer path is detected automatically.
3. Double-click `setup.bat`.
4. Double-click `run.bat`.

## Build the Windows application

Double-click `build_exe.bat` on a Windows 11 computer. It runs the regression tests before building.

The output is:

```text
dist\Office2PDF\Office2PDF.exe
```

This is intentionally a **one-folder** release for the first dependable build. Copy or install the entire `dist\Office2PDF` folder. Do not copy only the EXE, because its supporting runtime files live beside it.

Once this version has been tested on Dad's Windows 11 Home machine, it can be wrapped in an installer and optionally converted to a one-file portable build.

## Run tests

Double-click `test.bat`, or run:

```text
python -m unittest discover -s tests -v
```

## Current conversion engine

The current release uses LibreOffice. A later fidelity upgrade can prefer native Microsoft Word/Excel PDF export when Microsoft Office is installed, with LibreOffice retained as the fallback.

## Files

- `gui.py` — PyQt6 desktop interface
- `office2pdf.py` — hardened conversion backend and CLI
- `tests/` — regression tests
- `setup.bat` — creates the Python environment
- `run.bat` — starts the GUI without a console window
- `test.bat` — compiles and runs tests
- `build_exe.bat` — creates the Windows one-folder EXE release
- `README_cli.md` — command-line usage
- `CHANGELOG.md` — completed reliability work
