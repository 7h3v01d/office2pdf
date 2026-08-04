# Office2PDF command line

## Requirements

- Python 3.11
- LibreOffice installed for the default backend
- Windows + pywin32 + Microsoft Office only when `--native-office` is requested

## Examples

```text
python office2pdf.py report.docx budget.xlsx
python office2pdf.py documents --output-dir pdfs
python office2pdf.py documents --recursive --output-dir pdfs
python office2pdf.py documents --jobs 2 --overwrite
python office2pdf.py report.docx --native-office
```

Defaults are conservative: LibreOffice, one conversion at a time, a 120-second per-attempt timeout, and two retries.

`--native-office` opts into installed Word, Excel or PowerPoint on Windows. Native automation runs in an isolated helper process. If it fails and LibreOffice is available, conversion falls back to LibreOffice.

## Exit codes

- `0` — all sources converted or valid existing outputs were skipped
- `1` — at least one conversion failed
- `2` — no inputs were found or output-name collisions were detected
- `3` — LibreOffice was unavailable when required

## Safety behaviour

- Output is created in a private staging directory.
- A staged file must pass PDF sanity checks.
- The destination is committed atomically only after validation.
- Without `--overwrite`, a valid existing PDF is skipped.
- Without `--overwrite`, an invalid existing PDF is reported as a failure.
- Timed-out LibreOffice process trees are terminated before retrying.
- Timed-out native Office work is isolated in a helper process and cleanup is PID-scoped.
- Office macros are force-disabled for native automation.
- Folder discovery excludes Office `~$` lock files and hidden dot-files.
- Two sources that map to the same PDF filename are refused before work begins.

## Supported formats

DOC/DOCX, XLS/XLSX, PPT/PPTX, ODT/ODS/ODP, RTF, CSV and related template formats listed by `office2pdf.SUPPORTED_EXTENSIONS`.

## Version

```bat
python office2pdf.py --version
```

Reports `office2pdf 1.0.0`.
