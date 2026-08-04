# Office2PDF command line

A standard-library Python wrapper around LibreOffice headless conversion.

## Requirements

- Python 3.10 or newer
- LibreOffice installed

No third-party Python packages are required for the CLI backend.

## Examples

```text
python office2pdf.py report.docx budget.xlsx
python office2pdf.py documents --output-dir pdfs
python office2pdf.py documents --recursive --output-dir pdfs
python office2pdf.py documents --jobs 2 --overwrite
```

Defaults are deliberately modest: one conversion at a time, a 120-second timeout, and two retries.

## Exit codes

- `0` — all sources converted or valid existing outputs were skipped
- `1` — at least one conversion failed
- `2` — no inputs were found or output-name collisions were detected
- `3` — LibreOffice was not found

## Safety behaviour

- Output is produced in a private staging directory.
- A staged file must pass PDF sanity checks.
- The destination is committed with atomic replacement only after validation.
- Without `--overwrite`, a valid existing PDF is skipped.
- Without `--overwrite`, an invalid existing PDF causes a failure rather than a silent skip.
- A timed-out LibreOffice process tree is terminated before retrying.
- Folder discovery is case-insensitive and excludes Office `~$` lock files.
- Two sources that would produce the same destination filename are refused before work begins.

## Supported formats

DOC/DOCX, XLS/XLSX, PPT/PPTX, ODT/ODS/ODP, RTF, CSV, and related template formats supported by the backend.
