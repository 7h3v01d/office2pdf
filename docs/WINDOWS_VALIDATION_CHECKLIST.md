# Windows 11 validation checklist

Use this on the development Windows 11 machine before placing the release on Dad's PC.

## 1. Prepare

- Install 64-bit Python 3.11.
- Install LibreOffice.
- Confirm Word and Excel are installed only if testing the optional native backend.
- Extract the project to a normal local folder, not inside the ZIP.

## 2. Create the environment

Run:

```text
setup.bat
```

Expected: `.venv` is created and PyQt6 plus pywin32 install without errors.

## 3. Run all source tests

Run:

```text
test.bat
```

Expected:

- all general and real-LibreOffice tests pass;
- native Word tests run rather than skip when Word is installed;
- native Excel tests run rather than skip when Excel is installed;
- no `WINWORD.EXE` or `EXCEL.EXE` remains after the suite.

A skipped native test means the matching Office application or pywin32 was not detected. Treat that as **not validated**, not as a native-backend pass.

## 4. Prove the default Dad workflow

Run `run.bat` and leave **Use Microsoft Office for best layout fidelity (experimental)** unchecked.

Convert:

1. a real village newsletter DOCX;
2. a real treasurer XLSX;
3. both together;
4. each again with overwrite disabled;
5. each again with overwrite enabled.

Confirm:

- PDFs open correctly;
- the newsletter matches expectations;
- the workbook uses its intended print area/orientation/scaling;
- overwrite disabled produces a clear skip;
- overwrite enabled replaces the PDF;
- cancelling a batch stops unstarted files and leaves active output valid;
- closing during conversion waits safely.

## 5. Prove native Word and Excel

Close test documents, then enable the experimental Microsoft Office checkbox.

Convert the same DOCX and XLSX. The log should show `[word]` and `[excel]` respectively. Compare their layout with the LibreOffice PDFs.

Then deliberately test fallback by temporarily unchecking native mode or, on a development copy only, making the worker unavailable. The application must fail clearly or use LibreOffice; it must not report a stale PDF as new.

Do not intentionally force a hung Office document on Dad's production machine.

## 6. Build

Run:

```text
build_exe.bat
```

Expected release:

```text
dist\Office2PDF\Office2PDF.exe
dist\Office2PDF\Office2PDFNativeWorker.exe
```

The build script also starts the worker without arguments and expects its usage exit code, proving that the companion executable launches.

## 7. Test the built release

From `dist\Office2PDF`, launch `Office2PDF.exe` and repeat one DOCX and one XLSX conversion with the default LibreOffice backend.

Then enable native mode and repeat only after the source native tests passed.

Confirm Task Manager shows no hidden Word, Excel, PowerPoint or LibreOffice process left behind after conversions finish.

## 8. Dad-machine acceptance

Copy the entire `dist\Office2PDF` folder to Dad's Windows 11 Home PC.

For the first acceptance run:

- leave native Office mode unchecked;
- use copies of two real documents;
- save PDFs beside the copies;
- open and inspect both PDFs;
- do not enable automatic deletion or source modification—Office2PDF never needs either.

Only enable native mode on Dad's PC after the same Word/Excel tests pass there.
