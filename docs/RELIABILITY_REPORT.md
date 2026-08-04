# Office2PDF reliability report

**Pass date:** 2026-08-04  
**Target:** dependable DOCX/XLSX-to-PDF desktop utility for Windows 11 Home

## Current status

The LibreOffice conversion path is suitable for Windows build and target-PC acceptance testing. The optional Microsoft Office path has been structurally hardened but remains opt-in until it passes real Windows Word and Excel tests.

## Proven controls

1. **Staged output:** every attempt writes into a private directory under the output filesystem.
2. **Pre-commit validation:** minimum size, `%PDF-` header and `%%EOF` marker are required.
3. **Atomic commit:** `os.replace()` occurs only after staged validation.
4. **Stale-output resistance:** an older destination cannot create false success.
5. **Overwrite preservation:** a failed attempt does not alter the previous good PDF.
6. **Failure containment:** process, filesystem and unexpected worker errors return structured results.
7. **LibreOffice isolation:** each attempt uses a separate user profile.
8. **LibreOffice timeout cleanup:** the launched process tree is terminated before retry.
9. **Deterministic discovery:** extension checks are case-insensitive; results are sorted and deduplicated.
10. **Input hygiene:** Office lock files and hidden dot-files are excluded.
11. **Collision refusal:** sources mapping to one destination PDF are rejected before work starts.
12. **Cancellation safety:** unstarted jobs are not submitted after cancellation; active jobs finish safely.
13. **Window-close safety:** the GUI waits for active conversion work before destruction.

## Native Office hardening

The incoming update originally used a timed Python thread and image-wide `taskkill`. Both were release blockers.

The replacement uses a companion process:

- `Office2PDFNativeWorker` owns COM automation;
- `DispatchEx` creates a dedicated Office instance;
- the worker records the owning Office PID;
- timeout is enforced by the parent process;
- PID cleanup verifies the executable image and terminates it through the same Win32 handle;
- native timeout cleanup launches no `taskkill` subprocess and adds at most a tightly bounded reap;
- no `/IM WINWORD.EXE`, `/IM EXCEL.EXE` or `/IM POWERPNT.EXE` command remains;
- macros are force-disabled before opening source files;
- native failure falls back to LibreOffice when possible;
- native mode is off by default until platform validation.

## Automated verification

```text
27 tests total
19 passed, 8 platform-specific skipped, 0 failed (local Linux run)
18 passed, 9 Office/platform-specific skipped, 0 failed (expected Windows run without Office)
```

Coverage includes:

- safe overwrite and stale destination regression;
- invalid existing PDF handling;
- launch-error containment;
- deterministic discovery and lock-file exclusion;
- destination collision detection;
- CLI range validation;
- actual LibreOffice conversion, overwrite, skip and failure preservation;
- true native-worker timeout behaviour, including the Windows cleanup-delay regression;
- native failure-to-LibreOffice fallback;
- native mode being opt-in;
- real Word export tests when Word is available;
- real Excel export tests when Excel is available.

## Real conversion verification

A generated village newsletter DOCX and treasurer workbook XLSX were converted concurrently through the actual LibreOffice binary.

```text
newsletter.pdf — valid, 1 page, 33,536 bytes
treasurer.pdf  — valid, 1 page, 28,309 bytes
```

Both opened with a PDF parser and rendered correctly for visual inspection.

## Honest remaining boundary

This environment cannot execute Microsoft COM automation or produce a trustworthy Windows PyInstaller binary. Therefore:

- the isolated native worker is reasoned, unit-tested around its process boundary, and protected by Windows-only end-to-end tests, but those tests must run on a real Windows + Office machine;
- the final `Office2PDF.exe` and `Office2PDFNativeWorker.exe` must be built and tested on Windows;
- native mode must remain unchecked for Dad until `WINDOWS_VALIDATION_CHECKLIST.md` passes.

## Release recommendation

Proceed to a Windows one-folder build and LibreOffice-only acceptance test. Do not yet call the native backend production-proven. Once Word and Excel tests run successfully on the target environment with no orphan processes, native mode can be promoted from experimental.
