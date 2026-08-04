# Native Word headless-window hotfix

**Date:** 2026-08-04  
**Trigger:** Windows build regression gate failed three real Word COM tests before PyInstaller packaging began.

## Observed failure

Word was installed and detected, but the isolated worker exited with:

```text
RuntimeError: Office did not expose a usable application window handle
```

Excel native export passed on the same machine, confirming that icon integration, pywin32 installation, the worker executable design and the general native process boundary were not the cause.

## Root cause

The worker assumed every Office `Application` object exposed `Hwnd`. Excel did on the target machine; Word did not expose a usable application-level handle while running invisibly. Word's object model exposes the reliable handle on the document `Window` object (`Document.ActiveWindow.Hwnd`).

## Correction

The Word path now:

1. creates a dedicated `Word.Application` with `DispatchEx`;
2. force-disables macros before opening the source;
3. opens the document read-only and outside the recent-files list;
4. records the owning PID from `Document.ActiveWindow.Hwnd`;
5. if the hidden window has not published a handle, briefly minimises and reveals only the dedicated automation instance for a maximum of two seconds, captures the handle, then hides it immediately;
6. exports to the staged PDF and quits Word normally.

Timeout cleanup remains exact-PID and executable-image verified. The patch does not use `taskkill /IM`, enumerate or terminate unrelated Word sessions, or relax the hard process deadline.

## Verification

Local platform-neutral suite:

```text
Ran 31 tests
OK (skipped=8)
```

Two new tests prove:

- Word succeeds when only the document window exposes `Hwnd`;
- the visibility probe restores the original hidden state and records the exact window/PID path.

The definitive remaining check is rerunning `build_exe.bat` on the same Windows + Word machine. Expected result: all three real Word tests pass, followed by the PyInstaller build.
