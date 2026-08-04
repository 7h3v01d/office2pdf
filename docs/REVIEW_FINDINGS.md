# Review findings — incoming 2026-08-04 update

## Verdict

The update made a strong architectural move by adding native Word/Excel/PowerPoint export and real end-to-end LibreOffice tests. It also preserved the earlier staged/atomic overwrite fix.

It nevertheless introduced two release-blocking hazards in the native timeout implementation.

## P0 — timeout was not a real timeout

The native converter ran inside a `ThreadPoolExecutor` context manager. `future.result(timeout=...)` raised at the deadline, but leaving the context manager called `shutdown(wait=True)`, which waited for the same stuck thread.

A controlled reproduction used a two-second worker with a 0.1-second timeout. The function raised `TimeoutError` only after approximately 2.0 seconds.

### Resolution

Native COM automation now runs in a separate companion process. The parent uses `Popen.communicate(timeout=...)`, so it has a genuine process boundary that can be terminated.

## P0 — timeout cleanup could destroy unrelated unsaved work

The update used:

```text
taskkill /IM WINWORD.EXE /F
taskkill /IM EXCEL.EXE /F
taskkill /IM POWERPNT.EXE /F
```

That targets every matching process, including an unrelated Word or Excel session containing unsaved user work.

### Resolution

The companion worker creates a dedicated Office instance with `DispatchEx`, obtains its window-owned PID, and records that PID for the parent. Timeout cleanup:

1. reads only that recorded PID;
2. queries the current executable image through Win32;
3. confirms it matches the expected Office executable;
4. terminates that process through the same verified Win32 handle;
5. directly terminates the isolated helper and performs only a 0.5-second bounded reap.

No image-wide Office termination remains, and the native timeout path launches no `taskkill` subprocess. This also fixes the later Windows regression where a one-second timeout took approximately nine seconds during cleanup.

## P1 — native path was enabled before target-platform proof

The new backend was on by default despite the documentation correctly stating that it had never run against real Windows Office in the development environment.

### Resolution

LibreOffice is again the default. Native Office is an explicit GUI/CLI opt-in until the Windows validation checklist passes.

## P1 — no conversion fallback after a native export failure

Backend selection fell back to LibreOffice when Office was absent, but not when Office was detected and COM export then failed.

### Resolution

After native attempts fail, a valid LibreOffice installation receives the same staged conversion request. Results identify the backend that ultimately produced the PDF and retain the native failure context in the message.

## P1 — native automation opened legacy Office formats without an explicit macro gate

Legacy `.doc`, `.xls` and `.ppt` files can contain macros. Hidden automation must not depend solely on local Trust Center defaults.

### Resolution

The worker requires `AutomationSecurity = 3` before opening a source document. If macro disabling cannot be established, native conversion aborts and can fall back to LibreOffice.

## P2 — Windows native tests covered Word but not the required XLSX path

The project goal explicitly includes Excel workbooks, but native end-to-end coverage existed only for Word.

### Resolution

Real Windows-only Excel tests now cover backend selection, validated export and safe overwrite. They run automatically when Excel is installed and skip cleanly elsewhere.

## Outcome

The incoming update was directionally correct, but the original native timeout path was not safe enough for Dad's machine. The replacement keeps the fidelity benefit while containing COM in a companion process and retaining LibreOffice as the proven default.
