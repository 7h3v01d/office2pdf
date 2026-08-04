# Windows native-timeout hotfix

## Reported failure

On Windows, `test_native_timeout_is_a_real_process_deadline` measured approximately nine seconds for a one-second timeout.

## Root cause

The conversion deadline itself fired correctly. Cleanup then launched `taskkill` with its own ten-second timeout and performed an additional five-second `communicate()` wait. Slow startup or completion of `taskkill` therefore extended the public conversion deadline.

## Correction

- The dedicated Office process is opened with query and terminate rights.
- Its executable image is verified and terminated using the same Win32 handle.
- The isolated native worker is terminated directly with `Popen.kill()` on Windows.
- Pipe cleanup is capped at 0.5 seconds.
- Generic LibreOffice process-tree cleanup remains separate because LibreOffice can spawn a different process topology.

## Expected Windows test result without Microsoft Office installed

```text
Ran 27 tests
OK (skipped=9)
```

The timeout regression should complete in roughly one second and remain below the existing 2.5-second assertion.
