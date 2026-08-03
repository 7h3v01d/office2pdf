@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Run setup.bat first.
    pause
    exit /b 1
)

.venv\Scripts\python.exe -m py_compile office2pdf.py gui.py
if errorlevel 1 goto :failed
.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 goto :failed

echo.
echo All tests passed.
pause
exit /b 0

:failed
echo.
echo TESTS FAILED.
pause
exit /b 1
