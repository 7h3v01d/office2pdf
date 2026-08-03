@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python Launcher ^(py.exe^) was not found.
    echo Install Python 3.11 or newer from python.org and enable the launcher.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 goto :failed
)

echo Installing application requirements...
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto :failed
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :failed

echo.
echo Setup complete. Run run.bat to start Office2PDF.
pause
exit /b 0

:failed
echo.
echo SETUP FAILED. Review the message above.
pause
exit /b 1
