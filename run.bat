@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo Office2PDF is not set up yet. Running setup...
    call setup.bat
    if errorlevel 1 exit /b 1
)

start "Office2PDF" ".venv\Scripts\pythonw.exe" "gui.py"
exit /b 0
