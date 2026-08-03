@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Run setup.bat first.
    pause
    exit /b 1
)

echo Installing build tools...
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

echo Running regression tests...
.venv\Scripts\python.exe -m py_compile office2pdf.py gui.py
if errorlevel 1 goto :failed
.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 goto :failed

echo Building reliable one-folder Windows release...
.venv\Scripts\python.exe -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --onedir ^
    --name Office2PDF ^
    --paths "%CD%" ^
    gui.py
if errorlevel 1 goto :failed

echo.
echo BUILD COMPLETE:
echo   %CD%\dist\Office2PDF\Office2PDF.exe
echo.
echo Copy the entire dist\Office2PDF folder, not only the EXE.
pause
exit /b 0

:failed
echo.
echo BUILD FAILED. Review the message above.
pause
exit /b 1
