@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Build Office2PDF 1.0.0

for %%F in (
    "assets\office2pdf.ico"
    "LICENSE.txt"
    "THIRD_PARTY_NOTICES.txt"
    "SOURCE_OFFER.txt"
    "windows_version_info.txt"
    "windows_worker_version_info.txt"
) do (
    if not exist "%%~F" (
        echo ERROR: Required release file is missing: %%~F
        pause
        exit /b 1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Run setup.bat first.
    pause
    exit /b 1
)

echo Installing build tools...
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

echo Running regression tests...
.venv\Scripts\python.exe -m py_compile office2pdf.py gui.py native_office_worker.py version_info.py package_source.py
if errorlevel 1 goto :failed
.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 goto :failed

echo Building isolated Microsoft Office worker...
.venv\Scripts\python.exe -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --console ^
    --name Office2PDFNativeWorker ^
    --icon "%CD%\assets\office2pdf.ico" ^
    --version-file "%CD%\windows_worker_version_info.txt" ^
    --paths "%CD%" ^
    native_office_worker.py
if errorlevel 1 goto :failed

echo Building reliable one-folder Windows release...
.venv\Scripts\python.exe -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --onedir ^
    --name Office2PDF ^
    --icon "%CD%\assets\office2pdf.ico" ^
    --version-file "%CD%\windows_version_info.txt" ^
    --add-data "%CD%\assets\office2pdf.ico;assets" ^
    --add-data "%CD%\LICENSE.txt;." ^
    --add-data "%CD%\THIRD_PARTY_NOTICES.txt;." ^
    --add-data "%CD%\SOURCE_OFFER.txt;." ^
    --paths "%CD%" ^
    gui.py
if errorlevel 1 goto :failed

copy /Y "dist\Office2PDFNativeWorker.exe" "dist\Office2PDF\Office2PDFNativeWorker.exe" >nul
if errorlevel 1 goto :failed
copy /Y "create_shortcuts.bat" "dist\Office2PDF\Create Office2PDF Shortcuts.bat" >nul
if errorlevel 1 goto :failed
copy /Y "LICENSE.txt" "dist\Office2PDF\LICENSE.txt" >nul
if errorlevel 1 goto :failed
copy /Y "THIRD_PARTY_NOTICES.txt" "dist\Office2PDF\THIRD_PARTY_NOTICES.txt" >nul
if errorlevel 1 goto :failed
copy /Y "SOURCE_OFFER.txt" "dist\Office2PDF\SOURCE_OFFER.txt" >nul
if errorlevel 1 goto :failed

 echo Creating corresponding-source archive...
.venv\Scripts\python.exe package_source.py --output "dist\Office2PDF\Office2PDF-1.0.0-source.zip"
if errorlevel 1 goto :failed

if not exist "dist\Office2PDF\Office2PDF.exe" goto :failed
if not exist "dist\Office2PDF\_internal\python311.dll" goto :failed
if not exist "dist\Office2PDF\Office2PDFNativeWorker.exe" goto :failed
if not exist "dist\Office2PDF\Create Office2PDF Shortcuts.bat" goto :failed
if not exist "dist\Office2PDF\LICENSE.txt" goto :failed
if not exist "dist\Office2PDF\THIRD_PARTY_NOTICES.txt" goto :failed
if not exist "dist\Office2PDF\Office2PDF-1.0.0-source.zip" goto :failed

rem The helper should start and return usage error 2 when called without args.
"dist\Office2PDF\Office2PDFNativeWorker.exe" >nul 2>nul
if not "%errorlevel%"=="2" goto :failed

echo.
echo BUILD COMPLETE:
echo   %CD%\dist\Office2PDF\Office2PDF.exe
echo.
echo This release includes:
echo   - five-second branded splash screen
 echo   - About tab with version and licence information
 echo   - Windows file-version metadata
 echo   - licence and third-party notices
 echo   - GPL corresponding-source archive
 echo   - validated no-PowerShell shortcut creator
 echo.
echo Copy the entire dist\Office2PDF folder, not only the EXE.
echo After placing the folder permanently, run:
echo   Create Office2PDF Shortcuts.bat
pause
exit /b 0

:failed
echo.
echo BUILD FAILED. Review the message above.
pause
exit /b 1
