@echo off
setlocal EnableExtensions
title Repair Office2PDF Shortcuts

set "SCRIPT_DIR=%~dp0"
set "APP_EXE="
set "APP_DIR="

rem Preferred: script is run from the project root.
if exist "%SCRIPT_DIR%dist\Office2PDF\Office2PDF.exe" (
    if exist "%SCRIPT_DIR%dist\Office2PDF\_internal\python311.dll" (
        set "APP_EXE=%SCRIPT_DIR%dist\Office2PDF\Office2PDF.exe"
        set "APP_DIR=%SCRIPT_DIR%dist\Office2PDF"
    )
)

rem Alternative: script is placed inside the finished dist\Office2PDF folder.
if not defined APP_EXE (
    if exist "%SCRIPT_DIR%Office2PDF.exe" (
        if exist "%SCRIPT_DIR%_internal\python311.dll" (
            set "APP_EXE=%SCRIPT_DIR%Office2PDF.exe"
            set "APP_DIR=%SCRIPT_DIR:~0,-1%"
        )
    )
)

if not defined APP_EXE (
    echo.
    echo ERROR: A complete Office2PDF release folder was not found.
    echo.
    echo This repair script must be either:
    echo   1. In the project root containing dist\Office2PDF
    echo      OR
    echo   2. Inside the finished dist\Office2PDF folder.
    echo.
    echo The valid release must contain:
    echo   Office2PDF.exe
    echo   _internal\python311.dll
    echo.
    echo Do NOT run or shortcut:
    echo   build\Office2PDF\Office2PDF.exe
    echo.
    pause
    exit /b 1
)

set "CSCRIPT=%SystemRoot%\System32\cscript.exe"
if not exist "%CSCRIPT%" (
    echo.
    echo ERROR: Windows Script Host was not found:
    echo   %CSCRIPT%
    echo.
    pause
    exit /b 2
)

set "VBS_FILE=%TEMP%\Office2PDF_RepairShortcuts_%RANDOM%_%RANDOM%.vbs"

> "%VBS_FILE%" echo Option Explicit
>>"%VBS_FILE%" echo Dim shell, fso, args, exePath, workDir, desktopPath, programsPath
>>"%VBS_FILE%" echo Dim desktopLink, startMenuLink
>>"%VBS_FILE%" echo Set shell = CreateObject("WScript.Shell")
>>"%VBS_FILE%" echo Set fso = CreateObject("Scripting.FileSystemObject")
>>"%VBS_FILE%" echo Set args = WScript.Arguments
>>"%VBS_FILE%" echo If args.Count ^< 1 Then WScript.Quit 2
>>"%VBS_FILE%" echo exePath = fso.GetAbsolutePathName(args(0))
>>"%VBS_FILE%" echo workDir = fso.GetParentFolderName(exePath)
>>"%VBS_FILE%" echo desktopPath = shell.SpecialFolders("Desktop")
>>"%VBS_FILE%" echo programsPath = shell.SpecialFolders("Programs")
>>"%VBS_FILE%" echo desktopLink = fso.BuildPath(desktopPath, "Office2PDF.lnk")
>>"%VBS_FILE%" echo startMenuLink = fso.BuildPath(programsPath, "Office2PDF.lnk")
>>"%VBS_FILE%" echo CreateLink desktopLink
>>"%VBS_FILE%" echo CreateLink startMenuLink
>>"%VBS_FILE%" echo WScript.Echo "Desktop shortcut updated:" ^& vbCrLf ^& desktopLink
>>"%VBS_FILE%" echo WScript.Echo "Start Menu shortcut updated:" ^& vbCrLf ^& startMenuLink
>>"%VBS_FILE%" echo WScript.Quit 0
>>"%VBS_FILE%" echo Sub CreateLink(linkPath)
>>"%VBS_FILE%" echo     Dim shortcut
>>"%VBS_FILE%" echo     Set shortcut = shell.CreateShortcut(linkPath)
>>"%VBS_FILE%" echo     shortcut.TargetPath = exePath
>>"%VBS_FILE%" echo     shortcut.WorkingDirectory = workDir
>>"%VBS_FILE%" echo     shortcut.IconLocation = exePath ^& ",0"
>>"%VBS_FILE%" echo     shortcut.Description = "Convert Office documents to PDF"
>>"%VBS_FILE%" echo     shortcut.Save
>>"%VBS_FILE%" echo End Sub

echo.
echo Validated release executable:
echo   %APP_EXE%
echo.
echo Recreating Office2PDF shortcuts...
"%CSCRIPT%" //nologo "%VBS_FILE%" "%APP_EXE%"
set "RESULT=%ERRORLEVEL%"

del /q "%VBS_FILE%" >nul 2>&1

if not "%RESULT%"=="0" (
    echo.
    echo ERROR: Windows could not repair the shortcuts.
    echo Windows Script Host returned exit code %RESULT%.
    echo.
    pause
    exit /b %RESULT%
)

echo.
echo SUCCESS: Shortcuts now point to the complete release:
echo   %APP_EXE%
echo.
echo You may now launch Office2PDF from the Desktop or Start Menu.
echo.
pause
exit /b 0
