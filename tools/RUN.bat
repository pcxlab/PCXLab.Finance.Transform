@echo off
setlocal

Title Running Automation
Color 0b

set INPUT_FOLDER=C:\TEST

if not exist "%INPUT_FOLDER%" (
    echo.
    echo Default folder not found: %INPUT_FOLDER%
    set /p INPUT_FOLDER=Enter input folder:
)

REM Get script directory
set SCRIPT_DIR=%~dp0

REM Define PowerShell script path
set PS_SCRIPT=%SCRIPT_DIR%main.ps1

echo =====================================
echo PCXLab Automation Tool
echo =====================================

REM Check script exists
if not exist "%PS_SCRIPT%" (
    echo ERROR: main.ps1 not found!
    echo Expected path: %PS_SCRIPT%
    pause
    exit /b
)

REM Check PowerShell
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: PowerShell not found.
    pause
    exit /b
)

echo Running script...
echo.

powershell.exe -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Folder "%INPUT_FOLDER%"

if %errorlevel% neq 0 (
    echo Script execution failed. Check logs.
) else (
    echo Script completed successfully.
)

pause