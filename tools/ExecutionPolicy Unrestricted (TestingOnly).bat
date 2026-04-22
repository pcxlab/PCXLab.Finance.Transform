@echo off
title PowerShell Execution Policy Tool
color 0A

echo ============================================
echo   PowerShell Execution Policy Tool
echo ============================================
echo.

:: Check for Administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
set IS_ADMIN=0
echo [WARNING] Not running as Administrator.
echo Switching to CurrentUser scope...
echo.
) else (
set IS_ADMIN=1
echo [INFO] Running with Administrator privileges.
echo.
)

echo This is only for testing.
echo This message should appear when the script runs.
echo.

echo Applying PowerShell Execution Policy...
echo.

if %IS_ADMIN%==1 (
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy Unrestricted -Scope LocalMachine -Force"
) else (
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
)

if %errorLevel% equ 0 (
echo.
echo [SUCCESS] Execution Policy applied successfully.
) else (
echo.
echo [ERROR] Failed to apply Execution Policy.
)

echo.
echo ============================================
echo Script execution completed.
echo ============================================

pause
