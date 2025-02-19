@echo off
setlocal enabledelayedexpansion

REM Check if PowerShell 7 is installed
where pwsh.exe >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo PowerShell 7 is not installed or not in PATH
    echo Please install PowerShell 7 from: https://aka.ms/powershell-release?tag=stable
    pause
    exit /b 1
)

REM Get the directory of the current batch file
set "SCRIPT_DIR=%~dp0"

REM Check if the PowerShell script exists
if not exist "%SCRIPT_DIR%SharePointVersionControl.ps1" (
    echo Error: SharePointVersionControl.ps1 not found in the current directory
    echo Expected location: %SCRIPT_DIR%SharePointVersionControl.ps1
    pause
    exit /b 1
)

REM Run the PowerShell script with PowerShell 7
echo Running SharePointVersionControl.ps1 with PowerShell 7...
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SharePointVersionControl.ps1"

REM Check if the script executed successfully
if %ERRORLEVEL% neq 0 (
    echo Error: Script execution failed with error code %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)

echo Script execution completed successfully
pause
exit /b 0