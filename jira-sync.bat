@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0jira-sync.ps1"
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: script failed with exit code %ERRORLEVEL%
    pause
)
