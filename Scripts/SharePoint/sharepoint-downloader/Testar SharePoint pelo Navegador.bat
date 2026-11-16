@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Preparar-Downloader.ps1" -PrepareOnly
if errorlevel 1 (
    pause
    exit /b 1
)
".venv\Scripts\python.exe" main.py --browser-test %*
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
