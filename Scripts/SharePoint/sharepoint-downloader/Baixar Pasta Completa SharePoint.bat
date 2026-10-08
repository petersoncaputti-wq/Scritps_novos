@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Preparar-Downloader.ps1" %*
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
