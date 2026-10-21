@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -NoExit -File "%~dp0Download-SharePoint.ps1" -BrowserTest %*
pause
