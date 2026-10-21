@echo off
setlocal
cd /d "%~dp0"
powershell.exe -STA -NoLogo -NoProfile -NoExit -File "%~dp0Fila-Projetos.ps1"
pause
