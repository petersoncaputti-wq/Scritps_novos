@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Executar-Kartado.ps1" %*
if errorlevel 1 pause
