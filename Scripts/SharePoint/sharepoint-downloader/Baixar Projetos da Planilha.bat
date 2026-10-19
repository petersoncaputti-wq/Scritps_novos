@echo off
set /p "PROJECT_LIST=Caminho completo da planilha XLSX: "
set "PROJECT_LIST=%PROJECT_LIST:"=%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Preparar-Downloader.ps1" -ProjectList "%PROJECT_LIST%"
pause
