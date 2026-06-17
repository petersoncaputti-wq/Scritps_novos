@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%06 - Criar Buscas Salvas em Projetos.ps1"

echo Executando criacao de buscas salvas no ProjectWise...
echo.

powershell.exe -NoProfile -NoExit -MTA -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

endlocal
