@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%Criar_Buscas_Salvas_Projetos_PW.ps1"

echo Executando criacao de buscas salvas no ProjectWise...
echo.

powershell.exe -NoProfile -NoExit -MTA -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

endlocal
