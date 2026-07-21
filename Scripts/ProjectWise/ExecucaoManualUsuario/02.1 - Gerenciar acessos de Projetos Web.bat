@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%13 - Incluir Usuarios em Projetos PW Web por Planilha.ps1"

if not exist "%SCRIPT_PATH%" (
    echo Script PowerShell nao encontrado:
    echo %SCRIPT_PATH%
    pause
    exit /b 1
)

echo Executando gerenciamento de acessos de projetos PW Web...
echo.

powershell.exe -NoProfile -NoExit -MTA -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

endlocal
