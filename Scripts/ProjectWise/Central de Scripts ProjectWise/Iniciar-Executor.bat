@echo off
setlocal
set "EXECUTOR_DIR=%~dp0"
set "EXECUTOR_PS1=%EXECUTOR_DIR%Executor.ps1"

if not exist "%EXECUTOR_PS1%" (
    echo Executor nao encontrado:
    echo %EXECUTOR_PS1%
    pause
    exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%EXECUTOR_PS1%"

if errorlevel 1 pause
endlocal
