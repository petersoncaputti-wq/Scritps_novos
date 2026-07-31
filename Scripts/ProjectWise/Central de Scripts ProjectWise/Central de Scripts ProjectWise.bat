@echo off
setlocal

rem Este arquivo pode ficar em qualquer uma das Areas de Trabalho do usuario.
set "EXECUTOR_PS1="

rem 1. Executor ao lado deste inicializador.
if exist "%~dp0Executor.ps1" set "EXECUTOR_PS1=%~dp0Executor.ps1"

rem 2. Projeto abaixo da mesma Area de Trabalho.
if not defined EXECUTOR_PS1 if exist "%~dp0Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1" set "EXECUTOR_PS1=%~dp0Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1"

rem 3. Procura nas pastas sincronizadas do OneDrive corporativo e pessoal.
if not defined EXECUTOR_PS1 if defined OneDriveCommercial for /d %%D in ("%OneDriveCommercial%\*") do if exist "%%~fD\Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1" set "EXECUTOR_PS1=%%~fD\Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1"
if not defined EXECUTOR_PS1 if defined OneDriveConsumer for /d %%D in ("%OneDriveConsumer%\*") do if exist "%%~fD\Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1" set "EXECUTOR_PS1=%%~fD\Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise\Executor.ps1"

if not exist "%EXECUTOR_PS1%" (
    echo Nao foi possivel localizar a Central de Scripts ProjectWise.
    echo.
    echo O inicializador procurou nas Areas de Trabalho pessoal e corporativa.
    echo Verifique se a pasta Scritps_novos continua dentro de uma delas.
    pause
    exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%EXECUTOR_PS1%"

if errorlevel 1 pause
endlocal
