@echo off
setlocal
cd /d "%~dp0"
echo Iniciando downloader SharePoint em PowerShell...
echo A janela permanecera aberta para mostrar qualquer erro.
echo.
powershell.exe -NoLogo -NoProfile -NoExit -File "%~dp0Download-SharePoint.ps1" %*
echo.
echo O PowerShell foi encerrado. Pressione uma tecla para fechar esta janela.
pause >nul
