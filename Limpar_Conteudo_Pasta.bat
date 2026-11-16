@echo off
if not exist "%~dp0Limpar_Conteudo_Pasta.ps1" (
    echo Nao foi encontrado o arquivo Limpar_Conteudo_Pasta.ps1 nesta pasta.
    echo.
    echo Se houver um arquivo .ps1.txt com aviso de quarentena, solicite
    echo a analise e liberacao do script pela TI ou seguranca corporativa.
    echo.
    pause
    exit /b 1
)
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Limpar_Conteudo_Pasta.ps1"
if errorlevel 1 (
    echo.
    echo O script terminou com erro. Confira a mensagem acima.
    pause
    exit /b 1
)
