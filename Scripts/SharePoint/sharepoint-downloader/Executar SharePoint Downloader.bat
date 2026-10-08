@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Configure o ambiente virtual seguindo o README.md.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" main.py %*
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
