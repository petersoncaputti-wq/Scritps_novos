@echo off
setlocal
cd /d "%~dp0"
".venv\Scripts\python.exe" main.py --browser-test %*
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
