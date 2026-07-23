@echo off
cd /d "%~dp0PWDM_Gerenciamento_Participantes_V2"

if not exist "diagnostico_exclusao_participante_v2.py" (
    echo Arquivo Python nao encontrado: %CD%\diagnostico_exclusao_participante_v2.py
    pause
    exit /b 1
)

python diagnostico_exclusao_participante_v2.py
pause
