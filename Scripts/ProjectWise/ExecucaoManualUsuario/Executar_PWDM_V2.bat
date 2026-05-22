@echo off
cd /d "%~dp0PWDM_Gerenciamento_Participantes_V2"

if not exist "gerenciar_participante_pwdm_connected_v2.py" (
    echo Arquivo Python nao encontrado: %CD%\gerenciar_participante_pwdm_connected_v2.py
    pause
    exit /b 1
)

python gerenciar_participante_pwdm_connected_v2.py
pause
