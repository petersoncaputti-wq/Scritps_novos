@echo off
rem Abre a tela para escolher a pasta. Tambem aceita arrastar uma pasta sobre este BAT.
rem A pasta escolhida e mantida; somente o conteudo dentro dela e excluido.
start "" powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Excluir-Permanente.ps1" %*
