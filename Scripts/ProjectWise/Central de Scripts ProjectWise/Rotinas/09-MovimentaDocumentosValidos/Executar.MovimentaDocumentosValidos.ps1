[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$centralRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$projectWiseRoot = Split-Path -Parent $centralRoot
$scriptOperacional = Join-Path $projectWiseRoot 'ExecucaoAutomaticaServidor\MovimentaDocumentosValidos_v3.ps1'

Write-Host 'Movimenta Documentos Válidos - execução manual'
Write-Host 'Este fluxo copia documentos para as pastas de disciplina, trata versões anteriores e altera estados no ProjectWise.'
Write-Host ("Código operacional: {0}" -f $scriptOperacional)

if (-not (Test-Path -LiteralPath $scriptOperacional -PathType Leaf)) {
    throw "Script operacional não encontrado: $scriptOperacional"
}

$inicio = Get-Date
try {
    Write-Host ''
    Write-Host '[1/1] Processando os documentos válidos...'
    & $scriptOperacional
    $duracao = (Get-Date) - $inicio
    Write-Host ''
    Write-Host ("[OK] Fluxo Movimenta Documentos Válidos finalizado em {0:hh\:mm\:ss}." -f $duracao)
}
catch {
    try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Error ("Falha no fluxo Movimenta Documentos Válidos: {0}" -f $_.Exception.Message)
    exit 1
}
