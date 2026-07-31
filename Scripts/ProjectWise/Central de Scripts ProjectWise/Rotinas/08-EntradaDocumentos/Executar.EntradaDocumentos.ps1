[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$centralRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$projectWiseRoot = Split-Path -Parent $centralRoot
$validacao = Join-Path $projectWiseRoot 'ExecucaoAutomaticaServidor\ValidaDocumentos.ps1'
$criacaoPastas = Join-Path $projectWiseRoot 'ExecucaoAutomaticaServidor\CriaPastaDisciplina.ps1'

Write-Host 'Script Entrada de Documentos - execução manual'
Write-Host 'Este fluxo altera estados de documentos e pode criar pastas no ProjectWise.'
Write-Host ("Etapa 1: {0}" -f $validacao)
Write-Host ("Etapa 2: {0}" -f $criacaoPastas)

foreach ($arquivo in @($validacao, $criacaoPastas)) {
    if (-not (Test-Path -LiteralPath $arquivo -PathType Leaf)) {
        throw "Arquivo da rotina não encontrado: $arquivo"
    }
}

$inicio = Get-Date
try {
    Write-Host ''
    Write-Host '[1/2] Validando os documentos recebidos...'
    & $validacao
    Write-Host '[OK] Validação de documentos concluída.'

    Write-Host ''
    Write-Host '[2/2] Criando ou garantindo as pastas de disciplina...'
    & $criacaoPastas
    Write-Host '[OK] Criação de pastas concluída.'

    $duracao = (Get-Date) - $inicio
    Write-Host ''
    Write-Host ("[OK] Fluxo Entrada de Documentos finalizado em {0:hh\:mm\:ss}." -f $duracao)
}
catch {
    try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Error ("Falha no fluxo Entrada de Documentos: {0}" -f $_.Exception.Message)
    exit 1
}
