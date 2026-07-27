[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptPrincipal = Join-Path -Path $PSScriptRoot -ChildPath 'GeraGRDsDocumentosComUnidade.ps1'

if (-not (Test-Path -LiteralPath $scriptPrincipal -PathType Leaf)) {
    throw "Script principal não encontrado: $scriptPrincipal"
}

Write-Host 'ATENÇÃO: iniciando execução REAL para todos os projetos.' -ForegroundColor Yellow
Write-Host "Script principal: $scriptPrincipal"

& $scriptPrincipal `
    -Executar `
    -ConfirmarExecucao 'CRIAR GRDS PARA TODOS OS PROJETOS'
