[CmdletBinding()]
param([switch]$Simular)

$ErrorActionPreference = 'Stop'
$centralRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$projectWiseRoot = Split-Path -Parent $centralRoot
$scriptOperacional = Join-Path $projectWiseRoot 'ExecucaoAutomaticaServidor\GeraGRDsDocumentosComUnidade.ps1'
$modeloExcel = Join-Path $env:ProgramData 'Ecorodovias\GRD\Modelos\ModeloGRD.xlsx'
$logOperacional = Join-Path $centralRoot 'Logs\10-GeraGRDUnidade\Operacional'

Write-Host ("Gera GRD para Unidade - {0}" -f $(if ($Simular) { 'simulação' } else { 'execução real' }))
Write-Host 'Este fluxo cria a estrutura da GRD e importa a planilha de documentos no ProjectWise.'
Write-Host ("Código operacional: {0}" -f $scriptOperacional)
Write-Host ("Modelo Excel: {0}" -f $modeloExcel)

if (-not (Test-Path -LiteralPath $scriptOperacional -PathType Leaf)) {
    throw "Script operacional não encontrado: $scriptOperacional"
}
if (-not (Test-Path -LiteralPath $modeloExcel -PathType Leaf)) {
    throw "Modelo Excel não encontrado: $modeloExcel"
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "O módulo PowerShell 'ImportExcel' não está instalado."
}
if (-not (Test-Path -LiteralPath $logOperacional)) {
    New-Item -ItemType Directory -Path $logOperacional -Force | Out-Null
}

$inicio = Get-Date
try {
    Write-Host ''
    Write-Host $(if ($Simular) { '[1/1] Simulando a geração das GRDs para as unidades...' } else { '[1/1] Gerando as GRDs para as unidades...' })
    if ($Simular) {
        & $scriptOperacional -Simular -CaminhoModeloExcel $modeloExcel -LogDirectory $logOperacional
    }
    else {
        & $scriptOperacional -Executar -CaminhoModeloExcel $modeloExcel -LogDirectory $logOperacional
    }

    $duracao = (Get-Date) - $inicio
    Write-Host ''
    $descricao = if ($Simular) { 'Simulação do fluxo Gera GRD para Unidade' } else { 'Fluxo Gera GRD para Unidade' }
    Write-Host ("[OK] {0} finalizada em {1:hh\:mm\:ss}." -f $descricao, $duracao)
}
catch {
    try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Error ("Falha no fluxo Gera GRD para Unidade: {0}" -f $_.Exception.Message)
    exit 1
}
