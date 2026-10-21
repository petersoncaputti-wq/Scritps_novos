#requires -Version 5.1
$ErrorActionPreference='Stop'
try {
    . (Join-Path $PSScriptRoot 'Interface.ps1')
    $path=SelectProjectList
    if(!$path){Write-Host 'Selecao da planilha cancelada.';return}
    Write-Host "Planilha selecionada: $path"
    & (Join-Path $PSScriptRoot 'Download-SharePoint.ps1') -ProjectList $path
}catch{Write-Host "Nao foi possivel iniciar a fila: $($_.Exception.Message)" -ForegroundColor Red}
