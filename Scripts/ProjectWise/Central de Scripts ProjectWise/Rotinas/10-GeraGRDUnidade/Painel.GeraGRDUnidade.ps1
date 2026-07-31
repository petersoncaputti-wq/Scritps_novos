function New-GeraGRDUnidadePanel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack
    )

    $sharedPanel = Join-Path $CentralRoot 'Rotinas\03-IncluirUsuariosWeb\Painel.IncluirUsuariosWeb.ps1'
    if (-not (Test-Path -LiteralPath $sharedPanel -PathType Leaf)) {
        throw "Controlador integrado não encontrado: $sharedPanel"
    }
    . $sharedPanel

    New-IncluirUsuariosWebPanel `
        -Parent $Parent `
        -CentralRoot $CentralRoot `
        -OnBack $OnBack `
        -PanelTitle '10 - Gera GRD para Unidade' `
        -PythonRelativePath 'Rotinas\10-GeraGRDUnidade\Executar.GeraGRDUnidade.ps1' `
        -LogPrefix 'PW_GeraGRDUnidade' `
        -LogDirectory '10-GeraGRDUnidade' `
        -Runtime PowerShell `
        -HelpText "Execução manual da rotina semanal do servidor, normalmente agendada para segunda-feira às 08:00.`r`nA saída e o resumo permanecerão visíveis nesta tela." `
        -StartConfirmationText "Esta rotina executará o processamento real: poderá criar pastas de GRD e importar planilhas no ProjectWise.`r`n`r`nDeseja iniciar a geração das GRDs para as unidades?" `
        -SecondaryButtonText 'Simular' `
        -SecondaryRuntimeArguments '-Simular'
}
