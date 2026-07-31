function New-EntradaDocumentosPanel {
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
        -PanelTitle '08 - Script Entrada de Documentos' `
        -PythonRelativePath 'Rotinas\08-EntradaDocumentos\Executar.EntradaDocumentos.ps1' `
        -LogPrefix 'PW_EntradaDocumentos' `
        -LogDirectory '08-EntradaDocumentos' `
        -Runtime PowerShell `
        -HelpText "Execução manual do fluxo automático para agilizar testes de chamados.`r`nEtapas: validar documentos recebidos e criar/garantir as pastas de disciplina." `
        -StartConfirmationText "Este fluxo altera estados de documentos e pode criar pastas no ProjectWise.`r`n`r`nDeseja executar agora as duas etapas?"
}
