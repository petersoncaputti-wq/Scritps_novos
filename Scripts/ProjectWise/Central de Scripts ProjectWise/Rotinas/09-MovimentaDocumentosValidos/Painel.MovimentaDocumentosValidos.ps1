function New-MovimentaDocumentosValidosPanel {
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
        -PanelTitle '09 - Movimenta Documentos Válidos' `
        -PythonRelativePath 'Rotinas\09-MovimentaDocumentosValidos\Executar.MovimentaDocumentosValidos.ps1' `
        -LogPrefix 'PW_MovimentaDocumentosValidos' `
        -LogDirectory '09-MovimentaDocumentosValidos' `
        -Runtime PowerShell `
        -HelpText "Execução manual da rotina automática do servidor para agilizar testes de chamados.`r`nA saída e o andamento permanecerão visíveis nesta tela." `
        -StartConfirmationText "Esta rotina copia documentos para as pastas de disciplina, pode mover versões anteriores para Superados e altera estados no ProjectWise.`r`n`r`nDeseja iniciar a execução manual?"
}
