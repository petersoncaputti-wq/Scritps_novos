function New-CriarBuscasSalvasPanel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack
    )

    $sharedPanel = Join-Path $CentralRoot 'Rotinas\03-IncluirUsuariosWeb\Painel.IncluirUsuariosWeb.ps1'
    if (-not (Test-Path -LiteralPath $sharedPanel -PathType Leaf)) {
        throw "Controlador visual compartilhado não encontrado: $sharedPanel"
    }
    . $sharedPanel

    New-IncluirUsuariosWebPanel `
        -Parent $Parent `
        -CentralRoot $CentralRoot `
        -OnBack $OnBack `
        -PanelTitle '07 - Criar buscas salvas' `
        -PythonRelativePath '..\ExecucaoManualUsuario\06 - Criar Buscas Salvas em Projetos.ps1' `
        -LogPrefix 'PW_CriarBuscasSalvas' `
        -LogDirectory '07-CriarBuscasSalvas' `
        -Runtime 'PowerShell' `
        -HelpText "Selecione as buscas, a concessão e os projetos nas janelas apresentadas.`r`nBuscas existentes serão ignoradas e todo o andamento permanecerá visível nesta tela."
}
