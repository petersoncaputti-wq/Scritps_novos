function New-CriarProjetosPanel {
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
        -PanelTitle '06 - Criar projetos no ProjectWise' `
        -PythonRelativePath '..\ExecucaoManualUsuario\01 - Criar Projetos no ProjectWise.ps1' `
        -LogPrefix 'PW_CriarProjetos' `
        -LogDirectory '06-CriarProjetos' `
        -Runtime 'PowerShell' `
        -HelpText "Ao iniciar, faça o login no ProjectWise e selecione a planilha de projetos.`r`nO andamento e o resultado de cada projeto serão exibidos nesta tela."
}
