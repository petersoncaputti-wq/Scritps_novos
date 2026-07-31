function New-GerenciarParticipantesPWDMPanel {
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
        -PanelTitle '04 - Gerenciar participantes PWDM' `
        -PythonRelativePath '..\ExecucaoManualUsuario\PWDM_Gerenciamento_Participantes_V2\gerenciar_participante_pwdm_connected_v2.py' `
        -LogPrefix 'PWDM_GerenciarParticipantes' `
        -LogDirectory '04-GerenciarParticipantesPWDM'
}
