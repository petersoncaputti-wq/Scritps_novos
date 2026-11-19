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
        -Etapas @('Escolha a concessão e os projetos no ProjectWise.', 'Faça login no navegador Bentley quando ele abrir.', 'Informe o usuário (ou a planilha) e a ação: incluir, alterar ou excluir.', 'Confira o resumo e confirme.', 'Acompanhe a aplicação em cada projeto.') `
        -PythonRelativePath '..\ExecucaoManualUsuario\PWDM_Gerenciamento_Participantes_V2\gerenciar_participante_pwdm_connected_v2.py' `
        -LogPrefix 'PWDM_GerenciarParticipantes' `
        -LogDirectory '04-GerenciarParticipantesPWDM'
}
