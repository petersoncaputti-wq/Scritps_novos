function Open-GerenciarUsuariosInativosPanel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack
    )

    $legacyScript = [IO.Path]::GetFullPath((Join-Path $CentralRoot '..\ExecucaoManualUsuario\04 - Gerenciar Usuarios Inativos.ps1'))
    if (-not (Test-Path -LiteralPath $legacyScript -PathType Leaf)) {
        throw "Script original não encontrado: $legacyScript"
    }

    $source = Get-Content -LiteralPath $legacyScript -Raw
    $endPattern = '\[void\]\$form\.ShowDialog\(\)\s*$'
    if (-not [regex]::IsMatch($source, $endPattern)) {
        throw 'A estrutura final do script original mudou e não pôde ser integrada.'
    }

    $closeOriginal = @'
$btnFechar.Text = "Fechar"
$btnFechar.Location = New-Object System.Drawing.Point(1055, 18)
$btnFechar.Size = New-Object System.Drawing.Size(90, 28)
$btnFechar.Anchor = "Top,Right"
$btnFechar.Add_Click({ $form.Close() })
'@
    $closeIntegrated = @'
$btnFechar.Text = "Voltar à Central"
$btnFechar.Location = New-Object System.Drawing.Point(1015, 18)
$btnFechar.Size = New-Object System.Drawing.Size(130, 28)
$btnFechar.Anchor = "Top,Right"
$btnFechar.Add_Click({
    $form.Close()
    if ($script:CentralBackAction) { & $script:CentralBackAction }
}.GetNewClosure())
'@
    if (-not $source.Contains($closeOriginal.Trim())) {
        throw 'Não foi possível adaptar o botão de retorno da rotina original.'
    }
    $source = $source.Replace($closeOriginal.Trim(), $closeIntegrated.Trim())

    $uiDoEvents = '[System.Windows.Forms.Application]::DoEvents()'
    $uiLogIntegrated = @'
if ($script:CentralLogPath) {
    try {
        $linhaCentral = "{0} | {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Mensagem
        [IO.File]::AppendAllText($script:CentralLogPath, $linhaCentral + [Environment]::NewLine, [Text.UTF8Encoding]::new($true))
    }
    catch {}
}
[System.Windows.Forms.Application]::DoEvents()
'@
    $source = [regex]::Replace($source, [regex]::Escape($uiDoEvents), [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $uiLogIntegrated.Trim() }, 1)

    $integratedEnd = @'
$form.TopLevel = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Add_FormClosing({
    if ($script:ConectadoPW -and (Get-Command Undo-PWLogin -ErrorAction SilentlyContinue)) {
        try {
            Undo-PWLogin | Out-Null
            Write-UiLog 'Sessão do ProjectWise encerrada.'
        }
        catch {}
        $script:ConectadoPW = $false
    }
})
$script:CentralHostControl.Controls.Add($form)
$form.Show()
'@
    $source = [regex]::Replace($source, $endPattern, $integratedEnd.TrimEnd())
    $runner = [scriptblock]::Create($source)

    $logsDir = Join-Path $CentralRoot 'Logs\05-GerenciarUsuariosInativos'
    if (-not (Test-Path -LiteralPath $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
    $logPath = Join-Path $logsDir ("PW_UsuariosInativos_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    $moduleName = 'CentralUsuariosInativos_{0}' -f ([Guid]::NewGuid().ToString('N'))
    $routineModule = New-Module -Name $moduleName -ArgumentList @($runner, $Parent, $OnBack, $logPath) -ScriptBlock {
        param($RoutineScript, $HostControl, $BackAction, $CentralLog)
        $script:CentralHostControl = $HostControl
        $script:CentralBackAction = $BackAction
        $script:CentralLogPath = $CentralLog
        . $RoutineScript
    }
    $Parent.Tag = [PSCustomObject]@{ Process = $null; Module = $routineModule; LogPath = $logPath }

    # Visual da Central sobre a janela original (cores e bordas; posicoes nao mudam).
    if (Get-Command Set-TemaControles -ErrorAction SilentlyContinue) {
        $Parent.BackColor = $script:Tema.Fundo
        Set-TemaControles -Raiz $Parent -Titulo '05 - Gerenciar usuários inativos'
    }

    # Layout que se ajusta a altura (o original exige 760px e era cortado em telas
    # com escala). A grade de usuarios recebe a sobra; Etapa 4 e Acompanhamento
    # ficam lado a lado no rodape.
    $janela = @($Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Form] }) | Select-Object -First 1
    if ($janela -and (Get-Command Set-TemaLayoutEtapas -ErrorAction SilentlyContinue)) {
        $grupos = @($janela.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] })
        $g1 = $grupos | Where-Object { $_.Text -like 'Etapa 1*' } | Select-Object -First 1
        $g2 = $grupos | Where-Object { $_.Text -like 'Etapa 2*' } | Select-Object -First 1
        $g3 = $grupos | Where-Object { $_.Text -like 'Etapa 3*' } | Select-Object -First 1
        $g4 = $grupos | Where-Object { $_.Text -like 'Etapa 4*' } | Select-Object -First 1
        $gLog = $grupos | Where-Object { $_.Text -eq 'Acompanhamento' } | Select-Object -First 1
        if ($g1 -and $g2 -and $g3 -and $g4 -and $gLog) {
            Set-TemaLayoutEtapas -Janela $janela -Topo $g1.Top -Superiores @($g1, $g2) -Principal $g3 `
                -Inferiores @($g4, $gLog) -AlturaInferiorMin 80 -AlturaInferiorMax 120 -ProporcaoInferior 0.70 -AlturaPrincipalMin 200
        }
    }
}
