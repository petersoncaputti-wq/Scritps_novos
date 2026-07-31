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
}
