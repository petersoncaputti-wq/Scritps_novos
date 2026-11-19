function Open-GerenciarAcessosPanel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack
    )

    $legacyScript = [System.IO.Path]::GetFullPath((Join-Path $CentralRoot '..\ExecucaoManualUsuario\02 - Gerenciar Acessos de Projetos.ps1'))
    if (-not (Test-Path -LiteralPath $legacyScript -PathType Leaf)) {
        throw "Script original não encontrado: $legacyScript"
    }

    $source = Get-Content -LiteralPath $legacyScript -Raw
    $startMarker = '$form = New-Object System.Windows.Forms.Form'
    $endPattern = 'Atualizar-EstadoInterface\r?\n\[void\]\$form\.ShowDialog\(\)\s*$'
    if (-not $source.Contains($startMarker) -or -not [regex]::IsMatch($source, $endPattern)) {
        throw 'A estrutura do script original mudou e não pôde ser integrada com segurança.'
    }

    $centralLogs = Join-Path $CentralRoot 'Logs\02-GerenciarAcessos'
    if (-not (Test-Path -LiteralPath $centralLogs)) { New-Item -ItemType Directory -Path $centralLogs -Force | Out-Null }
    $escapedLogs = $centralLogs.Replace("'", "''")
    $source = $source.Replace('$PastaLog = Join-Path $PSScriptRoot "Logs"', "`$PastaLog = '$escapedLogs'")

    $doEventsOriginal = '[System.Windows.Forms.Application]::DoEvents()'
    $doEventsIntegrated = @'
if ($script:CentralStatusLabel) {
    $script:CentralStatusLabel.Text = $Mensagem
    if ($script:CentralProgressBar) {
        if ($Mensagem -like 'Carregando projetos*' -or $Mensagem -like 'Lendo acessos do projeto*') {
            $script:CentralProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            $script:CentralProgressBar.MarqueeAnimationSpeed = 28
            $script:CentralProgressBar.Visible = $true
        }
        elseif ($Mensagem -like "Concessao * carregada com sucesso.*" -or $Mensagem -like "Concessão * carregada com sucesso.*") {
            $script:CentralProgressBar.MarqueeAnimationSpeed = 0
            $script:CentralProgressBar.Visible = $false
        }
    }
}
[System.Windows.Forms.Application]::DoEvents()
'@
    if (-not $source.Contains($doEventsOriginal)) {
        throw 'Não foi possível integrar o acompanhamento visual da rotina original.'
    }
    $source = $source.Replace($doEventsOriginal, $doEventsIntegrated.Trim())

    # O evento Shown pode ser disparado fora do escopo em que a função foi
    # declarada. Usa uma referência direta para impedir CommandNotFoundException.
    $layoutEventsOriginal = @'
$form.Add_Shown({ Ajustar-LayoutCampos })
$form.Add_Resize({ Ajustar-LayoutCampos })
$groupUsuario.Add_Resize({ Ajustar-LayoutCampos })
$groupSelecao.Add_Resize({ Ajustar-LayoutCampos })
'@
    $layoutEventsIntegrated = @'
$script:AjustarLayoutCamposAction = ${function:Ajustar-LayoutCampos}.GetNewClosure()
$form.Add_Shown({ & $script:AjustarLayoutCamposAction }.GetNewClosure())
$form.Add_Resize({ & $script:AjustarLayoutCamposAction }.GetNewClosure())
$groupUsuario.Add_Resize({ & $script:AjustarLayoutCamposAction }.GetNewClosure())
$groupSelecao.Add_Resize({ & $script:AjustarLayoutCamposAction }.GetNewClosure())
'@
    if (-not $source.Contains($layoutEventsOriginal.Trim())) {
        throw 'Não foi possível adaptar os eventos de redimensionamento da rotina original.'
    }
    $source = $source.Replace($layoutEventsOriginal.Trim(), $layoutEventsIntegrated.Trim())

    $closeButtonOriginal = @'
$btnFechar.Text = "Fechar"
$btnFechar.Location = New-Object System.Drawing.Point(1070,20)
$btnFechar.Size = New-Object System.Drawing.Size(90,28)
$btnFechar.Add_Click({ $form.Close() })
'@
    $closeButtonIntegrated = @'
$btnFechar.Text = "Voltar à Central"
$btnFechar.Location = New-Object System.Drawing.Point(1035,20)
$btnFechar.Size = New-Object System.Drawing.Size(125,28)
$btnFechar.Add_Click({
    $form.Close()
    if ($script:CentralBackAction) { & $script:CentralBackAction }
}.GetNewClosure())
'@
    if (-not $source.Contains($closeButtonOriginal.Trim())) {
        throw 'Não foi possível adaptar o botão de retorno da rotina original.'
    }
    $source = $source.Replace($closeButtonOriginal.Trim(), $closeButtonIntegrated.Trim())

    $integratedEnd = @'
Atualizar-EstadoInterface
$form.TopLevel = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$centralStatus = New-Object System.Windows.Forms.StatusStrip
$centralStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom
$centralStatus.SizingGrip = $false
$script:CentralStatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:CentralStatusLabel.Spring = $true
$script:CentralStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:CentralStatusLabel.Text = 'Pronto.'
$script:CentralProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$script:CentralProgressBar.Size = New-Object System.Drawing.Size(180, 16)
$script:CentralProgressBar.Visible = $false
[void]$centralStatus.Items.Add($script:CentralStatusLabel)
[void]$centralStatus.Items.Add($script:CentralProgressBar)
$form.Controls.Add($centralStatus)
$centralStatus.BringToFront()
$script:CentralHostControl.Controls.Add($form)
$form.Show()
'@
    $source = [regex]::Replace($source, $endPattern, $integratedEnd.TrimEnd())
    $runner = [scriptblock]::Create($source)

    # O módulo dinâmico preserva funções e variáveis usadas pelos eventos da
    # interface mesmo depois que esta função de carregamento termina.
    $moduleName = 'CentralGerenciarAcessos_{0}' -f ([Guid]::NewGuid().ToString('N'))
    $routineModule = New-Module -Name $moduleName -ArgumentList @($runner, $Parent, $OnBack) -ScriptBlock {
        param($RoutineScript, $HostControl, $BackAction)
        $script:CentralHostControl = $HostControl
        $script:CentralBackAction = $BackAction
        . $RoutineScript
    }
    $Parent.Tag = [PSCustomObject]@{
        Process = $null
        Module  = $routineModule
    }

    # Visual da Central sobre a janela original (cores e bordas; posicoes nao mudam).
    if (Get-Command Set-TemaControles -ErrorAction SilentlyContinue) {
        $Parent.BackColor = $script:Tema.Fundo
        Set-TemaControles -Raiz $Parent -Titulo '02 - Gerenciar acessos de projetos'
    }

    # ---------------------------------------------------------
    # Layout que se ajusta a altura disponivel.
    # O original empilha 4 etapas e exige 810px de altura: em tela 1920x1080
    # com escala de 125% (793px uteis maximizado) o fim era cortado e a lista
    # de projetos ficava com ~100px. Aqui "Resumo" e "Acompanhamento" ficam
    # lado a lado e a lista de projetos recebe a sobra de altura.
    # ---------------------------------------------------------
    $janela = @($Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Form] }) | Select-Object -First 1
    if ($janela -and (Get-Command Set-TemaLayoutEtapas -ErrorAction SilentlyContinue)) {
        $grupos = @($janela.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] })
        $gUsuario = $grupos | Where-Object { $_.Text -like 'Etapa 1*' } | Select-Object -First 1
        $gSelecao = $grupos | Where-Object { $_.Text -like 'Etapa 2*' } | Select-Object -First 1
        $gResumo = $grupos | Where-Object { $_.Text -like 'Etapa 3*' } | Select-Object -First 1
        $gLog = $grupos | Where-Object { $_.Text -eq 'Acompanhamento' } | Select-Object -First 1
        if ($gUsuario -and $gSelecao -and $gResumo -and $gLog) {
            Set-TemaLayoutEtapas -Janela $janela -Topo $gUsuario.Top -Superiores @($gUsuario) -Principal $gSelecao `
                -Inferiores @($gResumo, $gLog) -AlturaInferiorMin 130 -AlturaInferiorMax 200 -AlturaPrincipalMin 240
        }
    }
}
