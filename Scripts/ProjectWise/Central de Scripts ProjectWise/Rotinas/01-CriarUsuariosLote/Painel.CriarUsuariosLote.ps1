function New-CriarUsuariosLotePanel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack
    )

    $workerScript = Join-Path $CentralRoot 'Rotinas\01-CriarUsuariosLote\CriarUsuariosLote.Integrado.ps1'
    $legacyScript = [System.IO.Path]::GetFullPath((Join-Path $CentralRoot '..\ExecucaoManualUsuario\03 - Criar Usuarios ProjectWise em Lote.ps1'))
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true
    $panel.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $panelLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $panelLayout.Dock = 'Fill'
    $panelLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    $panelLayout.ColumnCount = 1
    $panelLayout.RowCount = 2
    [void]$panelLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$panelLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))
    [void]$panelLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $panel.Controls.Add($panelLayout)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Fill'
    $header.Margin = New-Object System.Windows.Forms.Padding(0)
    $header.BackColor = [System.Drawing.Color]::FromArgb(22, 54, 92)
    $panelLayout.Controls.Add($header, 0, 0)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = '< Voltar'
    $backButton.FlatStyle = 'Flat'
    $backButton.FlatAppearance.BorderSize = 0
    $backButton.ForeColor = [System.Drawing.Color]::White
    $backButton.BackColor = [System.Drawing.Color]::FromArgb(22, 54, 92)
    $backButton.Location = New-Object System.Drawing.Point(12, 17)
    $backButton.Size = New-Object System.Drawing.Size(80, 30)
    $header.Controls.Add($backButton)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Criar usuários em lote'
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(105, 15)
    $header.Controls.Add($title)

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = 'Fill'
    $content.AutoScroll = $true
    $content.Margin = New-Object System.Windows.Forms.Padding(0)
    $content.Padding = New-Object System.Windows.Forms.Padding(18)
    $panelLayout.Controls.Add($content, 0, 1)

    $contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $contentLayout.Dock = 'Fill'
    $contentLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    $contentLayout.ColumnCount = 1
    $contentLayout.RowCount = 6
    [void]$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 122)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    $content.Controls.Add($contentLayout)

    $inputGroup = New-Object System.Windows.Forms.GroupBox
    $inputGroup.Text = 'Arquivo de entrada'
    $inputGroup.Dock = 'Fill'
    $inputGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $contentLayout.Controls.Add($inputGroup, 0, 0)

    $fileLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $fileLayout.Dock = 'Fill'
    $fileLayout.ColumnCount = 2
    $fileLayout.RowCount = 1
    $fileLayout.Padding = New-Object System.Windows.Forms.Padding(8, 7, 8, 7)
    [void]$fileLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$fileLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 140)))
    [void]$fileLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $inputGroup.Controls.Add($fileLayout)

    $fileBox = New-Object System.Windows.Forms.TextBox
    $fileBox.Dock = 'Fill'
    $fileBox.Margin = New-Object System.Windows.Forms.Padding(0, 7, 8, 5)
    $fileLayout.Controls.Add($fileBox, 0, 0)

    $fileButton = New-Object System.Windows.Forms.Button
    $fileButton.Text = 'Selecionar arquivo...'
    $fileButton.Dock = 'Fill'
    $fileButton.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
    $fileLayout.Controls.Add($fileButton, 1, 0)

    $optionsGroup = New-Object System.Windows.Forms.GroupBox
    $optionsGroup.Text = 'Opções da execução'
    $optionsGroup.Dock = 'Fill'
    $optionsGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $contentLayout.Controls.Add($optionsGroup, 0, 1)

    $validateOnly = New-Object System.Windows.Forms.CheckBox
    $validateOnly.Text = 'Somente validar (nenhuma alteração no ProjectWise)'
    $validateOnly.Checked = $true
    $validateOnly.AutoSize = $true
    $validateOnly.Location = New-Object System.Drawing.Point(16, 27)
    $optionsGroup.Controls.Add($validateOnly)

    $updateExisting = New-Object System.Windows.Forms.CheckBox
    $updateExisting.Text = 'Atualizar usuários existentes'
    $updateExisting.AutoSize = $true
    $updateExisting.Enabled = $false
    $updateExisting.Location = New-Object System.Drawing.Point(16, 55)
    $optionsGroup.Controls.Add($updateExisting)

    $skipAccess = New-Object System.Windows.Forms.CheckBox
    $skipAccess.Text = 'Não adicionar usuários a grupos e listas de usuários'
    $skipAccess.AutoSize = $true
    $skipAccess.Location = New-Object System.Drawing.Point(330, 55)
    $optionsGroup.Controls.Add($skipAccess)

    $safetyLabel = New-Object System.Windows.Forms.Label
    $safetyLabel.Text = 'Desmarcar “Somente validar” habilita alterações reais e exigirá confirmação.'
    $safetyLabel.ForeColor = [System.Drawing.Color]::FromArgb(170, 80, 20)
    $safetyLabel.AutoSize = $true
    $safetyLabel.Location = New-Object System.Drawing.Point(16, 82)
    $optionsGroup.Controls.Add($safetyLabel)

    $actionPanel = New-Object System.Windows.Forms.Panel
    $actionPanel.Dock = 'Fill'
    $actionPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $contentLayout.Controls.Add($actionPanel, 0, 2)

    $actionLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $actionLayout.Dock = 'Fill'
    $actionLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    $actionLayout.ColumnCount = 5
    $actionLayout.RowCount = 1
    [void]$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 17)))
    [void]$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 16)))
    [void]$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 16)))
    [void]$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 18)))
    [void]$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$actionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $actionPanel.Controls.Add($actionLayout)

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = 'Validar planilha'
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(31, 105, 178)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = 'Flat'
    $runButton.Dock = 'Fill'
    $runButton.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 5)
    $actionLayout.Controls.Add($runButton, 0, 0)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancelar execução'
    $cancelButton.Enabled = $false
    $cancelButton.Dock = 'Fill'
    $cancelButton.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 5)
    $actionLayout.Controls.Add($cancelButton, 1, 0)

    $openOutputButton = New-Object System.Windows.Forms.Button
    $openOutputButton.Text = 'Abrir última exportação'
    $openOutputButton.Enabled = $false
    $openOutputButton.Dock = 'Fill'
    $openOutputButton.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 5)
    $actionLayout.Controls.Add($openOutputButton, 2, 0)

    $legacyButton = New-Object System.Windows.Forms.Button
    $legacyButton.Text = 'Abrir versão original'
    $legacyButton.Dock = 'Fill'
    $legacyButton.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 5)
    $actionLayout.Controls.Add($legacyButton, 3, 0)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Style = 'Marquee'
    $progress.MarqueeAnimationSpeed = 0
    $progress.Dock = 'Fill'
    $progress.Margin = New-Object System.Windows.Forms.Padding(6, 13, 0, 13)
    $actionLayout.Controls.Add($progress, 4, 0)

    $logLabel = New-Object System.Windows.Forms.Label
    $logLabel.Text = 'Acompanhamento e resultados'
    $logLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $logLabel.AutoSize = $true
    $logLabel.Dock = 'Fill'
    $logLabel.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
    $contentLayout.Controls.Add($logLabel, 0, 3)

    $resultTabs = New-Object System.Windows.Forms.TabControl
    $resultTabs.Dock = 'Fill'
    $resultTabs.Margin = New-Object System.Windows.Forms.Padding(0)
    $contentLayout.Controls.Add($resultTabs, 0, 4)

    $logTab = New-Object System.Windows.Forms.TabPage
    $logTab.Text = 'Acompanhamento'
    $resultTabs.TabPages.Add($logTab)

    $tableTab = New-Object System.Windows.Forms.TabPage
    $tableTab.Text = 'Resultado por usuário'
    $resultTabs.TabPages.Add($tableTab)

    $resultLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $resultLayout.Dock = 'Fill'
    $resultLayout.ColumnCount = 1
    $resultLayout.RowCount = 2
    [void]$resultLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$resultLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 38)))
    [void]$resultLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $tableTab.Controls.Add($resultLayout)

    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.ReadOnly = $true
    $logBox.BackColor = [System.Drawing.Color]::FromArgb(28, 31, 36)
    $logBox.ForeColor = [System.Drawing.Color]::Gainsboro
    $logBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $logBox.Dock = 'Fill'
    $logBox.Margin = New-Object System.Windows.Forms.Padding(0)
    $logTab.Controls.Add($logBox)

    $resultSummary = New-Object System.Windows.Forms.Label
    $resultSummary.Text = 'O resumo será exibido depois da execução.'
    $resultSummary.Dock = 'Top'
    $resultSummary.Height = 38
    $resultSummary.Padding = New-Object System.Windows.Forms.Padding(8, 10, 8, 4)
    $resultSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $resultLayout.Controls.Add($resultSummary, 0, 0)

    $resultGrid = New-Object System.Windows.Forms.DataGridView
    $resultGrid.Dock = 'Fill'
    $resultGrid.ReadOnly = $true
    $resultGrid.AllowUserToAddRows = $false
    $resultGrid.AllowUserToDeleteRows = $false
    $resultGrid.AllowUserToResizeRows = $false
    $resultGrid.RowHeadersVisible = $false
    $resultGrid.SelectionMode = 'FullRowSelect'
    $resultGrid.AutoGenerateColumns = $false
    $resultGrid.BackgroundColor = [System.Drawing.Color]::White
    $resultLayout.Controls.Add($resultGrid, 0, 1)

    foreach ($columnInfo in @(
        @('Linha', 'Linha', 55, 'NotSet'),
        @('Usuário', 'UserName', 165, 'NotSet'),
        @('E-mail', 'Email', 220, 'NotSet'),
        @('Status', 'Status', 75, 'NotSet'),
        @('Ação', 'Acao', 100, 'NotSet'),
        @('Mensagem', 'Mensagem', 260, 'Fill')
    )) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.HeaderText = $columnInfo[0]
        $column.DataPropertyName = $columnInfo[1]
        $column.Width = [int]$columnInfo[2]
        if ($columnInfo[3] -eq 'Fill') { $column.AutoSizeMode = 'Fill' }
        [void]$resultGrid.Columns.Add($column)
    }

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = 'Pronto para validar uma planilha.'
    $statusLabel.AutoSize = $true
    $statusLabel.Dock = 'Fill'
    $statusLabel.Margin = New-Object System.Windows.Forms.Padding(0, 7, 0, 0)
    $contentLayout.Controls.Add($statusLabel, 0, 5)

    $state = [pscustomobject]@{
        Process     = $null
        Finished    = $false
        Cancelled   = $false
        StartedAt   = $null
        ReportPath  = $null
        LogPath     = $null
        LogLinesRead = 0
        ShowCompletionDialog = $true
        ExportPath  = $null
        ExportAction = $null
    }
    $panel.Tag = $state

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 150

    $appendLine = {
        param([string]$Line, [System.Drawing.Color]$Color)
        if ([string]::IsNullOrWhiteSpace($Line)) { return }
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.SelectionColor = $Color
        $logBox.AppendText($Line + [Environment]::NewLine)
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.ScrollToCaret()
    }.GetNewClosure()

    $setRunning = {
        param([bool]$Running)
        $runButton.Enabled = -not $Running
        $fileButton.Enabled = -not $Running
        $validateOnly.Enabled = -not $Running
        $updateExisting.Enabled = (-not $Running -and -not $validateOnly.Checked)
        $skipAccess.Enabled = -not $Running
        $legacyButton.Enabled = -not $Running
        $openOutputButton.Enabled = (-not $Running -and -not [string]::IsNullOrWhiteSpace($state.ExportPath))
        $cancelButton.Enabled = $Running
        $backButton.Enabled = -not $Running
        $progress.MarqueeAnimationSpeed = if ($Running) { 30 } else { 0 }
    }.GetNewClosure()

    $quoteArgument = {
        param([string]$Value)
        return '"' + $Value.Replace('"', '\"') + '"'
    }

    $writeCentralLog = {
        param([string]$Message)
        if (Get-Command Write-ExecutorLog -ErrorAction SilentlyContinue) {
            Write-ExecutorLog $Message
        }
    }.GetNewClosure()

    $activateCentral = {
        try {
            $hostForm = $panel.FindForm()
            if ($hostForm -and -not $hostForm.IsDisposed) {
                if ($hostForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
                    $hostForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                }
                $hostForm.Show()
                $hostForm.Activate()
                $hostForm.BringToFront()
            }
        }
        catch { }
    }.GetNewClosure()

    $loadResults = {
        $logsFolder = Join-Path (Split-Path -Parent $workerScript) 'Logs'
        $report = Get-ChildItem -LiteralPath $logsFolder -Filter 'PW_CriarUsuariosLote_Relatorio_*.csv' -File -ErrorAction SilentlyContinue |
            Where-Object { $state.StartedAt -and $_.LastWriteTime -ge $state.StartedAt.AddSeconds(-2) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if (-not $report) {
            $resultSummary.Text = 'Nenhum relatório foi produzido nesta execução.'
            $resultGrid.DataSource = $null
            return $null
        }

        $state.ReportPath = $report.FullName
        $rows = @(Import-Csv -LiteralPath $report.FullName -Delimiter ';')
        $criados = @($rows | Where-Object { $_.Status -eq 'OK' -and $_.Acao -eq 'Criado' }).Count
        $atualizados = @($rows | Where-Object { $_.Status -eq 'OK' -and $_.Acao -eq 'Atualizado' }).Count
        $existentes = @($rows | Where-Object { $_.Status -eq 'OK' -and $_.Acao -eq 'JaExistia' }).Count
        $validos = @($rows | Where-Object { $_.Status -eq 'VALIDO' }).Count
        $invalidos = @($rows | Where-Object { $_.Status -eq 'INVALIDO' }).Count
        $errosResultado = @($rows | Where-Object { $_.Status -eq 'ERRO' }).Count

        $summary = [pscustomobject]@{
            Total       = $rows.Count
            Criados     = $criados
            Atualizados = $atualizados
            Existentes  = $existentes
            Validados   = $validos
            Invalidos   = $invalidos
            Erros       = $errosResultado
        }
        $resultSummary.Text = "Total: $($summary.Total)  |  Criados: $criados  |  Atualizados: $atualizados  |  Já existentes: $existentes  |  Inválidos: $invalidos  |  Erros: $errosResultado"
        $resultGrid.DataSource = [System.Collections.ArrayList]@($rows)
        return $summary
    }.GetNewClosure()

    $exportResults = {
        param([Parameter(Mandatory)][string]$DestinationFolder)

        if (-not $state.ReportPath -or -not (Test-Path -LiteralPath $state.ReportPath -PathType Leaf)) {
            throw 'O relatório interno desta execução não foi localizado.'
        }

        Import-Module ImportExcel -ErrorAction Stop
        $rows = @(Import-Csv -LiteralPath $state.ReportPath -Delimiter ';')
        if ($rows.Count -eq 0) { throw 'O relatório não possui registros para exportação.' }

        if (-not (Test-Path -LiteralPath $DestinationFolder)) {
            New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
        }
        $exportFolder = Join-Path $DestinationFolder ("UsuariosPorDescricao_{0:yyyyMMdd_HHmmss}" -f (Get-Date))
        New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null

        $files = New-Object System.Collections.Generic.List[string]
        foreach ($group in @($rows | Group-Object Description)) {
            $description = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Sem descrição' } else { $group.Name }
            $safeName = $description
            foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
                $safeName = $safeName.Replace($invalidChar, '_')
            }
            $safeName = ($safeName -replace '\s+', ' ').Trim()
            if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0, 80).Trim() }
            if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'SemDescricao' }

            $exportFile = Join-Path $exportFolder ("Usuarios_Descricao_{0}.xlsx" -f $safeName)
            @($group.Group | Sort-Object Email, UserName | Select-Object Email, Description, UserName, Status, Acao, Mensagem) |
                Export-Excel -Path $exportFile -WorksheetName 'Usuarios' -AutoSize -ClearSheet
            $files.Add($exportFile)
        }

        $state.ExportPath = $exportFolder
        $openOutputButton.Enabled = $true
        & $writeCentralLog "EXPORTAÇÃO CONCLUÍDA | criar-usuarios | $exportFolder | $($files.Count) arquivo(s)"
        return [pscustomobject]@{ Folder = $exportFolder; Files = $files.ToArray() }
    }.GetNewClosure()
    $state.ExportAction = $exportResults

    $validateOnly.Add_CheckedChanged({
        $updateExisting.Enabled = -not $validateOnly.Checked
        if ($validateOnly.Checked) {
            $updateExisting.Checked = $false
            $runButton.Text = 'Validar planilha'
        }
        else {
            $runButton.Text = 'Criar usuários'
        }
    }.GetNewClosure())

    $fileButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Arquivos CSV (*.csv)|*.csv'
        $dialog.Title = 'Selecione a planilha de usuários'
        if ($dialog.ShowDialog($panel.FindForm()) -eq [System.Windows.Forms.DialogResult]::OK) {
            $fileBox.Text = $dialog.FileName
        }
    }.GetNewClosure())

    $openOutputButton.Add_Click({
        if ($state.ExportPath -and (Test-Path -LiteralPath $state.ExportPath -PathType Container)) {
            Start-Process explorer.exe -ArgumentList ('"{0}"' -f $state.ExportPath)
        }
    }.GetNewClosure())

    $legacyButton.Add_Click({
        if (-not (Test-Path -LiteralPath $legacyScript -PathType Leaf)) {
            [void][System.Windows.Forms.MessageBox]::Show("Script original não encontrado:`r`n$legacyScript", 'Erro', 'OK', 'Error')
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            'A versão original será aberta em uma janela separada. Deseja continuar?',
            'Executar versão original', 'YesNo', 'Question')
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process powershell.exe -WorkingDirectory (Split-Path -Parent $legacyScript) -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -MTA -File "{0}"' -f $legacyScript)
        }
    }.GetNewClosure())

    $backButton.Add_Click({ & $OnBack }.GetNewClosure())

    $runButton.Add_Click({
        if (-not (Test-Path -LiteralPath $workerScript -PathType Leaf)) {
            [void][System.Windows.Forms.MessageBox]::Show("Arquivo integrado não encontrado:`r`n$workerScript", 'Erro', 'OK', 'Error')
            return
        }
        if (-not (Test-Path -LiteralPath $fileBox.Text -PathType Leaf)) {
            [void][System.Windows.Forms.MessageBox]::Show('Selecione uma planilha válida.', 'Arquivo obrigatório', 'OK', 'Warning')
            return
        }
        if ([IO.Path]::GetExtension($fileBox.Text).ToLowerInvariant() -notin @('.csv', '.xlsx', '.xlsm')) {
            [void][System.Windows.Forms.MessageBox]::Show('O arquivo deve ser CSV, XLSX ou XLSM.', 'Formato inválido', 'OK', 'Warning')
            return
        }
        if (-not $validateOnly.Checked) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
            'Esta execução poderá criar ou atualizar usuários no ProjectWise. Deseja continuar?',
                'Confirmar alteração real', 'YesNo', 'Warning')
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $arguments = @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-MTA',
            '-File', (& $quoteArgument $workerScript),
            '-ModoIntegrado', '-NaoPausar',
            '-CaminhoArquivo', (& $quoteArgument $fileBox.Text)
        )
        if ($validateOnly.Checked) { $arguments += '-SomenteValidar' }
        if ($updateExisting.Checked) { $arguments += '-AtualizarExistentes' }
        if ($skipAccess.Checked) { $arguments += '-NaoAdicionarAcessos' }

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = 'powershell.exe'
        $startInfo.Arguments = $arguments -join ' '
        $startInfo.WorkingDirectory = Split-Path -Parent $workerScript
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $state.Cancelled = $false
        $state.Finished = $false
        $state.StartedAt = Get-Date
        $state.ReportPath = $null
        $state.ExportPath = $null
        $openOutputButton.Enabled = $false
        $state.LogPath = $null
        $state.LogLinesRead = 0
        $logBox.Clear()
        $resultGrid.DataSource = $null
        $resultSummary.Text = 'Execução em andamento. Aguarde o processamento.'
        $resultTabs.SelectedTab = $logTab
        & $appendLine ('Iniciando em modo {0}...' -f $(if ($validateOnly.Checked) { 'SOMENTE VALIDAR' } else { 'ALTERAÇÃO REAL' })) ([System.Drawing.Color]::LightSkyBlue)

        try {
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $startInfo
            if (-not $process.Start()) { throw 'O processo não pôde ser iniciado.' }
            $state.Process = $process
            & $writeCentralLog "INÍCIO ROTINA INTEGRADA | criar-usuarios | PID $($process.Id) | $($fileBox.Text)"
            $statusLabel.Text = "Execução em andamento (PID $($process.Id))..."
            & $setRunning $true
            $timer.Start()
        }
        catch {
            & $setRunning $false
            & $appendLine ("Falha ao iniciar: " + $_.Exception.Message) ([System.Drawing.Color]::Salmon)
            $statusLabel.Text = 'Falha ao iniciar a execução.'
            & $writeCentralLog "ERRO AO INICIAR ROTINA INTEGRADA | criar-usuarios | $($_.Exception.Message)"
        }
    }.GetNewClosure())

    $cancelButton.Add_Click({
        if ($state.Process -and -not $state.Process.HasExited) {
            $answer = [System.Windows.Forms.MessageBox]::Show('Deseja interromper a execução atual?', 'Cancelar execução', 'YesNo', 'Warning')
            if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
                $state.Cancelled = $true
                $state.Process.Kill()
                & $appendLine 'Cancelamento solicitado pelo usuário.' ([System.Drawing.Color]::Khaki)
                & $writeCentralLog "CANCELAMENTO SOLICITADO | criar-usuarios | PID $($state.Process.Id)"
            }
        }
    }.GetNewClosure())

    $timer.Add_Tick({
        try {
            if ($state.Process -and -not $state.LogPath) {
                $logsFolder = Join-Path (Split-Path -Parent $workerScript) 'Logs'
                $currentLog = Get-ChildItem -LiteralPath $logsFolder -Filter 'PW_CriarUsuariosLote_*.log' -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -ge $state.StartedAt.AddSeconds(-2) } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1
                if ($currentLog) { $state.LogPath = $currentLog.FullName }
            }

            if ($state.LogPath -and (Test-Path -LiteralPath $state.LogPath)) {
                $logLines = @(Get-Content -LiteralPath $state.LogPath -Encoding UTF8 -ErrorAction SilentlyContinue)
                if ($logLines.Count -gt $state.LogLinesRead) {
                    for ($index = $state.LogLinesRead; $index -lt $logLines.Count; $index++) {
                        $line = [string]$logLines[$index]
                        $color = if ($line -match '\| ERROR \||Falha geral:') { [System.Drawing.Color]::Salmon } elseif ($line -match '\| WARN \|') { [System.Drawing.Color]::Khaki } elseif ($line -match '\| OK \|') { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Gainsboro }
                        & $appendLine $line $color
                        if ($line -match 'Autenticação realizada') {
                            & $activateCentral
                            $statusLabel.Text = 'Autenticação concluída. Processando os usuários...'
                        }
                    }
                    $state.LogLinesRead = $logLines.Count
                }
            }

            if ($state.Process -and $state.Process.HasExited) {
                $timer.Stop()
                $exitCode = $state.Process.ExitCode
                & $setRunning $false
                & $activateCentral
                $summary = & $loadResults
                if ($state.Cancelled) {
                    $statusLabel.Text = 'Execução cancelada.'
                    & $writeCentralLog "ROTINA CANCELADA | criar-usuarios | Código $exitCode"
                }
                elseif ($exitCode -eq 0) {
                    if ($summary) {
                        $statusLabel.Text = "Concluído: $($summary.Criados) criado(s), $($summary.Atualizados) atualizado(s) e $($summary.Existentes) já existente(s)."
                        $resultTabs.SelectedTab = $tableTab
                        $notification = "Execução concluída.`r`n`r`nTotal: $($summary.Total)`r`nCriados: $($summary.Criados)`r`nAtualizados: $($summary.Atualizados)`r`nJá existentes: $($summary.Existentes)`r`nInválidos: $($summary.Invalidos)`r`nErros: $($summary.Erros)`r`n`r`nDeseja exportar a lista de resultados?"
                        if ($state.ShowCompletionDialog) {
                            $exportAnswer = [System.Windows.Forms.MessageBox]::Show($panel.FindForm(), $notification, 'Resultado da criação de usuários', 'YesNo', 'Question')
                            if ($exportAnswer -eq [System.Windows.Forms.DialogResult]::Yes) {
                                $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                                $folderDialog.Description = 'Selecione a pasta onde a lista será exportada'
                                $folderDialog.ShowNewFolderButton = $true
                                if ($folderDialog.ShowDialog($panel.FindForm()) -eq [System.Windows.Forms.DialogResult]::OK) {
                                    try {
                                        $exportResult = & $exportResults $folderDialog.SelectedPath
                                        $statusLabel.Text = "Execução concluída. Lista exportada para: $($exportResult.Folder)"
                                        [void][System.Windows.Forms.MessageBox]::Show($panel.FindForm(), "Exportação concluída.`r`n`r`n$($exportResult.Files.Count) arquivo(s) criado(s) em:`r`n$($exportResult.Folder)", 'Exportação concluída', 'OK', 'Information')
                                    }
                                    catch {
                                        $statusLabel.Text = 'Execução concluída, mas ocorreu uma falha na exportação.'
                                        & $writeCentralLog "ERRO DE EXPORTAÇÃO | criar-usuarios | $($_.Exception.Message)"
                                        [void][System.Windows.Forms.MessageBox]::Show($panel.FindForm(), "Não foi possível exportar a lista:`r`n`r`n$($_.Exception.Message)", 'Falha na exportação', 'OK', 'Error')
                                    }
                                }
                                else {
                                    $statusLabel.Text = 'Execução concluída. Exportação cancelada pelo usuário.'
                                }
                            }
                        }
                    }
                    else {
                        $statusLabel.Text = 'Execução concluída, mas nenhum relatório foi localizado.'
                    }
                    & $appendLine 'Processo finalizado com sucesso.' ([System.Drawing.Color]::LightGreen)
                    & $writeCentralLog "FIM ROTINA INTEGRADA | criar-usuarios | Código 0 | Relatório: $($state.ReportPath)"
                }
                else {
                    $statusLabel.Text = "Execução concluída com erro (código $exitCode)."
                    & $appendLine "Processo finalizado com código de erro $exitCode." ([System.Drawing.Color]::Salmon)
                    $resultTabs.SelectedTab = $logTab
                    [void][System.Windows.Forms.MessageBox]::Show($panel.FindForm(), "A rotina terminou com erro (código $exitCode). Consulte a guia Acompanhamento.", 'Falha na criação de usuários', 'OK', 'Error')
                    & $writeCentralLog "FIM ROTINA INTEGRADA COM ERRO | criar-usuarios | Código $exitCode"
                }
                $state.Finished = $true
                $state.Process.Dispose()
                $state.Process = $null
            }
        }
        catch {
            $timer.Stop()
            & $setRunning $false
            $statusLabel.Text = 'Falha ao atualizar a interface. A Central permanecerá aberta.'
            & $appendLine ("Erro da interface: " + $_.Exception.Message) ([System.Drawing.Color]::Salmon)
            & $writeCentralLog "ERRO DE INTERFACE | criar-usuarios | $($_.Exception.Message)"
            & $activateCentral
        }
    }.GetNewClosure())

    $panel.Add_Disposed({
        $timer.Stop()
        $timer.Dispose()
        if ($state.Process -and -not $state.Process.HasExited) {
            $state.Process.Kill()
        }
    }.GetNewClosure())

    $Parent.Controls.Add($panel)
    $panel.BringToFront()
    return $panel
}
