function New-IncluirUsuariosWebPanel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack,
        [string]$PanelTitle = '03 - Incluir usuários em projetos PW Web',
        [string]$PythonRelativePath = '..\ExecucaoManualUsuario\PWDM_Gerenciamento_Participantes_V2\incluir_usuarios_projetos_roles_pwdm_v2.py',
        [string]$LogPrefix = 'PW_Web_IncluirUsuarios',
        [string]$LogDirectory = '03-IncluirUsuariosWeb',
        [ValidateSet('Python','PowerShell')][string]$Runtime = 'Python',
        [string]$HelpText = '',
        [string]$StartConfirmationText = '',
        [string]$RuntimeArguments = '',
        [string]$SecondaryButtonText = '',
        [string]$SecondaryRuntimeArguments = '',
        [string]$SecondaryConfirmationText = ''
    )

    if (-not ('CentralScripts.StreamPump' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Collections.Concurrent;
using System.IO;
using System.Threading.Tasks;
namespace CentralScripts {
    public static class StreamPump {
        public static Task Start(StreamReader reader, ConcurrentQueue<string> queue, string prefix) {
            return Task.Run(() => {
                int value;
                bool first = true;
                while ((value = reader.Read()) >= 0) {
                    string text = new string((char)value, 1);
                    queue.Enqueue((first ? prefix : "") + text);
                    first = false;
                }
            });
        }
        public static Task StartFilteredError(StreamReader reader, ConcurrentQueue<string> queue) {
            return Task.Run(() => {
                string text = reader.ReadToEnd();
                if (string.IsNullOrWhiteSpace(text)) return;
                bool isPowerShellProgress = text.StartsWith("#< CLIXML")
                    && text.Contains("S=\"progress\"")
                    && !text.Contains("S=\"Error\"");
                if (!isPowerShellProgress) queue.Enqueue("[ERRO] " + text);
            });
        }
    }
}
'@
    }

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $Parent.Controls.Add($panel)
    $panel.BringToFront()

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object System.Windows.Forms.Padding(16)
    $layout.ColumnCount = 1
    $layout.RowCount = 5
    [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 48)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 62)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 0)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
    $panel.Controls.Add($layout)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Fill'
    $layout.Controls.Add($header, 0, 0)
    $title = New-Object System.Windows.Forms.Label
    $title.Text = $PanelTitle
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(22, 54, 92)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(0, 6)
    $header.Controls.Add($title)
    $help = New-Object System.Windows.Forms.Label
    $help.Dock = 'Fill'
    $help.Text = if ([string]::IsNullOrWhiteSpace($HelpText)) {
        "A rotina abrirá o seletor de projetos e o navegador Bentley quando necessário.`r`nAs perguntas e confirmações serão apresentadas automaticamente em janelas de diálogo."
    }
    else { $HelpText }
    $help.ForeColor = [System.Drawing.Color]::FromArgb(70, 78, 88)
    $layout.Controls.Add($help, 0, 1)

    $output = New-Object System.Windows.Forms.RichTextBox
    $output.Dock = 'Fill'
    $output.ReadOnly = $true
    $output.BackColor = [System.Drawing.Color]::FromArgb(24, 28, 34)
    $output.ForeColor = [System.Drawing.Color]::Gainsboro
    $output.Font = New-Object System.Drawing.Font('Consolas', 9.5)
    $layout.Controls.Add($output, 0, 2)

    $answerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $answerLayout.Dock = 'Fill'
    $answerLayout.Visible = $false
    $answerLayout.ColumnCount = 2
    [void]$answerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$answerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 175)))
    $layout.Controls.Add($answerLayout, 0, 3)
    $answerBox = New-Object System.Windows.Forms.TextBox
    $answerBox.Dock = 'Fill'
    $answerBox.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 7)
    $answerLayout.Controls.Add($answerBox, 0, 0)
    $answerButton = New-Object System.Windows.Forms.Button
    $answerButton.Text = 'Responder / Continuar'
    $answerButton.Dock = 'Fill'
    $answerButton.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 4)
    $answerButton.Enabled = $false
    $answerLayout.Controls.Add($answerButton, 1, 0)

    $actions = New-Object System.Windows.Forms.FlowLayoutPanel
    $actions.Dock = 'Fill'
    $actions.FlowDirection = 'LeftToRight'
    $layout.Controls.Add($actions, 0, 4)
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = 'Iniciar rotina'
    $startButton.Size = New-Object System.Drawing.Size(135, 34)
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(31, 105, 178)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.FlatStyle = 'Flat'
    $actions.Controls.Add($startButton)
    $secondaryButton = New-Object System.Windows.Forms.Button
    $secondaryButton.Text = $SecondaryButtonText
    $secondaryButton.Size = New-Object System.Drawing.Size(135, 34)
    $secondaryButton.BackColor = [System.Drawing.Color]::FromArgb(76, 125, 76)
    $secondaryButton.ForeColor = [System.Drawing.Color]::White
    $secondaryButton.FlatStyle = 'Flat'
    $secondaryButton.Visible = -not [string]::IsNullOrWhiteSpace($SecondaryButtonText)
    $actions.Controls.Add($secondaryButton)
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancelar execução'
    $cancelButton.Size = New-Object System.Drawing.Size(145, 34)
    $cancelButton.Enabled = $false
    $actions.Controls.Add($cancelButton)
    $footerBackButton = New-Object System.Windows.Forms.Button
    $footerBackButton.Text = 'Voltar à Central'
    $footerBackButton.Size = New-Object System.Drawing.Size(135, 34)
    $actions.Controls.Add($footerBackButton)
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = 'Pronto para iniciar.'
    $statusLabel.AutoSize = $true
    $statusLabel.Margin = New-Object System.Windows.Forms.Padding(14, 9, 3, 3)
    $actions.Controls.Add($statusLabel)

    $state = [PSCustomObject]@{
        Process = $null
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
        LogPath = $null
        Finished = $false
        OutputTask = $null
        ErrorTask = $null
        PromptBuffer = ''
        DialogOpen = $false
        LastOutputUtc = [DateTime]::UtcNow
        NextRuntimeArguments = $RuntimeArguments
        NextConfirmationText = $StartConfirmationText
    }
    $panel.Tag = $state
    $logsDir = Join-Path (Join-Path $CentralRoot 'Logs') $LogDirectory
    if (-not (Test-Path -LiteralPath $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }

    $appendOutput = {
        param([string]$Line)
        if ($null -eq $Line) { return }
        $output.AppendText($Line + [Environment]::NewLine)
        $output.SelectionStart = $output.TextLength
        $output.ScrollToCaret()
        if ($state.LogPath) {
            try { [IO.File]::AppendAllText($state.LogPath, $Line + [Environment]::NewLine, [Text.UTF8Encoding]::new($true)) } catch {}
        }
    }.GetNewClosure()

    $appendProcessOutput = {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return }
        $output.AppendText($Text)
        $output.SelectionStart = $output.TextLength
        $output.ScrollToCaret()
        $state.PromptBuffer += $Text
        $state.LastOutputUtc = [DateTime]::UtcNow
        if ($state.PromptBuffer.Length -gt 3000) {
            $state.PromptBuffer = $state.PromptBuffer.Substring($state.PromptBuffer.Length - 3000)
        }
        if ($state.LogPath) {
            try { [IO.File]::AppendAllText($state.LogPath, $Text, [Text.UTF8Encoding]::new($true)) } catch {}
        }
    }.GetNewClosure()

    $writeResponse = {
        param([string]$Value)
        if (-not $state.Process -or $state.Process.HasExited) { return }
        $state.Process.StandardInput.WriteLine($Value)
        $state.Process.StandardInput.Flush()
        $displayValue = if ([string]::IsNullOrEmpty($Value)) { '[CONTINUAR]' } else { $Value }
        & $appendOutput "> $displayValue"
        $state.PromptBuffer = ''
    }.GetNewClosure()

    $askText = {
        param([string]$Question, [string]$Title)
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $Title
        $dialog.StartPosition = 'CenterParent'
        $dialog.FormBorderStyle = 'FixedDialog'
        $dialog.MinimizeBox = $false
        $dialog.MaximizeBox = $false
        $dialog.ShowInTaskbar = $false
        $dialog.ClientSize = New-Object System.Drawing.Size(500, 145)
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Question
        $label.Location = New-Object System.Drawing.Point(16, 15)
        $label.Size = New-Object System.Drawing.Size(468, 45)
        $dialog.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(18, 65)
        $textBox.Width = 464
        $dialog.Controls.Add($textBox)
        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'Confirmar'
        $ok.DialogResult = 'OK'
        $ok.Location = New-Object System.Drawing.Point(302, 103)
        $ok.Size = New-Object System.Drawing.Size(85, 30)
        $dialog.Controls.Add($ok)
        $cancel = New-Object System.Windows.Forms.Button
        $cancel.Text = 'Cancelar'
        $cancel.DialogResult = 'Cancel'
        $cancel.Location = New-Object System.Drawing.Point(397, 103)
        $cancel.Size = New-Object System.Drawing.Size(85, 30)
        $dialog.Controls.Add($cancel)
        $dialog.AcceptButton = $ok
        $dialog.CancelButton = $cancel
        $dialog.Add_Shown({ $textBox.Focus() })
        $result = $dialog.ShowDialog($panel.FindForm())
        $value = $textBox.Text
        $dialog.Dispose()
        if ($result -eq 'OK') { return $value }
        return $null
    }.GetNewClosure()

    $showPrompt = {
        param([string]$Prompt)
        $state.DialogOpen = $true
        try {
            if ($Prompt -match 'Escolha \[1/2\]') {
                if ($state.PromptBuffer -match 'Forma de selecao dos projetos') {
                    $value = & $askText 'Digite 1 para selecionar projetos pelos números ou 2 para selecionar por planilha.' 'Forma de seleção dos projetos'
                }
                else {
                    $value = & $askText 'Digite 1 para usuário único ou 2 para lote por planilha Excel.' 'Modo de inclusão'
                }
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'SELECIONAR_CONCESSAO') {
                $value = & $askText 'Informe o número da concessão exibida no acompanhamento.' 'Selecionar concessão'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'FORMA_SELECAO_PROJETOS') {
                $value = & $askText 'Digite 1 para selecionar projetos pelos números ou 2 para selecionar por planilha.' 'Forma de seleção dos projetos'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'SELECIONAR_PROJETOS') {
                $value = & $askText "Informe números separados por vírgula, um intervalo como 1-5, ou digite 'todos'." 'Selecionar projetos'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'SELECIONAR_BUSCAS_SALVAS') {
                $value = & $askText 'Use P ou deixe vazio para as buscas padrão; T para todas; ou informe números como 1,3,6 ou 1-5.' 'Selecionar buscas salvas'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'SELECIONAR_CONCESSAO_BUSCAS') {
                $value = & $askText 'Informe o número da concessão exibida no acompanhamento.' 'Selecionar concessão'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -eq 'SELECIONAR_PROJETOS_BUSCAS') {
                $value = & $askText 'Informe números separados por vírgula, um intervalo como 2-6, ou T para todos.' 'Selecionar projetos'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -match 'E-mail do usuario') {
                $value = & $askText 'Informe o e-mail do usuário que será incluído.' 'Usuário'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -match 'Acao desejada \[1/2/3\]') {
                $value = & $askText 'Digite 1 para incluir, 2 para alterar permissões ou 3 para excluir o participante.' 'Ação desejada'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -match 'Titulo/cargo do participante') {
                $value = & $askText 'Informe o título/cargo do participante. Deixe vazio para usar Participante.' 'Título do participante'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -match 'Numero ou nome da role') {
                $value = & $askText 'Informe o número ou o nome exato da role exibida no acompanhamento.' 'Selecionar role'
                if ($null -ne $value) { & $writeResponse $value }
            }
            elseif ($Prompt -match 'Continuar com este e-mail') {
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'Confirma o e-mail informado?', 'Confirmar usuário', 'YesNo', 'Question')
                & $writeResponse $(if ($yes -eq 'Yes') { 'S' } else { 'N' })
            }
            elseif ($Prompt -match 'Continuar com estes usuarios') {
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'Deseja continuar com os usuários apresentados?', 'Confirmar usuários', 'YesNo', 'Question')
                & $writeResponse $(if ($yes -eq 'Yes') { 'S' } else { 'N' })
            }
            elseif ($Prompt -match 'Aplicar as operacoes no Bentley RBAC') {
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'Deseja aplicar as operações apresentadas no Bentley RBAC?', 'Confirmação final', 'YesNo', 'Warning')
                & $writeResponse $(if ($yes -eq 'Yes') { 'S' } else { 'N' })
            }
            elseif ($Prompt -match 'Aplicar alteracoes\? \[S/N\]') {
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'Deseja aplicar as alterações apresentadas?', 'Confirmação final', 'YesNo', 'Warning')
                & $writeResponse $(if ($yes -eq 'Yes') { 'S' } else { 'N' })
            }
            elseif ($Prompt -match 'Confirmacao:') {
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'A exclusão removerá participantes dos projetos selecionados. Deseja confirmar?', 'Confirmar exclusão', 'YesNo', 'Warning')
                & $writeResponse $(if ($yes -eq 'Yes') { 'CONFIRMAR EXCLUSAO' } else { '' })
            }
            elseif ($Prompt -match '\[S/N\]:') {
                $question = ($Prompt -replace '\s*\[S/N\]:\s*$', '').Trim()
                $yes = [Windows.Forms.MessageBox]::Show($panel.FindForm(), $question, 'Permissão do participante', 'YesNo', 'Question')
                & $writeResponse $(if ($yes -eq 'Yes') { 'S' } else { 'N' })
            }
            else {
                [Windows.Forms.MessageBox]::Show($panel.FindForm(), 'Conclua a etapa indicada e clique em OK para continuar.', 'Continuar execução', 'OK', 'Information') | Out-Null
                & $writeResponse ''
            }
        }
        finally { $state.DialogOpen = $false }
    }.GetNewClosure()

    $sendAnswer = {
        if (-not $state.Process -or $state.Process.HasExited) { return }
        $answer = $answerBox.Text
        try {
            $state.Process.StandardInput.WriteLine($answer)
            $state.Process.StandardInput.Flush()
            & $appendOutput (if ([string]::IsNullOrEmpty($answer)) { '> [CONTINUAR]' } else { "> $answer" })
            $answerBox.Clear()
            $answerBox.Focus()
        }
        catch { & $appendOutput "[ERRO] Não foi possível enviar a resposta: $($_.Exception.Message)" }
    }.GetNewClosure()
    $answerButton.Add_Click($sendAnswer)
    $answerBox.Add_KeyDown({ param($s,$e) if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; & $sendAnswer } }.GetNewClosure())

    $startButton.Add_Click({
        $argumentsForRun = [string]$state.NextRuntimeArguments
        $confirmationForRun = [string]$state.NextConfirmationText
        $state.NextRuntimeArguments = $RuntimeArguments
        $state.NextConfirmationText = $StartConfirmationText
        if (-not [string]::IsNullOrWhiteSpace($confirmationForRun)) {
            $confirmation = [Windows.Forms.MessageBox]::Show(
                $panel.FindForm(),
                $confirmationForRun,
                'Confirmar execução',
                'YesNo',
                'Warning'
            )
            if ($confirmation -ne 'Yes') {
                $statusLabel.Text = 'Execução cancelada pelo operador.'
                return
            }
        }
        $pythonScript = [IO.Path]::GetFullPath((Join-Path $CentralRoot $PythonRelativePath))
        if (-not (Test-Path -LiteralPath $pythonScript -PathType Leaf)) {
            [Windows.Forms.MessageBox]::Show("Script Python não encontrado:`r`n$pythonScript", 'Arquivo não encontrado', 'OK', 'Error') | Out-Null
            return
        }
        try {
            $state.LogPath = Join-Path $logsDir ("{0}_{1}.log" -f $LogPrefix, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $state.Finished = $false
            $output.Clear()
            & $appendOutput "Iniciando rotina: $PanelTitle..."
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            if ($Runtime -eq 'PowerShell') {
                $command = "[Console]::OutputEncoding = [Text.UTF8Encoding]::new(`$false); & '$($pythonScript.Replace("'", "''"))' $argumentsForRun"
                $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand {0}' -f $encodedCommand
            }
            else {
                $psi.FileName = 'python.exe'
                $psi.Arguments = ('-u "{0}" {1}' -f $pythonScript, $argumentsForRun).Trim()
            }
            $psi.WorkingDirectory = Split-Path -Parent $pythonScript
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.RedirectStandardInput = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
            $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            if (-not $process.Start()) { throw 'O processo Python não foi iniciado.' }
            $state.Process = $process
            $state.OutputTask = [CentralScripts.StreamPump]::Start($process.StandardOutput, $state.Queue, '')
            $state.ErrorTask = [CentralScripts.StreamPump]::StartFilteredError($process.StandardError, $state.Queue)
            $startButton.Enabled = $false
            $secondaryButton.Enabled = $false
            $cancelButton.Enabled = $true
            $answerButton.Enabled = $true
            $statusLabel.Text = 'Rotina em execução.'
        }
        catch {
            & $appendOutput "[ERRO] $($_.Exception.Message)"
            $statusLabel.Text = 'Falha ao iniciar.'
        }
    }.GetNewClosure())

    $secondaryButton.Add_Click({
        if ($state.Process -and -not $state.Process.HasExited) { return }
        $state.NextRuntimeArguments = $SecondaryRuntimeArguments
        $state.NextConfirmationText = $SecondaryConfirmationText
        $startButton.PerformClick()
    }.GetNewClosure())

    $cancelButton.Add_Click({
        if ($state.Process -and -not $state.Process.HasExited) {
            if ([Windows.Forms.MessageBox]::Show('Deseja interromper a rotina em execução?', 'Cancelar execução', 'YesNo', 'Warning') -eq 'Yes') {
                $state.Process.Kill()
                & $appendOutput '[AVISO] Cancelamento solicitado pelo operador.'
            }
        }
    }.GetNewClosure())

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 150
    $timer.Add_Tick({
        $line = $null
        $batch = New-Object System.Text.StringBuilder
        while ($state.Queue.TryDequeue([ref]$line)) {
            [void]$batch.Append($line)
            $line = $null
        }
        if ($batch.Length -gt 0) { & $appendProcessOutput $batch.ToString() }
        if ($state.Process -and -not $state.Process.HasExited -and -not $state.DialogOpen) {
            $promptMatch = [regex]::Match(
                $state.PromptBuffer,
                '(?is)(Escolha \[1/2\]:|E-mail do usuario:|Continuar com este e-mail\? \[S/N\]:|Continuar com estes usuarios\? \[S/N\]:|Acao desejada \[1/2/3\]:|Titulo/cargo do participante.*?:|Quando estiver autenticado.*?pressione ENTER no terminal\.\.\.|Numero ou nome da role:|Aplicar as operacoes no Bentley RBAC\? \[S/N\]:|Aplicar alteracoes\? \[S/N\]:|Confirmacao:|[^\r\n]+\[S/N\]:|Pressione ENTER para fechar (?:o navegador|esta janela)(?:\.\.\.)?)\s*$'
            )
            if ($promptMatch.Success) { & $showPrompt $promptMatch.Value }
            elseif (([DateTime]::UtcNow - $state.LastOutputUtc).TotalMilliseconds -ge 700) {
                if ($state.PromptBuffer -match '(?is)Digite numeros separados por virgula.*intervalos como 1-5.*ou.*todos') {
                    & $showPrompt 'SELECIONAR_PROJETOS'
                }
                elseif ($state.PromptBuffer -match '(?is)Projetos encontrados:.*\[\d+\].*') {
                    & $showPrompt 'SELECIONAR_PROJETOS_BUSCAS'
                }
                elseif ($state.PromptBuffer -match '(?is)Concessoes encontradas:.*\[\d+\].*') {
                    & $showPrompt 'SELECIONAR_CONCESSAO_BUSCAS'
                }
                elseif ($state.PromptBuffer -match '(?is)Buscas disponiveis para criar:.*\[\d+\].*') {
                    & $showPrompt 'SELECIONAR_BUSCAS_SALVAS'
                }
                elseif ($state.PromptBuffer -match '(?is)Forma de selecao dos projetos:.*01\..*02\.') {
                    & $showPrompt 'FORMA_SELECAO_PROJETOS'
                }
                elseif ($state.PromptBuffer -match '(?is)Concessoes disponiveis:.*\r?\n\s*\d{2}\.\s+') {
                    & $showPrompt 'SELECIONAR_CONCESSAO'
                }
            }
        }
        if ($state.Process -and $state.Process.HasExited -and -not $state.Finished) {
            $state.Finished = $true
            $code = $state.Process.ExitCode
            & $appendOutput "Processo finalizado com código $code."
            $statusLabel.Text = if ($code -eq 0) { 'Execução concluída.' } else { "Execução finalizada com erro ($code)." }
            $startButton.Enabled = $true; $secondaryButton.Enabled = $true; $cancelButton.Enabled = $false; $answerButton.Enabled = $false
            $state.Process.Dispose(); $state.Process = $null
        }
    }.GetNewClosure())
    $timer.Start()

    $returnToCentral = {
        if ($state.Process -and -not $state.Process.HasExited) {
            [Windows.Forms.MessageBox]::Show('Cancele ou conclua a execução antes de voltar à Central.', 'Rotina em execução', 'OK', 'Information') | Out-Null
            return
        }
        $timer.Stop(); $timer.Dispose(); & $OnBack
    }.GetNewClosure()
    $footerBackButton.Add_Click($returnToCentral)

    return $panel
}
