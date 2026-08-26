[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ExecutorRoot = $PSScriptRoot
$script:CatalogPath = Join-Path $PSScriptRoot 'scripts.json'
$script:LogsPath = Join-Path $PSScriptRoot 'Logs'
$script:Catalog = @()
$script:SelectedScript = $null
$script:ActivePanel = $null
$script:CardControls = @()
$script:LogoImages = @()

function Get-UiImage {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        $sourceImage = [System.Drawing.Image]::FromFile($Path)
        try {
            $copy = New-Object System.Drawing.Bitmap($sourceImage)
            $script:LogoImages += $copy
            return $copy
        }
        finally {
            $sourceImage.Dispose()
        }
    }
    catch {
        Write-ExecutorLog "AVISO | Não foi possível carregar a imagem: $Path | $($_.Exception.Message)"
        return $null
    }
}

function Show-Message {
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$Title = 'Central de Scripts',
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $Icon)
}

function Write-ExecutorLog {
    param([Parameter(Mandatory)][string]$Message)

    if (-not (Test-Path -LiteralPath $script:LogsPath)) {
        New-Item -ItemType Directory -Path $script:LogsPath -Force | Out-Null
    }
    $line = '{0:yyyy-MM-dd HH:mm:ss} | {1}' -f (Get-Date), $Message
    Add-Content -LiteralPath (Join-Path $script:LogsPath ('Executor_{0:yyyyMMdd}.log' -f (Get-Date))) -Value $line -Encoding UTF8
}

function Get-ResolvedScriptPath {
    param([Parameter(Mandatory)]$Item)
    [System.IO.Path]::GetFullPath((Join-Path $script:ExecutorRoot ([string]$Item.arquivo)))
}

function Get-CentralPythonExecutable {
    $projectRoot = [IO.Path]::GetFullPath((Join-Path $script:ExecutorRoot '..\..\..'))
    $venvPython = Join-Path $projectRoot '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $venvPython -PathType Leaf) {
        try {
            $version = & $venvPython --version 2>&1
            if ($LASTEXITCODE -eq 0 -and $version) { return $venvPython }
        }
        catch {}
    }

    $command = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }
    return $null
}

function Test-PythonModule {
    param([Parameter(Mandatory)][string]$Name)
    $python = Get-CentralPythonExecutable
    if (-not $python) { return $false }
    try {
        & $python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('$Name') else 1)" 2>$null
        return $LASTEXITCODE -eq 0
    }
    catch { return $false }
}

function Test-PlaywrightChromium {
    $python = Get-CentralPythonExecutable
    if (-not $python -or -not (Test-PythonModule -Name 'playwright')) { return $false }
    try {
        & $python -c "import os; from playwright.sync_api import sync_playwright; p=sync_playwright().start(); ok=os.path.isfile(p.chromium.executable_path); p.stop(); raise SystemExit(0 if ok else 1)" 2>$null
        return $LASTEXITCODE -eq 0
    }
    catch { return $false }
}

function Get-DependencyStatus {
    param([Parameter(Mandatory)]$Item)

    $results = @()
    foreach ($dependency in @($Item.dependencias)) {
        $parts = ([string]$dependency).Split(':', 2)
        $kind = $parts[0]
        $name = if ($parts.Count -gt 1) { $parts[1] } else { $parts[0] }
        $available = $false

        switch ($kind) {
            'module'  { $available = $null -ne (Get-Module -ListAvailable -Name $name | Select-Object -First 1) }
            'command' { $available = $null -ne (Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1) }
            'pythonmodule' { $available = Test-PythonModule -Name $name }
            'playwrightbrowser' { $available = Test-PlaywrightChromium }
            default   { $available = $false }
        }

        $results += [pscustomobject]@{ Name = $name; Available = $available }
    }
    return @($results)
}

function Get-RiskLabel {
    param([string]$Risk)
    switch ($Risk.ToLowerInvariant()) {
        'critico' { 'CRÍTICO' }
        'alto'    { 'ALTO' }
        'medio'   { 'MÉDIO' }
        default   { 'BAIXO' }
    }
}

function Get-RiskColor {
    param([string]$Risk)
    switch ($Risk.ToLowerInvariant()) {
        'critico' { [System.Drawing.Color]::FromArgb(176, 32, 37) }
        'alto'    { [System.Drawing.Color]::FromArgb(210, 105, 30) }
        'medio'   { [System.Drawing.Color]::FromArgb(183, 139, 0) }
        default   { [System.Drawing.Color]::FromArgb(39, 125, 73) }
    }
}

function Start-CatalogScript {
    param([Parameter(Mandatory)]$Item)

    $target = Get-ResolvedScriptPath -Item $Item
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) {
        Show-Message -Text "O arquivo não foi encontrado:`r`n$target" -Icon Error
        Write-ExecutorLog "ERRO | Arquivo não encontrado | $($Item.id) | $target"
        return
    }

    $missing = @(Get-DependencyStatus -Item $Item | Where-Object { -not $_.Available })
    if ($missing.Count -gt 0) {
        $names = ($missing.Name -join ', ')
        Show-Message -Text "Não é possível iniciar esta rotina. Dependências ausentes:`r`n`r`n$names" -Icon Warning
        Write-ExecutorLog "BLOQUEADO | Dependências ausentes: $names | $($Item.id)"
        return
    }

    if ([string]$Item.risco -in @('alto', 'critico')) {
        $warning = "Esta rotina possui risco $((Get-RiskLabel $Item.risco).ToLowerInvariant()) e pode alterar dados.`r`n`r`nDeseja continuar com '$($Item.nome)'?"
        $answer = [System.Windows.Forms.MessageBox]::Show($warning, 'Confirmar execução', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $workingDirectory = Split-Path -Parent $target
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.WorkingDirectory = $workingDirectory
    $startInfo.UseShellExecute = $true

    switch ([string]$Item.tipo) {
        'powershell' {
            $startInfo.FileName = 'powershell.exe'
            $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -MTA -File "{0}"' -f $target.Replace('"', '\"')
        }
        'batch' {
            $startInfo.FileName = 'cmd.exe'
            $startInfo.Arguments = '/c ""{0}""' -f $target
        }
        'python' {
            $startInfo.FileName = 'python.exe'
            $startInfo.Arguments = '"{0}"' -f $target
        }
        default {
            Show-Message -Text "Tipo de execução não suportado: $($Item.tipo)" -Icon Error
            return
        }
    }

    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
        Write-ExecutorLog "INICIADO | PID $($process.Id) | $($Item.id) | $target"
        $script:StatusLabel.Text = "Rotina iniciada: $($Item.nome) (PID $($process.Id))"
    }
    catch {
        Write-ExecutorLog "ERRO | $($Item.id) | $($_.Exception.Message)"
        Show-Message -Text "Falha ao iniciar a rotina:`r`n`r`n$($_.Exception.Message)" -Icon Error
    }
}

function Open-IntegratedRoutine {
    param([Parameter(Mandatory)]$Item)

    switch ([string]$Item.id) {
        'criar-usuarios' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\01-CriarUsuariosLote\Painel.CriarUsuariosLote.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-CriarUsuariosLotePanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Criar usuários em lote'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'gerenciar-acessos' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\02-GerenciarAcessos\Painel.GerenciarAcessos.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $hostPanel = New-Object System.Windows.Forms.Panel
                $hostPanel.Dock = 'Fill'
                $hostPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
                $form.Controls.Add($hostPanel)
                $hostPanel.BringToFront()
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = $hostPanel
                Open-GerenciarAcessosPanel -Parent $hostPanel -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Gerenciar acessos de projetos'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                if ($hostPanel) {
                    $form.Controls.Remove($hostPanel)
                    $hostPanel.Dispose()
                }
                $script:ActivePanel = $null
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'incluir-usuarios-web' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\03-IncluirUsuariosWeb\Painel.IncluirUsuariosWeb.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-IncluirUsuariosWebPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Incluir usuários em projetos PW Web'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'participantes-pwdm' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\04-GerenciarParticipantesPWDM\Painel.GerenciarParticipantesPWDM.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Controlador visual não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-GerenciarParticipantesPWDMPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Gerenciar participantes PWDM'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'usuarios-inativos' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\05-GerenciarUsuariosInativos\Painel.GerenciarUsuariosInativos.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $hostPanel = New-Object System.Windows.Forms.Panel
                $hostPanel.Dock = 'Fill'
                $hostPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
                $form.Controls.Add($hostPanel)
                $hostPanel.BringToFront()
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = $hostPanel
                Open-GerenciarUsuariosInativosPanel -Parent $hostPanel -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Gerenciar usuários inativos'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                if ($hostPanel) { $form.Controls.Remove($hostPanel); $hostPanel.Dispose() }
                $script:ActivePanel = $null
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'criar-projetos' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\06-CriarProjetos\Painel.CriarProjetos.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-CriarProjetosPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Criar projetos no ProjectWise'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'buscas-salvas' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\07-CriarBuscasSalvas\Painel.CriarBuscasSalvas.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-CriarBuscasSalvasPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Criar buscas salvas'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'entrada-documentos' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\08-EntradaDocumentos\Painel.EntradaDocumentos.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-EntradaDocumentosPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Script Entrada de Documentos'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'movimenta-documentos-validos' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\09-MovimentaDocumentosValidos\Painel.MovimentaDocumentosValidos.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-MovimentaDocumentosValidosPanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Movimenta Documentos Válidos'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        'gera-grd-unidade' {
            $panelScript = Join-Path $script:ExecutorRoot 'Rotinas\10-GeraGRDUnidade\Painel.GeraGRDUnidade.ps1'
            if (-not (Test-Path -LiteralPath $panelScript -PathType Leaf)) {
                Show-Message -Text "Painel integrado não encontrado:`r`n$panelScript" -Icon Error
                return
            }
            try {
                . $panelScript
                $backAction = {
                    if ($script:ActivePanel) {
                        $form.Controls.Remove($script:ActivePanel)
                        $script:ActivePanel.Dispose()
                        $script:ActivePanel = $null
                    }
                    $form.Text = 'Central de Scripts'
                }
                $script:ActivePanel = New-GeraGRDUnidadePanel -Parent $form -CentralRoot $script:ExecutorRoot -OnBack $backAction
                $form.Text = 'Central de Scripts - Gera GRD para Unidade'
                Write-ExecutorLog "PAINEL INTEGRADO | $($Item.id)"
            }
            catch {
                Write-ExecutorLog "ERRO PAINEL | $($Item.id) | $($_.Exception.Message)"
                Show-Message -Text "Não foi possível abrir o painel integrado:`r`n`r`n$($_.Exception.Message)" -Icon Error
            }
        }
        default {
            Show-Message -Text "A rotina '$($Item.nome)' ainda não possui painel integrado." -Icon Information
        }
    }
}

function Invoke-SelectedRoutine {
    if (-not $script:SelectedScript) { return }
    if ([string]$script:SelectedScript.modo -eq 'em_migracao' -or [string]$script:SelectedScript.modo -eq 'integrado') {
        Open-IntegratedRoutine -Item $script:SelectedScript
    }
    else {
        Start-CatalogScript -Item $script:SelectedScript
    }
}

function Update-Details {
    if (-not $script:SelectedScript) {
        $script:RunButton.Enabled = $false
        $logsButton.Enabled = $false
        $folderButton.Enabled = $false
        return
    }

    $item = $script:SelectedScript
    $script:RunButton.Enabled = $true
    $logsButton.Enabled = $true
    $folderButton.Enabled = $true
    $script:RunButton.Text = if ([string]$item.modo -in @('em_migracao', 'integrado')) { 'Abrir formulário' } else { 'Executar rotina' }
    $number = ([int]([int]$item.ordem / 10)).ToString('00')
    $script:StatusLabel.Text = "Selecionado: $number - $($item.nome)"
}

function Update-CardWidths {
    if (-not $script:CardsPanel) { return }
    $cardWidth = [Math]::Max(280, $script:CardsPanel.ClientSize.Width - 28)
    foreach ($card in @($script:CardControls)) {
        $card.Width = $cardWidth
    }
}

function Select-ScriptCard {
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][System.Windows.Forms.Panel]$Card
    )

    foreach ($existingCard in @($script:CardControls)) {
        $existingCard.BackColor = [System.Drawing.Color]::White
    }
    $Card.BackColor = [System.Drawing.Color]::FromArgb(225, 237, 250)
    $script:SelectedScript = $Item
    Update-Details
}

function Update-ScriptList {
    $search = $script:SearchBox.Text.Trim()
    $category = [string]$script:CategoryBox.SelectedItem
    $script:CardsPanel.SuspendLayout()
    foreach ($oldCard in @($script:CardControls)) {
        $script:CardsPanel.Controls.Remove($oldCard)
        $oldCard.Dispose()
    }
    $script:CardControls = @()
    $script:SelectedScript = $null

    $filtered = @($script:Catalog | Where-Object {
        ($category -eq 'Todas' -or $_.categoria -eq $category) -and
        ([string]::IsNullOrWhiteSpace($search) -or $_.nome -like "*$search*" -or $_.descricao -like "*$search*")
    } | Sort-Object ordem, nome)

    foreach ($item in $filtered) {
        $card = New-Object System.Windows.Forms.Panel
        $card.Height = 92
        $card.Width = [Math]::Max(280, $script:CardsPanel.ClientSize.Width - 28)
        $card.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 8)
        $card.Padding = New-Object System.Windows.Forms.Padding(12, 9, 12, 8)
        $card.BackColor = [System.Drawing.Color]::White
        $card.BorderStyle = 'FixedSingle'
        $card.Cursor = [System.Windows.Forms.Cursors]::Hand
        $card.Tag = $item

        $number = ([int]([int]$item.ordem / 10)).ToString('00')
        $cardTitle = New-Object System.Windows.Forms.Label
        $cardTitle.Text = "$number - $($item.nome)"
        $cardTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
        $cardTitle.ForeColor = [System.Drawing.Color]::FromArgb(22, 54, 92)
        $cardTitle.Location = New-Object System.Drawing.Point(12, 10)
        $cardTitle.Size = New-Object System.Drawing.Size(($card.Width - 26), 26)
        $cardTitle.Anchor = 'Top,Left,Right'
        $cardTitle.Cursor = [System.Windows.Forms.Cursors]::Hand
        $card.Controls.Add($cardTitle)

        $cardDescription = New-Object System.Windows.Forms.Label
        $cardDescription.Text = [string]$item.descricao
        $cardDescription.ForeColor = [System.Drawing.Color]::FromArgb(75, 82, 90)
        $cardDescription.Location = New-Object System.Drawing.Point(13, 42)
        $cardDescription.Size = New-Object System.Drawing.Size(($card.Width - 28), 39)
        $cardDescription.Anchor = 'Top,Left,Right'
        $cardDescription.Cursor = [System.Windows.Forms.Cursors]::Hand
        $card.Controls.Add($cardDescription)

        $selectHandler = { Select-ScriptCard -Item $item -Card $card }.GetNewClosure()
        $openHandler = {
            Select-ScriptCard -Item $item -Card $card
            Invoke-SelectedRoutine
        }.GetNewClosure()
        foreach ($control in @($card, $cardTitle, $cardDescription)) {
            $control.Add_Click($selectHandler)
            $control.Add_DoubleClick($openHandler)
        }

        $script:CardControls += $card
        [void]$script:CardsPanel.Controls.Add($card)
    }
    $script:CardsPanel.ResumeLayout()
    Update-CardWidths
    Update-Details
    $script:StatusLabel.Text = if ($filtered.Count -eq 1) { '1 rotina encontrada' } else { "$($filtered.Count) rotinas encontradas" }
}

try {
    if (-not (Test-Path -LiteralPath $script:CatalogPath)) { throw "Catálogo não encontrado: $script:CatalogPath" }
    $catalogDocument = Get-Content -LiteralPath $script:CatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $script:Catalog = @($catalogDocument.scripts)
    if ($script:Catalog.Count -eq 0) { throw 'O catálogo não possui scripts cadastrados.' }
}
catch {
    Show-Message -Text "Não foi possível carregar o catálogo:`r`n`r`n$($_.Exception.Message)" -Icon Error
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Central de Scripts'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.MinimumSize = New-Object System.Drawing.Size(920, 600)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.Margin = New-Object System.Windows.Forms.Padding(0)
$layout.Padding = New-Object System.Windows.Forms.Padding(0)
$layout.ColumnCount = 1
$layout.RowCount = 4
[void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 86)))
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 62)))
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24)))
$form.Controls.Add($layout)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.Margin = New-Object System.Windows.Forms.Padding(0)
$header.BackColor = [System.Drawing.Color]::FromArgb(22, 54, 92)
$layout.Controls.Add($header, 0, 0)

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = 'Fill'
$headerLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$headerLayout.ColumnCount = 3
$headerLayout.RowCount = 1
[void]$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 215)))
[void]$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 225)))
[void]$headerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$header.Controls.Add($headerLayout)

$projectWiseLogo = New-Object System.Windows.Forms.PictureBox
$projectWiseLogo.Dock = 'Fill'
$projectWiseLogo.Margin = New-Object System.Windows.Forms.Padding(14, 10, 10, 10)
$projectWiseLogo.SizeMode = 'Zoom'
$projectWiseLogo.BackColor = [System.Drawing.Color]::Transparent
$projectWiseLogo.AccessibleName = 'Logotipo ProjectWise'
$projectWiseLogo.Image = Get-UiImage -Path (Join-Path $script:ExecutorRoot 'Outlook-ProjectWis.png')
$headerLayout.Controls.Add($projectWiseLogo, 0, 0)

$headerText = New-Object System.Windows.Forms.Panel
$headerText.Dock = 'Fill'
$headerText.Margin = New-Object System.Windows.Forms.Padding(0)
$headerText.BackColor = [System.Drawing.Color]::Transparent
$headerLayout.Controls.Add($headerText, 1, 0)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Central de Scripts'
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 21)
$title.AutoSize = $false
$title.Dock = 'Fill'
$title.TextAlign = 'MiddleCenter'
$headerText.Controls.Add($title)

$ecorodoviasLogo = New-Object System.Windows.Forms.PictureBox
$ecorodoviasLogo.Dock = 'Fill'
$ecorodoviasLogo.Margin = New-Object System.Windows.Forms.Padding(12, 13, 16, 13)
$ecorodoviasLogo.SizeMode = 'Zoom'
$ecorodoviasLogo.BackColor = [System.Drawing.Color]::Transparent
$ecorodoviasLogo.AccessibleName = 'Logotipo EcoRodovias'
$ecorodoviasLogo.Image = Get-UiImage -Path (Join-Path $script:ExecutorRoot 'Ecorodovias_Logo.png')
$headerLayout.Controls.Add($ecorodoviasLogo, 2, 0)

$filters = New-Object System.Windows.Forms.Panel
$filters.Dock = 'Fill'
$filters.Margin = New-Object System.Windows.Forms.Padding(0)
$filters.Padding = New-Object System.Windows.Forms.Padding(16, 14, 16, 8)
$layout.Controls.Add($filters, 0, 1)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = 'Pesquisar:'
$searchLabel.AutoSize = $true
$searchLabel.Location = New-Object System.Drawing.Point(18, 21)
$filters.Controls.Add($searchLabel)

$script:SearchBox = New-Object System.Windows.Forms.TextBox
$script:SearchBox.Location = New-Object System.Drawing.Point(88, 17)
$script:SearchBox.Width = 330
$filters.Controls.Add($script:SearchBox)

$categoryLabel = New-Object System.Windows.Forms.Label
$categoryLabel.Text = 'Categoria:'
$categoryLabel.AutoSize = $true
$categoryLabel.Location = New-Object System.Drawing.Point(442, 21)
$filters.Controls.Add($categoryLabel)

$script:CategoryBox = New-Object System.Windows.Forms.ComboBox
$script:CategoryBox.DropDownStyle = 'DropDownList'
$script:CategoryBox.Location = New-Object System.Drawing.Point(512, 17)
$script:CategoryBox.Width = 220
[void]$script:CategoryBox.Items.Add('Todas')
$script:Catalog | Select-Object -ExpandProperty categoria -Unique | Sort-Object | ForEach-Object { [void]$script:CategoryBox.Items.Add($_) }
$script:CategoryBox.SelectedIndex = 0
$filters.Controls.Add($script:CategoryBox)

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.Margin = New-Object System.Windows.Forms.Padding(0)
$main.Padding = New-Object System.Windows.Forms.Padding(16, 6, 16, 10)
$main.ColumnCount = 1
$main.RowCount = 2
[void]$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
$layout.Controls.Add($main, 0, 2)

$script:CardsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:CardsPanel.Dock = 'Fill'
$script:CardsPanel.AutoScroll = $true
$script:CardsPanel.FlowDirection = 'TopDown'
$script:CardsPanel.WrapContents = $false
$script:CardsPanel.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 247)
$script:CardsPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$main.Controls.Add($script:CardsPanel, 0, 0)

$homeActions = New-Object System.Windows.Forms.FlowLayoutPanel
$homeActions.Dock = 'Fill'
$homeActions.FlowDirection = 'LeftToRight'
$homeActions.WrapContents = $false
$homeActions.Padding = New-Object System.Windows.Forms.Padding(4, 7, 4, 4)
$homeActions.BackColor = [System.Drawing.Color]::White
$main.Controls.Add($homeActions, 0, 1)

$script:RunButton = New-Object System.Windows.Forms.Button
$script:RunButton.Text = 'Executar rotina'
$script:RunButton.Enabled = $false
$script:RunButton.BackColor = [System.Drawing.Color]::FromArgb(31, 105, 178)
$script:RunButton.ForeColor = [System.Drawing.Color]::White
$script:RunButton.FlatStyle = 'Flat'
$script:RunButton.Margin = New-Object System.Windows.Forms.Padding(3)
$script:RunButton.Size = New-Object System.Drawing.Size(160, 38)
$homeActions.Controls.Add($script:RunButton)

$logsButton = New-Object System.Windows.Forms.Button
$logsButton.Text = 'Abrir pasta de logs'
$logsButton.Enabled = $false
$logsButton.Margin = New-Object System.Windows.Forms.Padding(3)
$logsButton.Size = New-Object System.Drawing.Size(150, 38)
$homeActions.Controls.Add($logsButton)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = 'Abrir pasta do script'
$folderButton.Enabled = $false
$folderButton.Margin = New-Object System.Windows.Forms.Padding(3)
$folderButton.Size = New-Object System.Drawing.Size(170, 38)
$homeActions.Controls.Add($folderButton)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = 'Fill'
$statusStrip.Margin = New-Object System.Windows.Forms.Padding(0)
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Spring = $true
$script:StatusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($script:StatusLabel)
$layout.Controls.Add($statusStrip, 0, 3)

$script:SearchBox.Add_TextChanged({ Update-ScriptList })
$script:CategoryBox.Add_SelectedIndexChanged({ Update-ScriptList })
$script:CardsPanel.Add_Resize({ Update-CardWidths })
$script:RunButton.Add_Click({ Invoke-SelectedRoutine })
$logsButton.Add_Click({
    if (-not (Test-Path -LiteralPath $script:LogsPath)) { New-Item -ItemType Directory -Path $script:LogsPath -Force | Out-Null }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:LogsPath)
})
$folderButton.Add_Click({
    if ($script:SelectedScript) {
        $folder = Split-Path -Parent (Get-ResolvedScriptPath -Item $script:SelectedScript)
        if (Test-Path -LiteralPath $folder) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $folder) }
    }
})

$form.Add_Shown({
    Update-ScriptList
    Write-ExecutorLog "ABERTO | Catálogo com $($script:Catalog.Count) rotinas"
})
$form.Add_FormClosing({
    param($sender, $eventArgs)
    $activeProcess = if ($script:ActivePanel -and $script:ActivePanel.Tag) { $script:ActivePanel.Tag.Process } else { $null }
    if ($activeProcess -and -not $activeProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            $form,
            'Existe uma rotina em execução. Fechar a Central também interromperá o processamento. Deseja realmente sair?',
            'Rotina em execução',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $eventArgs.Cancel = $true
            $form.Activate()
            $form.BringToFront()
        }
    }
})
$form.Add_FormClosed({
    Write-ExecutorLog 'FECHADO'
    foreach ($logoImage in @($script:LogoImages)) {
        if ($logoImage) { $logoImage.Dispose() }
    }
    $script:LogoImages = @()
})

[void]$form.ShowDialog()
