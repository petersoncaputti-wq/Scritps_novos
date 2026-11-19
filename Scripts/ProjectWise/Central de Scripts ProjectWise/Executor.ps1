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

# Tema visual (cores do brand book, fontes Lexend, desenho dos cartoes)
. (Join-Path $PSScriptRoot 'Tema.ps1')
Initialize-TemaFontes -Raiz $PSScriptRoot

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

    # Valida o executavel para nao aceitar o alias da Microsoft Store.
    $launcher = Get-Command py.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($launcher) {
        try {
            $resolved = & $launcher.Source -c "import sys; print(sys.executable)" 2>$null
            if ($LASTEXITCODE -eq 0 -and $resolved -and (Test-Path -LiteralPath ([string]$resolved) -PathType Leaf)) {
                return [string]$resolved
            }
        }
        catch {}
    }
    foreach ($command in @(Get-Command python.exe, python3.exe -CommandType Application -ErrorAction SilentlyContinue)) {
        if ($command.Source -match '\\Microsoft\\WindowsApps\\') { continue }
        try {
            $version = & $command.Source --version 2>&1
            if ($LASTEXITCODE -eq 0 -and $version) { return $command.Source }
        }
        catch {}
    }
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
            'command' { if ($name -eq 'python') { $available = $null -ne (Get-CentralPythonExecutable) } else { $available = $null -ne (Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1) } }
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

function Get-RoutineNumber {
    param($Item)
    ([int]([int]$Item.ordem / 10)).ToString('00')
}

function Get-RoutineLogsPath {
    # Prefere a pasta de logs da propria rotina (Logs\NN-Nome); senao, a da Central.
    param($Item)
    $numero = Get-RoutineNumber -Item $Item
    $pasta = Get-ChildItem -LiteralPath $script:LogsPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$numero-*" } | Select-Object -First 1
    if ($pasta) { return $pasta.FullName }
    return $script:LogsPath
}

function Update-Details {
    $item = $script:SelectedScript
    $temItem = $null -ne $item
    $script:RunButton.Enabled = $temItem
    $script:LogsButton.Enabled = $temItem
    $script:MoreButton.Enabled = $temItem

    if (-not $temItem) {
        $script:DetailTitle.Text = 'Selecione uma rotina'
        $script:DetailInfo.Text = 'Clique em um cartão. Duplo clique ou Enter abre a rotina.'
        $script:RunButton.Text = 'Abrir rotina'
        return
    }

    $numero = Get-RoutineNumber -Item $item
    $script:DetailTitle.Text = "$numero · $($item.nome)"
    $partes = @([string]$item.categoria)
    switch ([string]$item.risco) {
        'critico' { $partes += 'Crítica: altera ou exclui dados' }
        'alto'    { $partes += 'Pede confirmação' }
    }
    if ([string]$item.modo -eq 'em_migracao') { $partes += 'Em migração' }
    $script:DetailInfo.Text = $partes -join '  ·  '
    $script:RunButton.Text = if ([string]$item.modo -in @('em_migracao', 'integrado')) { 'Abrir rotina' } else { 'Executar rotina' }
    $script:StatusLabel.Text = "Selecionado: $numero - $($item.nome)"
}

function Update-CardWidths {
    if (-not $script:CardsPanel) { return }
    $util = $script:CardsPanel.ClientSize.Width - $script:CardsPanel.Padding.Horizontal - 2
    if ($script:CardsPanel.VerticalScroll.Visible) { $util -= 2 }
    $colunas = if ($util -ge 760) { 2 } else { 1 }
    $margem = 12
    $largura = [Math]::Max(300, [int][Math]::Floor(($util - ($colunas * $margem)) / $colunas))
    $script:CardsPanel.SuspendLayout()
    foreach ($card in @($script:CardControls)) {
        $card.Width = $largura
        $card.Invalidate()
    }
    $script:CardsPanel.ResumeLayout()
}

function Select-ScriptCard {
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Card
    )

    foreach ($existingCard in @($script:CardControls)) {
        if ($existingCard.Tag.Selecionado) {
            $existingCard.Tag.Selecionado = $false
            $existingCard.Invalidate()
        }
    }
    $Card.Tag.Selecionado = $true
    $Card.Invalidate()
    $script:CardsPanel.ScrollControlIntoView($Card)
    $script:SelectedScript = $Item
    Update-Details
}

function Move-CardSelection {
    param([int]$Passo)
    $cards = @($script:CardControls)
    if ($cards.Count -eq 0) { return }
    $atual = -1
    for ($i = 0; $i -lt $cards.Count; $i++) { if ($cards[$i].Tag.Selecionado) { $atual = $i; break } }
    $novo = [Math]::Min($cards.Count - 1, [Math]::Max(0, $atual + $Passo))
    if ($atual -lt 0) { $novo = 0 }
    Select-ScriptCard -Item $cards[$novo].Tag.Item -Card $cards[$novo]
}

# --- Desenho do cartao (tudo pintado, para cantos e selos suaves) ---
$script:CardPaint = {
    param($ctl, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $estado = $ctl.Tag
    $item = $estado.Item
    $T = $script:Tema
    $F = $script:Fontes

    $g.Clear($script:CardsPanel.BackColor)
    $r = New-Object System.Drawing.RectangleF(0.5, 0.5, ($ctl.Width - 1.5), ($ctl.Height - 1.5))
    $path = New-TemaRetanguloArredondado -R $r -Raio 8
    $fundo = New-Object System.Drawing.SolidBrush($T.Camada)
    $corBorda = if ($estado.Selecionado) { $T.VerdeMedio } elseif ($estado.Hover) { $T.ContornoForte } else { $T.Contorno }
    $espessura = if ($estado.Selecionado) { 2 } else { 1 }
    $borda = New-Object System.Drawing.Pen($corBorda, $espessura)
    $g.FillPath($fundo, $path)
    $g.DrawPath($borda, $path)
    $fundo.Dispose(); $borda.Dispose(); $path.Dispose()

    # Icone dentro da forma "janela para o mundo"
    $corJanela = if ($estado.Selecionado) { $T.VerdeMedio } else { $T.VerdeSuave }
    $corGlifo = if ($estado.Selecionado) { [System.Drawing.Color]::White } else { $T.VerdeEscuro }
    Draw-TemaJanela -G $g -R (New-Object System.Drawing.RectangleF(16, 16, 50, 44)) -Cor $corJanela
    $glifo = if ($item.icone) { [string][char][Convert]::ToInt32([string]$item.icone, 16) } else { [string][char]0xE7C3 }
    $flagsCentro = [System.Windows.Forms.TextFormatFlags]'HorizontalCenter,VerticalCenter,NoPadding'
    [System.Windows.Forms.TextRenderer]::DrawText($g, $glifo, $F.Icone, (New-Object System.Drawing.Rectangle(22, 22, 44, 38)), $corGlifo, $flagsCentro)

    # Textos
    $x = 82
    $larguraTexto = $ctl.Width - $x - 16
    $numero = Get-RoutineNumber -Item $item
    $flagsLinha = [System.Windows.Forms.TextFormatFlags]'NoPadding,SingleLine,EndEllipsis'
    $tamNum = [System.Windows.Forms.TextRenderer]::MeasureText($g, $numero, $F.Legenda, (New-Object System.Drawing.Size(200, 40)), $flagsLinha)
    [System.Windows.Forms.TextRenderer]::DrawText($g, $numero, $F.Legenda, (New-Object System.Drawing.Point(($ctl.Width - 16 - $tamNum.Width), 18)), $T.TextoSuave, $flagsLinha)
    [System.Windows.Forms.TextRenderer]::DrawText($g, [string]$item.nome, $F.CardTitulo, (New-Object System.Drawing.Rectangle($x, 15, ($larguraTexto - $tamNum.Width - 8), 24)), $T.Texto, $flagsLinha)
    $flagsDesc = [System.Windows.Forms.TextFormatFlags]'NoPadding,WordBreak,EndEllipsis'
    [System.Windows.Forms.TextRenderer]::DrawText($g, [string]$item.descricao, $F.Texto, (New-Object System.Drawing.Rectangle($x, 41, $larguraTexto, 36)), $T.TextoSuave, $flagsDesc)

    # Selos
    $sx = $x
    $sy = $ctl.Height - 31
    $sx += 6 + (Draw-TemaSelo -G $g -Texto ([string]$item.categoria) -Fonte $F.Legenda -X $sx -Y $sy -Fundo $T.Fundo -Cor $T.TextoSuave)
    switch ([string]$item.risco) {
        'critico' { $sx += 6 + (Draw-TemaSelo -G $g -Texto 'Crítica' -Fonte $F.LegendaForte -X $sx -Y $sy -Fundo $T.CriticoSuave -Cor $T.Critico -Ponto) }
        'alto'    { $sx += 6 + (Draw-TemaSelo -G $g -Texto 'Pede confirmação' -Fonte $F.Legenda -X $sx -Y $sy -Fundo $T.Fundo -Cor $T.TextoSuave) }
    }
    if ([string]$item.modo -eq 'em_migracao') {
        [void](Draw-TemaSelo -G $g -Texto 'Em migração' -Fonte $F.Legenda -X $sx -Y $sy -Fundo $T.AtencaoSuave -Cor $T.Atencao -Ponto)
    }
}

$script:CardMouseEnter = { $this.Tag.Hover = $true; $this.Invalidate() }
$script:CardMouseLeave = { $this.Tag.Hover = $false; $this.Invalidate() }
$script:CardClick = { Select-ScriptCard -Item $this.Tag.Item -Card $this; $this.Focus() }
$script:CardDoubleClick = {
    Select-ScriptCard -Item $this.Tag.Item -Card $this
    Invoke-SelectedRoutine
}

# --- Filtros de categoria em formato de pilula ---
$script:PillPaint = {
    param($ctl, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $T = $script:Tema
    $ativo = $ctl.Tag.Categoria -eq $script:CategoriaAtual
    $g.Clear($ctl.Parent.BackColor)
    $r = New-Object System.Drawing.RectangleF(0.5, 0.5, ($ctl.Width - 1.5), ($ctl.Height - 1.5))
    $path = New-TemaRetanguloArredondado -R $r -Raio (($ctl.Height - 1) / 2)
    if ($ativo) {
        $b = New-Object System.Drawing.SolidBrush($T.VerdeEscuro); $g.FillPath($b, $path); $b.Dispose()
        $corTexto = [System.Drawing.Color]::White
    }
    else {
        $corFundo = if ($ctl.Tag.Hover) { $T.VerdeSuave } else { $T.Camada }
        $b = New-Object System.Drawing.SolidBrush($corFundo); $g.FillPath($b, $path); $b.Dispose()
        $p = New-Object System.Drawing.Pen($T.Contorno); $g.DrawPath($p, $path); $p.Dispose()
        $corTexto = $T.Texto
    }
    $path.Dispose()
    $flags = [System.Windows.Forms.TextFormatFlags]'HorizontalCenter,VerticalCenter,NoPadding,SingleLine'
    [System.Windows.Forms.TextRenderer]::DrawText($g, $ctl.Tag.Texto, $script:Fontes.Pilula, $ctl.ClientRectangle, $corTexto, $flags)
}

function New-CategoryPill {
    param([string]$Categoria, [string]$Texto)
    $pill = New-Object System.Windows.Forms.Panel
    Enable-TemaBufferDuplo -Controle $pill
    $tam = [System.Windows.Forms.TextRenderer]::MeasureText($Texto, $script:Fontes.Pilula)
    $pill.Size = New-Object System.Drawing.Size(($tam.Width + 26), 32)
    $pill.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $pill.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pill.Tag = @{ Categoria = $Categoria; Texto = $Texto; Hover = $false }
    $pill.Add_Paint($script:PillPaint)
    $pill.Add_MouseEnter({ $this.Tag.Hover = $true; $this.Invalidate() })
    $pill.Add_MouseLeave({ $this.Tag.Hover = $false; $this.Invalidate() })
    $pill.Add_Click({
        $script:CategoriaAtual = $this.Tag.Categoria
        foreach ($p in $script:CategoryPills) { $p.Invalidate() }
        Update-ScriptList
    })
    return $pill
}

function Update-ScriptList {
    $search = $script:SearchBox.Text.Trim()
    $category = $script:CategoriaAtual
    $anterior = if ($script:SelectedScript) { [string]$script:SelectedScript.id } else { $null }

    $script:CardsPanel.SuspendLayout()
    foreach ($oldCard in @($script:CardControls)) {
        $script:CardsPanel.Controls.Remove($oldCard)
        $oldCard.Dispose()
    }
    $script:CardControls = @()
    $script:SelectedScript = $null

    $filtered = @($script:Catalog | Where-Object {
        ([string]::IsNullOrWhiteSpace($category) -or $_.categoria -eq $category) -and
        ([string]::IsNullOrWhiteSpace($search) -or $_.nome -like "*$search*" -or $_.descricao -like "*$search*" -or $_.categoria -like "*$search*")
    } | Sort-Object ordem, nome)

    foreach ($item in $filtered) {
        $card = New-Object System.Windows.Forms.Panel
        Enable-TemaBufferDuplo -Controle $card
        $card.Height = 116
        $card.Width = 320
        $card.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 12)
        $card.Cursor = [System.Windows.Forms.Cursors]::Hand
        $card.TabStop = $true
        $card.AccessibleName = "$(Get-RoutineNumber -Item $item) - $($item.nome)"
        $card.Tag = @{ Item = $item; Hover = $false; Selecionado = $false }
        $card.Add_Paint($script:CardPaint)
        $card.Add_MouseEnter($script:CardMouseEnter)
        $card.Add_MouseLeave($script:CardMouseLeave)
        $card.Add_Click($script:CardClick)
        $card.Add_DoubleClick($script:CardDoubleClick)
        $script:CardControls += $card
        [void]$script:CardsPanel.Controls.Add($card)
    }

    $script:EmptyLabel.Visible = $filtered.Count -eq 0
    if ($filtered.Count -eq 0) {
        $script:EmptyLabel.Text = if ($search) { "Nenhuma rotina encontrada para '$search'." } else { 'Nenhuma rotina nesta categoria.' }
    }
    $script:CardsPanel.ResumeLayout()
    Update-CardWidths

    $reselecionar = @($script:CardControls | Where-Object { [string]$_.Tag.Item.id -eq $anterior } | Select-Object -First 1)
    if ($reselecionar.Count -gt 0) { Select-ScriptCard -Item $reselecionar[0].Tag.Item -Card $reselecionar[0] } else { Update-Details }
    $script:StatusLabel.Text = if ($filtered.Count -eq 1) { '1 rotina' } else { "$($filtered.Count) rotinas" }
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

$T = $script:Tema
$script:Fontes = @{
    Titulo       = Get-TemaFonte -Papel Titulo -Tamanho 15
    Marca        = Get-TemaFonte -Papel Texto -Tamanho 9
    CardTitulo   = Get-TemaFonte -Papel Destaque -Tamanho 10.5
    Detalhe      = Get-TemaFonte -Papel Destaque -Tamanho 11
    Texto        = Get-TemaFonte -Papel Texto -Tamanho 9
    Legenda      = Get-TemaFonte -Papel Legenda -Tamanho 8
    LegendaForte = Get-TemaFonte -Papel TextoForte -Tamanho 8
    Pilula       = Get-TemaFonte -Papel Texto -Tamanho 9
    Icone        = Get-TemaFonte -Papel Icone -Tamanho 17
    IconeBusca   = Get-TemaFonte -Papel Icone -Tamanho 10
    IconeBotao   = Get-TemaFonte -Papel Icone -Tamanho 11
}
$script:CategoriaAtual = ''
$script:CategoryPills = @()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Central de Scripts'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1180, 760)
# Minimo suficiente para as telas das rotinas (as maiores pedem ~1000x640 de area util).
$form.MinimumSize = New-Object System.Drawing.Size(1040, 700)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
# Fonte e fundo do form ficam como antes: os paineis das rotinas herdam esses valores.
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.KeyPreview = $true
$iconePath = Join-Path $script:ExecutorRoot 'CentralScripts-Icone-v2.ico'
if (Test-Path -LiteralPath $iconePath) {
    try { $form.Icon = New-Object System.Drawing.Icon($iconePath) } catch { }
}
Set-TemaJanela -Form $form

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.Margin = New-Object System.Windows.Forms.Padding(0)
$layout.Padding = New-Object System.Windows.Forms.Padding(0)
$layout.BackColor = $T.Fundo
$layout.ColumnCount = 1
$layout.RowCount = 5
[void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 68)))   # cabecalho
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))   # busca e filtros
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))   # cartoes
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 68)))   # acoes
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))   # status
$form.Controls.Add($layout)

# ---------------- Cabecalho ----------------
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Fill'
$header.Margin = New-Object System.Windows.Forms.Padding(0)
$header.Padding = New-Object System.Windows.Forms.Padding(20, 0, 24, 0)
$header.BackColor = $T.Camada
$header.ColumnCount = 3
$header.RowCount = 1
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$header.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$header.Add_Paint({
    param($ctl, $e)
    $p = New-Object System.Drawing.Pen($script:Tema.Contorno)
    $e.Graphics.DrawLine($p, 0, ($ctl.Height - 1), $ctl.Width, ($ctl.Height - 1))
    $p.Dispose()
})
$layout.Controls.Add($header, 0, 0)

$appIcon = New-Object System.Windows.Forms.PictureBox
$appIcon.Size = New-Object System.Drawing.Size(40, 40)
$appIcon.Anchor = 'Left'
$appIcon.Margin = New-Object System.Windows.Forms.Padding(0)
$appIcon.SizeMode = 'Zoom'
$appIcon.AccessibleName = 'Ícone da Central de Scripts'
# Usa o PNG v1 (transparencia limpa; o v2 tem halo magenta) reduzido em alta qualidade.
$iconeOrigem = Get-UiImage -Path (Join-Path $script:ExecutorRoot 'CentralScripts-Icone.png')
if ($iconeOrigem) {
    $iconeReduzido = New-Object System.Drawing.Bitmap(80, 80)
    $gIcone = [System.Drawing.Graphics]::FromImage($iconeReduzido)
    $gIcone.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $gIcone.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $gIcone.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $gIcone.DrawImage($iconeOrigem, 0, 0, 80, 80)
    $gIcone.Dispose()
    $script:LogoImages += $iconeReduzido
    $appIcon.Image = $iconeReduzido
}
$header.Controls.Add($appIcon, 0, 0)

$headerText = New-Object System.Windows.Forms.FlowLayoutPanel
$headerText.FlowDirection = 'TopDown'
$headerText.WrapContents = $false
$headerText.AutoSize = $true
$headerText.Anchor = 'Left'
$headerText.Margin = New-Object System.Windows.Forms.Padding(4, 0, 0, 0)
$header.Controls.Add($headerText, 1, 0)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Central de Scripts'
$title.Font = $script:Fontes.Titulo
$title.ForeColor = $T.VerdeEscuro
$title.AutoSize = $true
$title.Margin = New-Object System.Windows.Forms.Padding(0)
$headerText.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'Rotinas ProjectWise · Engenharia'
$subtitle.Font = $script:Fontes.Marca
$subtitle.ForeColor = $T.TextoSuave
$subtitle.AutoSize = $true
$subtitle.Margin = New-Object System.Windows.Forms.Padding(1, 0, 0, 0)
$headerText.Controls.Add($subtitle)

$ecorodoviasLogo = New-Object System.Windows.Forms.PictureBox
$ecorodoviasLogo.Size = New-Object System.Drawing.Size(124, 40)
$ecorodoviasLogo.Anchor = 'Right'
$ecorodoviasLogo.Margin = New-Object System.Windows.Forms.Padding(0)
$ecorodoviasLogo.SizeMode = 'Zoom'
$ecorodoviasLogo.AccessibleName = 'Logotipo EcoRodovias'
$ecorodoviasLogo.Image = Get-UiImage -Path (Join-Path $script:ExecutorRoot 'Ecorodovias_Logo.png')
$header.Controls.Add($ecorodoviasLogo, 2, 0)

# ---------------- Busca e filtros ----------------
$filters = New-Object System.Windows.Forms.FlowLayoutPanel
$filters.Dock = 'Fill'
$filters.WrapContents = $false
$filters.Margin = New-Object System.Windows.Forms.Padding(0)
$filters.Padding = New-Object System.Windows.Forms.Padding(20, 16, 20, 0)
$filters.BackColor = $T.Fundo
$layout.Controls.Add($filters, 0, 1)

$searchHost = New-Object System.Windows.Forms.Panel
Enable-TemaBufferDuplo -Controle $searchHost
$searchHost.Size = New-Object System.Drawing.Size(320, 32)
$searchHost.Margin = New-Object System.Windows.Forms.Padding(0, 0, 18, 0)
$searchHost.BackColor = $T.Fundo
$searchHost.Add_Paint({
    param($ctl, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $r = New-Object System.Drawing.RectangleF(0.5, 0.5, ($ctl.Width - 1.5), ($ctl.Height - 1.5))
    $path = New-TemaRetanguloArredondado -R $r -Raio 6
    $b = New-Object System.Drawing.SolidBrush($script:Tema.Camada); $g.FillPath($b, $path); $b.Dispose()
    $foco = $script:SearchBox -and $script:SearchBox.Focused
    $cor = if ($foco) { $script:Tema.VerdeMedio } else { $script:Tema.Contorno }
    $p = New-Object System.Drawing.Pen($cor, $(if ($foco) { 1.5 } else { 1 })); $g.DrawPath($p, $path); $p.Dispose(); $path.Dispose()
    $flags = [System.Windows.Forms.TextFormatFlags]'VerticalCenter,NoPadding'
    [System.Windows.Forms.TextRenderer]::DrawText($g, [string][char]0xE721, $script:Fontes.IconeBusca, (New-Object System.Drawing.Rectangle(11, 0, 20, $ctl.Height)), $script:Tema.TextoSuave, $flags)
})
$searchHost.Add_Click({ $script:SearchBox.Focus() })
$filters.Controls.Add($searchHost)

$script:SearchBox = New-Object System.Windows.Forms.TextBox
$script:SearchBox.BorderStyle = 'None'
$script:SearchBox.Font = $script:Fontes.Texto
$script:SearchBox.BackColor = $T.Camada
$script:SearchBox.ForeColor = $T.Texto
$script:SearchBox.Location = New-Object System.Drawing.Point(34, 8)
$script:SearchBox.Width = 276
$script:SearchBox.Anchor = 'Left,Right,Top'
Set-TemaTextoExemplo -Caixa $script:SearchBox -Texto 'Buscar rotina (Ctrl+F)'
$script:SearchBox.Add_GotFocus({ $searchHost.Invalidate() })
$script:SearchBox.Add_LostFocus({ $searchHost.Invalidate() })
$searchHost.Controls.Add($script:SearchBox)

$categorias = @($script:Catalog | Group-Object categoria | Sort-Object Name)
$script:CategoryPills += New-CategoryPill -Categoria '' -Texto "Todas  $($script:Catalog.Count)"
foreach ($grupo in $categorias) {
    $script:CategoryPills += New-CategoryPill -Categoria $grupo.Name -Texto "$($grupo.Name)  $($grupo.Count)"
}
foreach ($pill in $script:CategoryPills) { $filters.Controls.Add($pill) }

# Em janelas estreitas a busca encolhe para os filtros continuarem visiveis.
$filters.Add_Resize({
    $larguraFiltros = 0
    foreach ($p in $script:CategoryPills) { $larguraFiltros += $p.Width + $p.Margin.Horizontal }
    $disponivel = $this.ClientSize.Width - $this.Padding.Horizontal - $larguraFiltros - $searchHost.Margin.Horizontal - 4
    $searchHost.Width = [Math]::Max(180, [Math]::Min(320, $disponivel))
    $script:SearchBox.Width = $searchHost.Width - 44
})

# ---------------- Cartoes ----------------
$cardsHost = New-Object System.Windows.Forms.Panel
$cardsHost.Dock = 'Fill'
$cardsHost.Margin = New-Object System.Windows.Forms.Padding(0)
$cardsHost.BackColor = $T.Fundo
$layout.Controls.Add($cardsHost, 0, 2)

$script:CardsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:CardsPanel.Dock = 'Fill'
$script:CardsPanel.AutoScroll = $true
$script:CardsPanel.FlowDirection = 'LeftToRight'
$script:CardsPanel.WrapContents = $true
$script:CardsPanel.BackColor = $T.Fundo
$script:CardsPanel.Padding = New-Object System.Windows.Forms.Padding(20, 6, 8, 8)
Enable-TemaBufferDuplo -Controle $script:CardsPanel
$cardsHost.Controls.Add($script:CardsPanel)

$script:EmptyLabel = New-Object System.Windows.Forms.Label
$script:EmptyLabel.Dock = 'Top'
$script:EmptyLabel.Height = 80
$script:EmptyLabel.TextAlign = 'MiddleCenter'
$script:EmptyLabel.Font = $script:Fontes.Texto
$script:EmptyLabel.ForeColor = $T.TextoSuave
$script:EmptyLabel.Visible = $false
$cardsHost.Controls.Add($script:EmptyLabel)
$script:EmptyLabel.BringToFront()

# ---------------- Barra de acoes ----------------
$actionBar = New-Object System.Windows.Forms.TableLayoutPanel
$actionBar.Dock = 'Fill'
$actionBar.Margin = New-Object System.Windows.Forms.Padding(0)
$actionBar.Padding = New-Object System.Windows.Forms.Padding(20, 0, 20, 0)
$actionBar.BackColor = $T.Camada
$actionBar.ColumnCount = 2
$actionBar.RowCount = 1
[void]$actionBar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$actionBar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$actionBar.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$actionBar.Add_Paint({
    param($ctl, $e)
    $p = New-Object System.Drawing.Pen($script:Tema.Contorno)
    $e.Graphics.DrawLine($p, 0, 0, $ctl.Width, 0)
    $p.Dispose()
})
$layout.Controls.Add($actionBar, 0, 3)

$detailText = New-Object System.Windows.Forms.FlowLayoutPanel
$detailText.FlowDirection = 'TopDown'
$detailText.WrapContents = $false
$detailText.AutoSize = $true
$detailText.Anchor = 'Left'
$detailText.Margin = New-Object System.Windows.Forms.Padding(0)
$actionBar.Controls.Add($detailText, 0, 0)

$script:DetailTitle = New-Object System.Windows.Forms.Label
$script:DetailTitle.Font = $script:Fontes.Detalhe
$script:DetailTitle.ForeColor = $T.Texto
$script:DetailTitle.AutoSize = $true
$script:DetailTitle.Margin = New-Object System.Windows.Forms.Padding(0)
$detailText.Controls.Add($script:DetailTitle)

$script:DetailInfo = New-Object System.Windows.Forms.Label
$script:DetailInfo.Font = $script:Fontes.Texto
$script:DetailInfo.ForeColor = $T.TextoSuave
$script:DetailInfo.AutoSize = $true
$script:DetailInfo.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$detailText.Controls.Add($script:DetailInfo)

$buttons = New-Object System.Windows.Forms.FlowLayoutPanel
$buttons.FlowDirection = 'LeftToRight'
$buttons.WrapContents = $false
$buttons.AutoSize = $true
$buttons.Anchor = 'Right'
$buttons.Margin = New-Object System.Windows.Forms.Padding(0)
$actionBar.Controls.Add($buttons, 1, 0)

$script:MoreButton = New-Object System.Windows.Forms.Button
$script:MoreButton.Text = [string][char]0xE712
Set-TemaBotao -Botao $script:MoreButton -Tipo Secundario
$script:MoreButton.Font = $script:Fontes.IconeBotao
$script:MoreButton.Size = New-Object System.Drawing.Size(40, 38)
$script:MoreButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$script:MoreButton.AccessibleName = 'Mais ações'
$buttons.Controls.Add($script:MoreButton)

$script:LogsButton = New-Object System.Windows.Forms.Button
$script:LogsButton.Text = 'Logs da rotina'
Set-TemaBotao -Botao $script:LogsButton -Tipo Secundario
$script:LogsButton.Size = New-Object System.Drawing.Size(128, 38)
$script:LogsButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$buttons.Controls.Add($script:LogsButton)

$script:RunButton = New-Object System.Windows.Forms.Button
$script:RunButton.Text = 'Abrir rotina'
Set-TemaBotao -Botao $script:RunButton -Tipo Principal
$script:RunButton.Size = New-Object System.Drawing.Size(156, 38)
$script:RunButton.Margin = New-Object System.Windows.Forms.Padding(0)
$script:RunButton.Add_EnabledChanged({
    if ($this.Enabled) { $this.BackColor = $script:Tema.VerdeEscuro; $this.FlatAppearance.BorderColor = $script:Tema.VerdeEscuro }
    else { $this.BackColor = $script:Tema.Contorno; $this.FlatAppearance.BorderColor = $script:Tema.Contorno }
})
$buttons.Controls.Add($script:RunButton)

$moreMenu = New-Object System.Windows.Forms.ContextMenuStrip
$moreMenu.Font = $script:Fontes.Texto
$menuFolder = $moreMenu.Items.Add('Abrir pasta do script')
$menuCentralLogs = $moreMenu.Items.Add('Abrir logs da Central')

# ---------------- Status ----------------
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = 'Fill'
$statusStrip.Margin = New-Object System.Windows.Forms.Padding(0)
$statusStrip.SizingGrip = $false
$statusStrip.BackColor = $T.Fundo
$statusStrip.Font = $script:Fontes.Legenda
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Spring = $true
$script:StatusLabel.TextAlign = 'MiddleLeft'
$script:StatusLabel.ForeColor = $T.TextoSuave
$script:EnvLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:EnvLabel.ForeColor = $T.TextoSuave
$script:EnvLabel.Text = 'Verificando ambiente...'
[void]$statusStrip.Items.Add($script:StatusLabel)
[void]$statusStrip.Items.Add($script:EnvLabel)
$layout.Controls.Add($statusStrip, 0, 4)

# ---------------- Eventos ----------------
$script:SearchBox.Add_TextChanged({ Update-ScriptList })
$script:CardsPanel.Add_Resize({ Update-CardWidths })
$script:RunButton.Add_Click({ Invoke-SelectedRoutine })
$script:LogsButton.Add_Click({
    if (-not $script:SelectedScript) { return }
    $pasta = Get-RoutineLogsPath -Item $script:SelectedScript
    if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $pasta)
})
$script:MoreButton.Add_Click({ $moreMenu.Show($script:MoreButton, 0, $script:MoreButton.Height) })
$menuFolder.Add_Click({
    if ($script:SelectedScript) {
        $folder = Split-Path -Parent (Get-ResolvedScriptPath -Item $script:SelectedScript)
        if (Test-Path -LiteralPath $folder) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $folder) }
    }
})
$menuCentralLogs.Add_Click({
    if (-not (Test-Path -LiteralPath $script:LogsPath)) { New-Item -ItemType Directory -Path $script:LogsPath -Force | Out-Null }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:LogsPath)
})

$form.Add_KeyDown({
    param($ctl, $e)
    if ($script:ActivePanel) { return }   # rotina aberta: teclado e da rotina
    if ($e.Control -and $e.KeyCode -eq 'F') {
        $script:SearchBox.Focus(); $script:SearchBox.SelectAll(); $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq 'Escape' -and $script:SearchBox.Text) {
        $script:SearchBox.Clear(); $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq 'Enter') {
        if (-not $script:SelectedScript -and $script:CardControls.Count -eq 1) {
            Select-ScriptCard -Item $script:CardControls[0].Tag.Item -Card $script:CardControls[0]
        }
        if ($script:SelectedScript) { Invoke-SelectedRoutine }
        $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -in @('Down', 'Right') -and -not ($script:SearchBox.Focused -and $e.KeyCode -eq 'Right')) {
        Move-CardSelection -Passo 1; $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -in @('Up', 'Left') -and -not ($script:SearchBox.Focused -and $e.KeyCode -eq 'Left')) {
        Move-CardSelection -Passo -1; $e.SuppressKeyPress = $true
    }
})

$form.Add_Shown({
    Update-ScriptList
    Write-ExecutorLog "ABERTO | Catálogo com $($script:Catalog.Count) rotinas"
    $form.Refresh()
    # Verificacao rapida: so modulos do PowerShell (Python/Chromium sao checados ao abrir a rotina).
    $faltando = @(foreach ($modulo in @('PWPS_DAB', 'ImportExcel')) {
        if (-not (Get-Module -ListAvailable -Name $modulo | Select-Object -First 1)) { $modulo }
    })
    if ($faltando.Count -eq 0) {
        $script:EnvLabel.Text = '●  Ambiente pronto: PWPS_DAB e ImportExcel encontrados'
        $script:EnvLabel.ForeColor = $script:Tema.VerdeEscuro
    }
    else {
        $script:EnvLabel.Text = "●  Faltando: $($faltando -join ', ')"
        $script:EnvLabel.ForeColor = $script:Tema.Critico
    }
})
$form.Add_FormClosing({
    param($ctl, $eventArgs)
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
