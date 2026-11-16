param(
    [string]$Url = 'https://cloud.engemap.com.br/s/SmfDfn34BFy25Mo',
    [string]$Destino,
    [switch]$Listar,
    [switch]$SemSenha
)
$ErrorActionPreference = 'Stop'
$python = Get-Command python -ErrorAction SilentlyContinue
if (!$python) { throw 'Instale Python 3.10 ou superior e habilite-o no PATH.' }
if (!$Listar -and !$Destino) {
    # Executar o seletor em STA permite usa-lo tambem no PowerShell 7 (MTA).
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $picker = [PowerShell]::Create()
    try {
        $runspace.Open()
        $picker.Runspace = $runspace
        [void]$picker.AddScript({
            Add-Type -AssemblyName System.Windows.Forms
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            try {
                $dialog.Description = 'Selecione a pasta de destino dos arquivos Engemap'
                $dialog.ShowNewFolderButton = $true
                if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $dialog.SelectedPath
                }
            }
            finally { $dialog.Dispose() }
        })
        $selection = @($picker.Invoke())
        if ($picker.HadErrors) { throw "Nao foi possivel abrir o seletor de pasta: $($picker.Streams.Error[0])" }
        if (!$selection.Count -or [string]::IsNullOrWhiteSpace([string]$selection[0])) {
            Write-Host 'Selecao cancelada. Nenhum download iniciado.'
            return
        }
        $Destino = [string]$selection[0]
    }
    finally { $picker.Dispose(); $runspace.Dispose() }
    Write-Host "Destino selecionado: $Destino"
}
$arguments = @((Join-Path $PSScriptRoot 'downloader.py'), '--url', $Url)
if ($Destino) { $arguments += @('--destino', $Destino) }
if ($Listar) { $arguments += '--listar' }
if ($SemSenha) { $arguments += '--sem-senha' }
& $python.Source @arguments
if ($LASTEXITCODE -ne 0) { throw "Downloader finalizou com codigo $LASTEXITCODE." }
