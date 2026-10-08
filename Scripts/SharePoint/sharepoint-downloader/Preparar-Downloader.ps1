param([switch]$PrepareOnly, [Parameter(ValueFromRemainingArguments=$true)][string[]]$DownloaderArgs)
$ErrorActionPreference = 'Stop'

function Find-CompatiblePython {
    $candidates = @()
    $pyCommand = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyCommand) { $candidates += @{Source=$pyCommand.Source; Args=@('-3')} }
    # Also probe the installer default directly: PATH in this process may be stale.
    $defaultPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'
    if (Test-Path -LiteralPath $defaultPython) { $candidates += @{Source=$defaultPython; Args=@()} }
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand -and $pythonCommand.Source -notlike '*\Microsoft\WindowsApps\*') {
        $candidates += @{Source=$pythonCommand.Source; Args=@()}
    }
    foreach ($candidate in $candidates) {
        try {
            $candidateArgs = $candidate.Args
            & $candidate.Source @candidateArgs -c 'import sys, tkinter, venv, ensurepip; assert sys.version_info >= (3,11)' 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch { }
    }
    return $null
}

function Get-OrInstallPython {
    $compatible = Find-CompatiblePython
    if ($compatible) { return $compatible }
    $wingetCommand = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $wingetCommand) {
        throw 'Python compativel e winget nao encontrados. Instale/atualize o Instalador de Aplicativos do Windows ou instale Python 3.13 com Tcl/Tk.'
    }
    Write-Host 'Python compativel nao encontrado. Instalando Python 3.13 para o usuario atual via winget...'
    & $wingetCommand.Source install --id Python.Python.3.13 --exact --source winget --scope user --silent --accept-package-agreements --accept-source-agreements | Out-Host
    $installExitCode = $LASTEXITCODE
    $compatible = Find-CompatiblePython
    if (-not $compatible) {
        throw "Python nao ficou disponivel (winget: $installExitCode). Verifique internet, proxy e politicas da TI. Feche e execute o BAT novamente se a instalacao terminou com sucesso."
    }
    return $compatible
}

try {
    Set-Location -LiteralPath $PSScriptRoot
    $envPath = Join-Path $PSScriptRoot '.venv'
    $pythonPath = Join-Path $envPath 'Scripts\python.exe'
    $markerPath = Join-Path $envPath '.downloader-setup.json'
    $requirementsPath = Join-Path $PSScriptRoot 'requirements.txt'
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'main.py')) -or -not (Test-Path -LiteralPath $requirementsPath)) {
        throw 'Copie a pasta completa do downloader, incluindo main.py e requirements.txt.'
    }
    $marker = $null
    if (Test-Path -LiteralPath $markerPath) {
        try { $marker = Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json } catch { $marker = $null }
    }
    $validEnvironment = $false
    if (Test-Path -LiteralPath $pythonPath) {
        & $pythonPath -c 'import sys, tkinter; assert sys.version_info >= (3,11)' 2>$null
        $validEnvironment = ($LASTEXITCODE -eq 0)
    }
    if ($marker -and $marker.folder -ne $PSScriptRoot) { $validEnvironment = $false }
    if (-not $validEnvironment) {
        $launcher = Get-OrInstallPython
        $pythonArgs = $launcher.Args
        if (Test-Path -LiteralPath $envPath) {
            # Preserve an unusable/copied environment; never delete it automatically.
            $resolvedEnv = (Resolve-Path -LiteralPath $envPath).Path
            if ([IO.Path]::GetDirectoryName($resolvedEnv) -ne $PSScriptRoot) { throw 'Ambiente virtual fora da pasta do downloader.' }
            $backupName = '.venv_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + [guid]::NewGuid().ToString('N').Substring(0,8)
            Rename-Item -LiteralPath $resolvedEnv -NewName $backupName
        }
        Write-Host 'Criando ambiente Python neste notebook...'
        & $launcher.Source @pythonArgs -m venv $envPath
        if ($LASTEXITCODE -ne 0) { throw 'Nao foi possivel criar o ambiente Python.' }
        $marker = $null
    }
    $requirementsHash = (Get-FileHash -LiteralPath $requirementsPath -Algorithm SHA256).Hash
    $needsInstall = (-not $marker -or $marker.requirements -ne $requirementsHash)
    if (-not $needsInstall) {
        & $pythonPath -c 'import playwright.sync_api, requests, msal, msal_extensions, tkinter' 2>$null
        $needsInstall = ($LASTEXITCODE -ne 0)
    }
    if ($needsInstall) {
        Write-Host 'Instalando/verificando dependencias. Esta etapa requer internet...'
        & $pythonPath -m pip install -r $requirementsPath
        if ($LASTEXITCODE -ne 0) { throw 'Falha na instalacao das dependencias. Verifique internet/proxy corporativo e execute novamente.' }
        & $pythonPath -c 'import playwright.sync_api, requests, msal, msal_extensions, tkinter'
        if ($LASTEXITCODE -ne 0) { throw 'Dependencias incompletas.' }
        @{folder=$PSScriptRoot; requirements=$requirementsHash} | ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding UTF8
    }
    Write-Host 'Ambiente pronto. O Microsoft Edge deve estar instalado para o login.'
    if ($PrepareOnly) { exit 0 }
    & $pythonPath (Join-Path $PSScriptRoot 'main.py') --browser-all @DownloaderArgs
    exit $LASTEXITCODE
} catch {
    Write-Host ('Erro ao preparar o downloader: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
