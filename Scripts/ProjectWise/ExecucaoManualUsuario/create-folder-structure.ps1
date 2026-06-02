param(
    [string]$BasePath,
    [switch]$WhatIf
)

function Select-BaseFolder {
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione a pasta onde a estrutura sera criada"
    $dialog.ShowNewFolderButton = $true
    $dialog.SelectedPath = [Environment]::GetFolderPath("Desktop")

    $result = $dialog.ShowDialog()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
        Write-Host "Operacao cancelada. Nenhuma pasta foi selecionada."
        exit 0
    }

    return $dialog.SelectedPath
}

$folders = @(
    "COLETA_DIURNA",
    "COLETA_DIURNA\Dispositivos",
    "COLETA_DIURNA\Vias",
    "COLETA_NOTURNA",
    "COLETA_NOTURNA\Dispositivos",
    "COLETA_NOTURNA\Vias"
)

$files = @()

if ([string]::IsNullOrWhiteSpace($BasePath)) {
    $BasePath = Select-BaseFolder
}

$resolvedBasePath = [System.IO.Path]::GetFullPath($BasePath)

Write-Host "Base: $resolvedBasePath"

if (-not (Test-Path -LiteralPath $resolvedBasePath)) {
    if ($WhatIf) {
        Write-Host "[WhatIf] Criaria pasta base: $resolvedBasePath"
    }
    else {
        New-Item -ItemType Directory -Path $resolvedBasePath -Force | Out-Null
        Write-Host "Criada pasta base: $resolvedBasePath"
    }
}

foreach ($folder in $folders) {
    $target = Join-Path $resolvedBasePath $folder

    if ($WhatIf) {
        Write-Host "[WhatIf] Criaria pasta: $target"
        continue
    }

    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Write-Host "Pasta pronta: $target"
}

foreach ($file in $files) {
    $target = Join-Path $resolvedBasePath $file
    $parent = Split-Path -Parent $target

    if ($WhatIf) {
        Write-Host "[WhatIf] Criaria arquivo: $target"
        continue
    }

    New-Item -ItemType Directory -Path $parent -Force | Out-Null

    if (-not (Test-Path -LiteralPath $target)) {
        New-Item -ItemType File -Path $target | Out-Null
        Write-Host "Arquivo criado: $target"
    }
    else {
        Write-Host "Arquivo ja existe: $target"
    }
}

Write-Host "Estrutura finalizada."
