[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$settingsPath = Join-Path $projectRoot "config\settings.json"

. (Join-Path $PSScriptRoot "Connect-GraphSharePoint.ps1")

try {
    Write-Host "Carregando configuracao: $settingsPath"

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        throw "Arquivo de configuracao nao encontrado: $settingsPath"
    }

    $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json

    if ([string]::IsNullOrWhiteSpace($settings.SharePoint.TenantUrl)) {
        throw "Configure SharePoint.TenantUrl no arquivo config/settings.json."
    }

    if ([string]::IsNullOrWhiteSpace($settings.SharePoint.TenantId)) {
        throw "Configure SharePoint.TenantId no arquivo config/settings.json."
    }

    if (-not $settings.SharePoint.Scopes -or $settings.SharePoint.Scopes.Count -eq 0) {
        throw "Configure SharePoint.Scopes no arquivo config/settings.json."
    }

    Write-Host "Abrindo login interativo do Microsoft Graph."
    Write-Host "Tenant: $($settings.SharePoint.TenantUrl)"
    Write-Host "TenantId: $($settings.SharePoint.TenantId)"
    Write-Host "Scopes: $($settings.SharePoint.Scopes -join ', ')"

    $connectionResult = Connect-GraphSharePoint `
        -TenantId $settings.SharePoint.TenantId `
        -Scopes $settings.SharePoint.Scopes
    $connectionResult | Format-List

    if (-not $connectionResult.Success) {
        exit 1
    }
}
catch {
    [pscustomobject]@{
        Success = $false
        Step    = "Main"
        Message = "Falha ao executar a fase inicial."
        Error   = [pscustomobject]@{
            Message = $_.Exception.Message
            Type    = $_.Exception.GetType().FullName
        }
    } | Format-List

    exit 1
}
