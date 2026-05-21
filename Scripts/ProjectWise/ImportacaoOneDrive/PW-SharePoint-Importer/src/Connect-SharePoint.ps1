function Connect-SharePoint {
    <#
    .SYNOPSIS
        Connects to SharePoint using PnP.PowerShell.

    .DESCRIPTION
        Opens an interactive login using Connect-PnPOnline, validates the
        connection with Get-PnPWeb, and returns a standardized result object.
        The URL can be the tenant root or a specific SharePoint site.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SharePointUrl,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId
    )

    try {
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            throw "Este script deve ser executado no PowerShell 7 ou superior. Versao atual: $($PSVersionTable.PSVersion)."
        }

        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            throw "Modulo PnP.PowerShell nao encontrado. Instale no PowerShell 7 com: Install-Module PnP.PowerShell -Scope CurrentUser"
        }

        Connect-PnPOnline -Url $SharePointUrl -Interactive -ClientId $ClientId -ErrorAction Stop

        # Basic validation: if this call succeeds, the current PnP connection is usable.
        $web = Get-PnPWeb -ErrorAction Stop

        return [pscustomobject]@{
            Success   = $true
            Step      = "Connect-SharePoint"
            Url       = $SharePointUrl
            ClientId  = $ClientId
            SiteTitle = $web.Title
            Message   = "Conexao com SharePoint realizada com sucesso."
            Error     = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Success   = $false
            Step      = "Connect-SharePoint"
            Url       = $SharePointUrl
            ClientId  = $ClientId
            SiteTitle = $null
            Message   = "Falha ao conectar ao SharePoint."
            Error     = [pscustomobject]@{
                Message = $_.Exception.Message
                Type    = $_.Exception.GetType().FullName
            }
        }
    }
}
