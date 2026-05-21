function Connect-GraphSharePoint {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph for SharePoint file access.

    .DESCRIPTION
        Opens an interactive Microsoft Graph login, validates the current
        Graph context, and returns a standardized result object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Scopes
    )

    try {
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            throw "Este script deve ser executado no PowerShell 7 ou superior. Versao atual: $($PSVersionTable.PSVersion)."
        }

        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            throw "Modulo Microsoft.Graph.Authentication nao encontrado. Instale no PowerShell 7 com: Install-Module Microsoft.Graph -Scope CurrentUser"
        }

        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome -ErrorAction Stop

        # Basic validation: if a context exists, Graph authentication succeeded.
        $context = Get-MgContext -ErrorAction Stop

        if (-not $context) {
            throw "Conexao Graph nao retornou contexto autenticado."
        }

        return [pscustomobject]@{
            Success   = $true
            Step      = "Connect-GraphSharePoint"
            TenantId  = $TenantId
            Account   = $context.Account
            Scopes    = $context.Scopes
            Message   = "Conexao com Microsoft Graph realizada com sucesso."
            Error     = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Success   = $false
            Step      = "Connect-GraphSharePoint"
            TenantId  = $TenantId
            Account   = $null
            Scopes    = $Scopes
            Message   = "Falha ao conectar ao Microsoft Graph."
            Error     = [pscustomobject]@{
                Message = $_.Exception.Message
                Type    = $_.Exception.GetType().FullName
            }
        }
    }
}
