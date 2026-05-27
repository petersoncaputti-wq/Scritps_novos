<#
.SYNOPSIS
    Ajuste de Atributos - preenche o atributo "Nº Documento" com o valor do campo Name.

.DESCRIPTION
    Script para ProjectWise Explorer CONNECT utilizando o modulo PWPS_DAB.
    Localiza a subpasta "Importacao" de um projeto informado, lista os documentos dessa pasta
    e atualiza o atributo "Nº Documento" com o Name da aba General do documento.

    Compatibilidade: Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
    # Quando informado, sobrescreve a variavel $ProjectPath configurada abaixo.
    [string]$ProjectPathParam,

    # Executa sem alterar documentos no ProjectWise.
    [switch]$DryRunParam,

    # Forca a execucao real pela linha de comando, mesmo se $DryRun estiver $true no bloco de configuracoes.
    [switch]$ExecuteParam
)

#requires -Version 5.1

# =================================================================================================
# CONFIGURACOES
# =================================================================================================

# Caminho do projeto no ProjectWise. Ajuste antes de executar.
# Exemplo: "Datasource\Projetos\Projeto-001"
$ProjectPath = "INFORME_AQUI_O_CAMINHO_DO_PROJETO"

# Nome da subpasta que sera localizada dentro do projeto.
$ImportFolderName = "Importacao"

# Nome exato do atributo na aba Attributes.
$NumeroDocumentoAttributeName = "Nº Documento"

# Modo simulacao: $true nao altera documentos; $false executa a atualizacao.
$DryRun = $true

# =================================================================================================
# INICIALIZACAO
# =================================================================================================

if (-not [string]::IsNullOrWhiteSpace($ProjectPathParam)) {
    $ProjectPath = $ProjectPathParam
}

if ($DryRunParam.IsPresent) {
    $DryRun = $true
}

if ($ExecuteParam.IsPresent) {
    $DryRun = $false
}

$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "OK", "WARN", "ERROR", "DRYRUN")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = "[$timestamp][$Level]"

    switch ($Level) {
        "OK"     { Write-Host "$prefix $Message" -ForegroundColor Green }
        "WARN"   { Write-Host "$prefix $Message" -ForegroundColor Yellow }
        "ERROR"  { Write-Host "$prefix $Message" -ForegroundColor Red }
        "DRYRUN" { Write-Host "$prefix $Message" -ForegroundColor Cyan }
        default  { Write-Host "$prefix $Message" -ForegroundColor Gray }
    }
}

function Get-DocumentPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,

        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        $property = $Document.PSObject.Properties[$propertyName]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Test-DocumentIsCheckedOut {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document
    )

    # Versoes diferentes do PWPS_DAB podem expor o estado de bloqueio com nomes diferentes.
    $booleanLockProperties = @(
        "IsCheckedOut",
        "CheckedOut",
        "IsLocked",
        "Locked",
        "DocumentLocked",
        "IsDocumentLocked"
    )

    foreach ($propertyName in $booleanLockProperties) {
        $value = Get-DocumentPropertyValue -Document $Document -PropertyNames @($propertyName)
        if ($null -ne $value) {
            try {
                if ([System.Convert]::ToBoolean($value)) {
                    return $true
                }
            }
            catch {
                # Se a propriedade existir mas nao for booleana, a avaliacao textual abaixo cobre o caso.
            }
        }
    }

    $checkedOutBy = Get-DocumentPropertyValue -Document $Document -PropertyNames @(
        "CheckedOutBy",
        "OutTo",
        "OutToUser",
        "UserNameOut",
        "DocumentOutTo"
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$checkedOutBy)) {
        return $true
    }

    return $false
}

function Test-AttributeExists {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,

        [Parameter(Mandatory = $true)]
        [string]$AttributeName
    )

    # Get-PWDocumentsBySearch -GetAttributes normalmente adiciona os atributos como propriedades do objeto.
    if ($null -ne $Document.PSObject.Properties[$AttributeName]) {
        return $true
    }

    # Algumas versoes retornam uma colecao/hashtable de atributos.
    $attributeContainers = @("Attributes", "DocumentAttributes", "EnvironmentAttributes")
    foreach ($containerName in $attributeContainers) {
        $container = Get-DocumentPropertyValue -Document $Document -PropertyNames @($containerName)
        if ($null -eq $container) {
            continue
        }

        if ($container -is [System.Collections.IDictionary] -and $container.Contains($AttributeName)) {
            return $true
        }

        $attributeProperty = $container.PSObject.Properties[$AttributeName]
        if ($null -ne $attributeProperty) {
            return $true
        }
    }

    return $false
}

function Get-DocumentDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document
    )

    $name = Get-DocumentPropertyValue -Document $Document -PropertyNames @("Name", "DocumentName")
    $fileName = Get-DocumentPropertyValue -Document $Document -PropertyNames @("FileName", "OriginalFileName")
    $documentId = Get-DocumentPropertyValue -Document $Document -PropertyNames @("DocumentID", "DocID", "Id")

    if (-not [string]::IsNullOrWhiteSpace([string]$name)) {
        return [string]$name
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$fileName)) {
        return [string]$fileName
    }

    if ($null -ne $documentId) {
        return "DocumentID=$documentId"
    }

    return "<documento sem identificador>"
}

function Join-PWPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    return ($ParentPath.TrimEnd("\", "/") + "\" + $ChildPath.TrimStart("\", "/"))
}

try {
    Write-Log "Iniciando preenchimento do atributo '$NumeroDocumentoAttributeName'." "INFO"

    if ([string]::IsNullOrWhiteSpace($ProjectPath) -or $ProjectPath -eq "INFORME_AQUI_O_CAMINHO_DO_PROJETO") {
        throw "Configure a variavel `$ProjectPath no inicio do script ou informe -ProjectPathParam."
    }

    Write-Log "Importando modulo PWPS_DAB." "INFO"
    Import-Module PWPS_DAB -ErrorAction Stop

    $importFolderPath = Join-PWPath -ParentPath $ProjectPath -ChildPath $ImportFolderName
    Write-Log "Projeto informado: $ProjectPath" "INFO"
    Write-Log "Validando subpasta '$ImportFolderName': $importFolderPath" "INFO"

    # Confirma que a subpasta Importacao existe antes de buscar documentos.
    $importFolder = Get-PWFolders -FolderPath $importFolderPath -JustOne -ErrorAction Stop
    if ($null -eq $importFolder) {
        throw "A subpasta '$ImportFolderName' nao foi encontrada no projeto informado."
    }

    Write-Log "Subpasta '$ImportFolderName' localizada com sucesso." "OK"

    if ($DryRun) {
        Write-Log "Modo simulacao ativo. Nenhum documento sera alterado." "DRYRUN"
    }

    # Busca somente documentos existentes diretamente na pasta Importacao, trazendo os atributos.
    Write-Log "Buscando documentos na pasta Importacao." "INFO"
    $documents = @(Get-PWDocumentsBySearch -FolderPath $importFolderPath -JustThisFolder -GetAttributes -PopulatePath -ErrorAction Stop)

    $totalFound = $documents.Count
    $successCount = 0
    $errorCount = 0
    $ignoredCount = 0

    Write-Log "Documentos encontrados: $totalFound" "INFO"

    foreach ($document in $documents) {
        $displayName = Get-DocumentDisplayName -Document $document

        try {
            Write-Log "Processando documento: $displayName" "INFO"

            # Le o campo Name da aba General.
            $documentName = Get-DocumentPropertyValue -Document $document -PropertyNames @("Name", "DocumentName")
            if ([string]::IsNullOrWhiteSpace([string]$documentName)) {
                $ignoredCount++
                Write-Log "Ignorado: campo Name vazio." "WARN"
                continue
            }

            # Evita alterar documentos bloqueados/check-out.
            if (Test-DocumentIsCheckedOut -Document $document) {
                $ignoredCount++
                Write-Log "Ignorado: documento bloqueado ou em check-out." "WARN"
                continue
            }

            # Confirma que o atributo existe no ambiente retornado para o documento.
            if (-not (Test-AttributeExists -Document $document -AttributeName $NumeroDocumentoAttributeName)) {
                $ignoredCount++
                Write-Log "Ignorado: atributo '$NumeroDocumentoAttributeName' nao encontrado." "WARN"
                continue
            }

            $attributesToUpdate = @{
                $NumeroDocumentoAttributeName = [string]$documentName
            }

            if ($DryRun) {
                $successCount++
                Write-Log "Simulacao: atributo '$NumeroDocumentoAttributeName' seria preenchido com '$documentName'." "DRYRUN"
                continue
            }

            # Atualiza e salva os atributos do documento no ProjectWise.
            $updateResult = Update-PWDocumentAttributes -InputDocuments @($document) -Attributes $attributesToUpdate -ReturnBoolean -ErrorAction Stop

            if ($updateResult -eq $false) {
                throw "Update-PWDocumentAttributes retornou False."
            }

            $successCount++
            Write-Log "Atualizado com sucesso: '$NumeroDocumentoAttributeName' = '$documentName'." "OK"
        }
        catch {
            $errorCount++
            Write-Log "Falha ao processar '$displayName'. Erro: $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Host ""
    Write-Log "Resumo final" "INFO"
    Write-Log "Documentos encontrados: $totalFound" "INFO"
    Write-Log "Atualizados com sucesso: $successCount" "OK"
    Write-Log "Ignorados: $ignoredCount" "WARN"
    Write-Log "Com erro: $errorCount" "ERROR"

    if ($DryRun) {
        Write-Log "Execucao finalizada em modo simulacao. Para alterar documentos, configure `$DryRun = `$false." "DRYRUN"
        Write-Log "Opcionalmente, execute com -ExecuteParam para forcar a atualizacao real." "DRYRUN"
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" "ERROR"
    exit 1
}
