<#
.SYNOPSIS
    Ajuste de Atributos - preenche o atributo "Nº Documento" com parte do valor do campo Name.

.DESCRIPTION
    Script para ProjectWise Explorer CONNECT utilizando o modulo PWPS_DAB.
    Solicita a concessao, lista os projetos, permite navegar ate a pasta dos arquivos,
    lista os documentos dessa pasta
    e atualiza o atributo "Nº Documento" com o codigo extraido do Name da aba General do documento.
    Exemplo: ERM-493RJ-112+815-EXE-DE-L3-001-R01a.dwg -> ERM-493RJ-112+815-EXE-DE-L3-001.

    Compatibilidade: Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
    # Opcional. Quando informado, pula a selecao por concessao/projeto e usa este caminho diretamente.
    [string]$ProjectPathParam,

    # Executa sem alterar documentos no ProjectWise.
    [switch]$DryRunParam,

    # Forca a execucao real pela linha de comando, mesmo se $DryRun estiver $true no bloco de configuracoes.
    [switch]$ExecuteParam,

    # Limita a quantidade de documentos processados. Use 0 para processar todos.
    [int]$LimiteDocumentosParam = -1,

    # Caminho opcional do arquivo de log. Se vazio, cria um log na pasta Logs ao lado do script.
    [string]$LogPathParam
)

#requires -Version 5.1

# =================================================================================================
# CONFIGURACOES
# =================================================================================================

# Caminho do projeto no ProjectWise.
# Deixe vazio para selecionar a concessao e o projeto pelo console.
# Se preenchido, o script usa este caminho diretamente.
$ProjectPath = ""

# Possiveis nomes da pasta raiz de engenharia e da pasta de projetos.
$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")

# Profundidade adicional para procurar projetos dentro da pasta "Projetos".
# 0 = somente filhos diretos. Aumente para 1 ou 2 se houver agrupamentos intermediarios.
$ProfundidadeProjetos = 0

# Nome do datasource ProjectWise.
# Deixe vazio para usar o login Bentley IMS padrao, como nos outros scripts.
# Se quiser forcar um datasource especifico, preencha esta variavel.
$DatasourceName = ""

# Tipo de login. BentleyIMS e o padrao para ambientes CONNECT.
$UsarBentleyIMS = $true

# Nome da subpasta usada apenas quando $ProjectPath ou -ProjectPathParam forem informados.
# No modo interativo, a pasta dos arquivos e escolhida navegando pelas subpastas do projeto.
$ImportFolderName = "Importacao"

# Nome exato do atributo na aba Attributes.
# Montado com [char] para evitar problema de encoding no Windows PowerShell 5.1.
$NumeroDocumentoAttributeName = "N$([char]0x00BA) Documento"

# Nome interno da coluna do atributo no ambiente ProjectWise.
# Deixe vazio para o script tentar descobrir automaticamente em um documento de amostra.
# Exemplos comuns: NumeroDocumento, N_Documento, NumDocumento.
$NumeroDocumentoAttributeColumnName = "NumeroPoderConcedente"

# Nome do ambiente ProjectWise da pasta/documentos.
# Deixe vazio para tentar descobrir automaticamente ou solicitar no console.
# Pela tela do ProjectWise, costuma aparecer como "Environment Name".
$ProjectWiseEnvironmentName = ""

# Modo simulacao: $true nao altera documentos; $false executa a atualizacao.
$DryRun = $true

# Quantidade maxima de documentos para processar.
# 0 = processa todos. Use 5, 10 etc. para teste inicial em producao.
$LimiteDocumentos = 0

# Busca rapida: nao carrega atributos de todos os documentos na busca inicial.
$ModoBuscaRapida = $true

# Valida o atributo somente uma vez em um documento de amostra antes do lote.
# Mantido desativado porque atributos vazios podem nao ser retornados pelo PWPS_DAB.
# A validacao principal passa a ser o campo Name preenchido.
$ValidarAtributoUmaVez = $false

# Caminho do arquivo de log. Deixe vazio para gerar automaticamente em .\Logs.
$LogFilePath = ""

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

if ($LimiteDocumentosParam -ge 0) {
    $LimiteDocumentos = $LimiteDocumentosParam
}

if (-not [string]::IsNullOrWhiteSpace($LogPathParam)) {
    $LogFilePath = $LogPathParam
}

function Restart-ScriptInMtaIfNeeded {
    param(
        [hashtable]$BoundParameters
    )

    $apartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartmentState -eq [System.Threading.ApartmentState]::MTA) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw "O PWPS_DAB precisa de PowerShell em modo MTA. Execute o script com: powershell.exe -Mta -File `"<caminho do script>`""
    }

    $powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $powershellExe)) {
        throw "Nao foi possivel localizar o Windows PowerShell 5.1 em '$powershellExe'."
    }

    Write-Host "[INFO] Sessao atual esta em $apartmentState. Reiniciando o script em PowerShell 5.1 MTA para carregar o PWPS_DAB..." -ForegroundColor Yellow

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Mta",
        "-File",
        $PSCommandPath
    )

    foreach ($key in $BoundParameters.Keys) {
        $value = $BoundParameters[$key]
        if ($value -is [switch] -or $value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) {
                $arguments += "-$key"
            }
            continue
        }

        if ($null -ne $value) {
            $arguments += "-$key"
            $arguments += [string]$value
        }
    }

    & $powershellExe @arguments
    exit $LASTEXITCODE
}

Restart-ScriptInMtaIfNeeded -BoundParameters $PSBoundParameters

$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"

if ([string]::IsNullOrWhiteSpace($LogFilePath)) {
    $logDirectory = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        Join-Path $PSScriptRoot "Logs"
    }
    else {
        Join-Path (Get-Location).Path "Logs"
    }

    $LogFilePath = Join-Path $logDirectory ("Ajuste_de_Atributos_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "OK", "WARN", "ERROR", "DRYRUN")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = "[$timestamp][$Level]"
    $line = "$prefix $Message"

    switch ($Level) {
        "OK"     { Write-Host $line -ForegroundColor Green }
        "WARN"   { Write-Host $line -ForegroundColor Yellow }
        "ERROR"  { Write-Host $line -ForegroundColor Red }
        "DRYRUN" { Write-Host $line -ForegroundColor Cyan }
        default  { Write-Host $line -ForegroundColor Gray }
    }

    try {
        $logFolder = Split-Path -Parent $script:LogFilePath
        if (-not [string]::IsNullOrWhiteSpace($logFolder) -and -not (Test-Path -LiteralPath $logFolder)) {
            New-Item -Path $logFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "[WARN] Nao foi possivel gravar no arquivo de log '$script:LogFilePath'. Erro: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Write-ExceptionLog {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [string]$Context = "Erro"
    )

    Write-Log "$Context`: $($ErrorRecord.Exception.Message)" "ERROR"

    if ($null -ne $ErrorRecord.InvocationInfo) {
        Write-Log "Linha: $($ErrorRecord.InvocationInfo.ScriptLineNumber) | Comando: $($ErrorRecord.InvocationInfo.Line.Trim())" "ERROR"
    }

    if (-not [string]::IsNullOrWhiteSpace($ErrorRecord.ScriptStackTrace)) {
        Write-Log "StackTrace PowerShell: $($ErrorRecord.ScriptStackTrace)" "ERROR"
    }
}

function Resolve-ExecutionMode {
    param(
        [bool]$DryRunWasRequested,
        [bool]$ExecuteWasRequested,
        [bool]$DefaultDryRun
    )

    if ($DryRunWasRequested -or $ExecuteWasRequested) {
        return [PSCustomObject]@{
            DryRun = $DefaultDryRun
            SelectedByPrompt = $false
        }
    }

    Write-Host ""
    Write-Host "Modo de execucao:" -ForegroundColor Cyan
    Write-Host "  S - Simulacao: nao altera documentos"
    Write-Host "  E - Executar: altera documentos no ProjectWise"
    Write-Host "  C - Cancelar"

    while ($true) {
        $selection = (Read-Host "Selecione o modo de execucao [S/E/C]").Trim()

        if ([string]::IsNullOrWhiteSpace($selection)) {
            $selection = "S"
        }

        switch -Regex ($selection) {
            "^(S|SIM|SIMULACAO)$" {
                return [PSCustomObject]@{
                    DryRun = $true
                    SelectedByPrompt = $true
                }
            }
            "^(E|EXECUTAR|REAL|N|NAO)$" {
                return [PSCustomObject]@{
                    DryRun = $false
                    SelectedByPrompt = $true
                }
            }
            "^(C|CANCELAR)$" {
                throw "Execucao cancelada pelo usuario antes de iniciar o processamento."
            }
            default {
                Write-Log "Opcao invalida. Digite S para simulacao, E para executar ou C para cancelar." "WARN"
            }
        }
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

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Get-EnvironmentNameFromObject {
    param(
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return ""
    }

    $environmentName = Get-ObjectPropertyValue -InputObject $InputObject -PropertyNames @(
        "EnvironmentName",
        "Environment",
        "EnvName",
        "EnvironmentDescription",
        "ProjectEnvironment",
        "FolderEnvironment"
    )

    if ([string]::IsNullOrWhiteSpace([string]$environmentName)) {
        return ""
    }

    return ([string]$environmentName).Trim()
}

function Connect-ProjectWiseDatasource {
    try {
        $currentDatasource = Get-PWCurrentDatasource -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$currentDatasource)) {
            Write-Log "Sessao ProjectWise ativa: $currentDatasource" "OK"
            return
        }
    }
    catch {
        Write-Log "Nenhuma sessao ProjectWise ativa encontrada. Sera realizado login." "WARN"
    }

    if ($UsarBentleyIMS) {
        if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
            Write-Log "Efetuando login via Bentley IMS." "INFO"
            New-PWLogin -BentleyIMS -ErrorAction Stop | Out-Null
        }
        else {
            Write-Log "Efetuando login via Bentley IMS no datasource '$DatasourceName'." "INFO"
            New-PWLogin -DatasourceName $DatasourceName -BentleyIMS -ErrorAction Stop | Out-Null
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
            Write-Log "Efetuando login pelo seletor padrao do ProjectWise." "INFO"
            New-PWLogin -ErrorAction Stop | Out-Null
        }
        else {
            Write-Log "Efetuando login no datasource '$DatasourceName'." "INFO"
            New-PWLogin -DatasourceName $DatasourceName -UseGui -ErrorAction Stop | Out-Null
        }
    }

    $currentDatasource = Get-PWCurrentDatasource -ErrorAction Stop
    Write-Log "Login realizado com sucesso: $currentDatasource" "OK"
}

function Get-FolderPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder,

        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        $property = $Folder.PSObject.Properties[$propertyName]
        if ($null -ne $property -and $null -ne $property.Value) {
            $value = ([string]$property.Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ""
}

function Get-FolderId {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @("ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID")
}

function Get-FolderPathValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @("FullPath", "Fullpath", "FolderPath", "Path", "ProjectPath")
}

function Get-FolderNameValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @("Name", "FolderName", "ProjectName", "ObjectName")
}

function Get-FolderDescriptionValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @("Description", "Descricao", "ProjectDescription", "FolderDescription")
}

function Get-FolderLabel {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    $description = Get-FolderDescriptionValue -Folder $Folder
    if (-not [string]::IsNullOrWhiteSpace($description)) {
        return $description
    }

    return Get-FolderNameValue -Folder $Folder
}

function Sort-FoldersByLabel {
    param(
        [object[]]$Folders
    )

    return @($Folders | Sort-Object -Property @{ Expression = { (Get-FolderLabel -Folder $_).ToLowerInvariant() } })
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)

    if ($params -contains "Root") {
        return @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $null -ne $_ })
    }

    return @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $null -ne $_ })
}

function Get-PWChildFoldersSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderId,

        [string]$Context = ""
    )

    try {
        if (-not [string]::IsNullOrWhiteSpace($Context)) {
            Write-Log "Listando subpastas de: $Context" "INFO"
        }

        return @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $null -ne $_ })
    }
    catch {
        Write-Log "Falha ao listar subpastas de '$Context'. Erro: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Find-FolderByPossibleNames {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Folders,

        [Parameter(Mandatory = $true)]
        [string[]]$PossibleNames,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    foreach ($expectedName in $PossibleNames) {
        $found = @(
            $Folders | Where-Object {
                (Get-FolderNameValue -Folder $_) -ieq $expectedName -or
                (Get-FolderLabel -Folder $_) -ieq $expectedName
            }
        )

        if ($found.Count -gt 0) {
            return $found[0]
        }
    }

    $available = @($Folders | ForEach-Object { Get-FolderLabel -Folder $_ }) -join ", "
    throw "Nao encontrei $Description. Pastas disponiveis: $available"
}

function Read-NumberSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [int]$Minimum,

        [Parameter(Mandatory = $true)]
        [int]$Maximum
    )

    while ($true) {
        $inputValue = (Read-Host $Message).Trim()

        if ($inputValue -match '^\d+$') {
            $number = [int]$inputValue
            if ($number -ge $Minimum -and $number -le $Maximum) {
                return $number
            }
        }

        Write-Log "Informe um numero entre $Minimum e $Maximum." "WARN"
    }
}

function Select-FolderFromList {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Folders,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan

    for ($i = 0; $i -lt $Folders.Count; $i++) {
        $label = Get-FolderLabel -Folder $Folders[$i]
        "{0:000}. {1}" -f ($i + 1), $label | Write-Host
    }

    $selection = Read-NumberSelection -Message $Prompt -Minimum 1 -Maximum $Folders.Count
    return $Folders[$selection - 1]
}

function Get-ProjectsFromProjectsFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProjectsFolder,

        [int]$MaximumDepth = 0
    )

    $projectsFolderId = Get-FolderId -Folder $ProjectsFolder
    if ([string]::IsNullOrWhiteSpace($projectsFolderId)) {
        throw "Nao foi possivel identificar o ID da pasta de projetos."
    }

    $result = @()
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue([PSCustomObject]@{
        FolderId = $projectsFolderId
        Level = 0
        Label = Get-FolderLabel -Folder $ProjectsFolder
    })

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $children = Get-PWChildFoldersSafe -FolderId $current.FolderId -Context $current.Label

        foreach ($child in $children) {
            $childId = Get-FolderId -Folder $child
            if ([string]::IsNullOrWhiteSpace($childId)) {
                continue
            }

            $result += $child

            if ($current.Level -lt $MaximumDepth) {
                $queue.Enqueue([PSCustomObject]@{
                    FolderId = $childId
                    Level = ($current.Level + 1)
                    Label = Get-FolderLabel -Folder $child
                })
            }
        }
    }

    $uniqueById = @{}
    foreach ($item in $result) {
        $id = Get-FolderId -Folder $item
        if (-not [string]::IsNullOrWhiteSpace($id) -and -not $uniqueById.ContainsKey($id)) {
            $uniqueById[$id] = $item
        }
    }

    return @(Sort-FoldersByLabel -Folders @($uniqueById.Values))
}

function Select-ProjectFromConsole {
    Write-Log "Carregando pastas raiz do ProjectWise." "INFO"
    $rootFolders = Get-PWRootFoldersSafe
    $engineeringFolder = Find-FolderByPossibleNames -Folders $rootFolders -PossibleNames $NomesPossiveisEngenhariaRaiz -Description "a pasta raiz de engenharia"

    $engineeringFolderId = Get-FolderId -Folder $engineeringFolder
    if ([string]::IsNullOrWhiteSpace($engineeringFolderId)) {
        throw "Nao foi possivel identificar o ID da pasta de engenharia."
    }

    $concessions = Sort-FoldersByLabel -Folders (Get-PWChildFoldersSafe -FolderId $engineeringFolderId -Context "Engenharia")
    if ($concessions.Count -eq 0) {
        throw "Nenhuma concessao encontrada dentro da pasta Engenharia."
    }

    $selectedConcession = Select-FolderFromList -Folders $concessions -Title "Concessoes encontradas:" -Prompt "Selecione a concessao"
    $selectedConcessionLabel = Get-FolderLabel -Folder $selectedConcession
    Write-Log "Concessao selecionada: $selectedConcessionLabel" "OK"

    $selectedConcessionId = Get-FolderId -Folder $selectedConcession
    if ([string]::IsNullOrWhiteSpace($selectedConcessionId)) {
        throw "Nao foi possivel identificar o ID da concessao selecionada."
    }

    $concessionChildren = Get-PWChildFoldersSafe -FolderId $selectedConcessionId -Context $selectedConcessionLabel
    $projectsFolder = Find-FolderByPossibleNames -Folders $concessionChildren -PossibleNames $NomesPossiveisPastaProjetos -Description "a pasta de projetos da concessao '$selectedConcessionLabel'"

    $projects = Get-ProjectsFromProjectsFolder -ProjectsFolder $projectsFolder -MaximumDepth $ProfundidadeProjetos
    if ($projects.Count -eq 0) {
        throw "Nenhum projeto encontrado na concessao '$selectedConcessionLabel'."
    }

    $selectedProject = Select-FolderFromList -Folders $projects -Title "Projetos encontrados em '$selectedConcessionLabel':" -Prompt "Selecione o projeto"

    Write-Log "Projeto selecionado: $(Get-FolderLabel -Folder $selectedProject)" "OK"
    return $selectedProject
}

function Select-DocumentFolderFromProject {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProjectFolder
    )

    $currentFolder = $ProjectFolder
    $currentTrail = @(Get-FolderLabel -Folder $ProjectFolder)

    while ($true) {
        $currentFolderId = Get-FolderId -Folder $currentFolder
        if ([string]::IsNullOrWhiteSpace($currentFolderId)) {
            throw "Nao foi possivel identificar o ID da pasta atual."
        }

        $children = Sort-FoldersByLabel -Folders (Get-PWChildFoldersSafe -FolderId $currentFolderId -Context ($currentTrail -join "\"))

        Write-Host ""
        Write-Host "Pasta atual: $($currentTrail -join '\')" -ForegroundColor Cyan
        Write-Host "000. Usar esta pasta para buscar documentos"

        for ($i = 0; $i -lt $children.Count; $i++) {
            "{0:000}. {1}" -f ($i + 1), (Get-FolderLabel -Folder $children[$i]) | Write-Host
        }

        if ($children.Count -eq 0) {
            Write-Log "A pasta atual nao possui subpastas. Ela sera usada para buscar documentos." "WARN"
            return $currentFolder
        }

        $selection = Read-NumberSelection -Message "Selecione uma subpasta ou 0 para usar a pasta atual" -Minimum 0 -Maximum $children.Count
        if ($selection -eq 0) {
            return $currentFolder
        }

        $currentFolder = $children[$selection - 1]
        $currentTrail += (Get-FolderLabel -Folder $currentFolder)
    }
}

function Read-ProcessingLimit {
    param(
        [int]$TotalDocuments
    )

    Write-Host ""
    Write-Host ("Documentos encontrados na pasta selecionada: {0}" -f $TotalDocuments) -ForegroundColor Cyan
    Write-Host "Informe quantos documentos deseja processar nesta execucao." -ForegroundColor Cyan
    Write-Host "Use 0 para processar todos." -ForegroundColor Cyan

    while ($true) {
        $inputValue = (Read-Host "Quantidade para processar").Trim()

        if ($inputValue -match '^\d+$') {
            $limit = [int]$inputValue

            if ($limit -eq 0) {
                return 0
            }

            if ($limit -ge 1 -and $limit -le $TotalDocuments) {
                return $limit
            }
        }

        Write-Log "Informe 0 ou um numero entre 1 e $TotalDocuments." "WARN"
    }
}

function Get-DocumentsFromSelectedFolder {
    param(
        [string]$FolderId,
        [string]$FolderPath,
        [string]$FolderLabel
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    if ($ModoBuscaRapida) {
        Write-Log "Consultando documentos no ProjectWise em modo rapido, sem carregar atributos de todos os documentos." "INFO"
    }
    else {
        Write-Log "Consultando documentos no ProjectWise com atributos. Esta etapa pode demorar." "INFO"
    }
    Write-Host "[BUSCA] Aguarde... buscando documentos na pasta '$FolderLabel'." -ForegroundColor Cyan

    try {
        if (-not [string]::IsNullOrWhiteSpace($FolderId)) {
            Write-Log "Buscando documentos por FolderID: $FolderId" "INFO"
            if ($ModoBuscaRapida) {
                $foundDocuments = @(Get-PWDocumentsBySearch -FolderID ([int]$FolderId) -JustThisFolder -PopulatePath -ErrorAction Stop)
            }
            else {
                $foundDocuments = @(Get-PWDocumentsBySearch -FolderID ([int]$FolderId) -JustThisFolder -GetAttributes -PopulatePath -ErrorAction Stop)
            }
        }
        else {
            Write-Log "Buscando documentos por caminho: $FolderPath" "INFO"
            if ($ModoBuscaRapida) {
                $foundDocuments = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
            }
            else {
                $foundDocuments = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -GetAttributes -PopulatePath -ErrorAction Stop)
            }
        }

        Write-Host ""
        Write-Host ("[CONTAGEM] Documentos encontrados na pasta selecionada: {0}" -f $foundDocuments.Count) -ForegroundColor Cyan

        if ($foundDocuments.Count -eq 0) {
            Write-Log "Nenhum documento foi encontrado diretamente nesta pasta. Verifique se os arquivos estao nesta pasta ou em uma subpasta." "WARN"
        }

        return $foundDocuments
    }
    finally {
        $stopwatch.Stop()
        Write-Log ("Busca finalizada em {0:n1} segundos." -f $stopwatch.Elapsed.TotalSeconds) "OK"
    }
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
    $attributeContainers = @("CustomAttributes", "Attributes", "DocumentAttributes", "EnvironmentAttributes")
    foreach ($containerName in $attributeContainers) {
        $container = Get-DocumentPropertyValue -Document $Document -PropertyNames @($containerName)
        if ($null -eq $container) {
            continue
        }

        if ($container -is [System.Collections.IDictionary]) {
            $containsKeyMethod = $container.PSObject.Methods["ContainsKey"]
            if ($null -ne $containsKeyMethod -and $container.ContainsKey($AttributeName)) {
                return $true
            }

            if ($null -ne $container.Keys -and $container.Keys -contains $AttributeName) {
                return $true
            }
        }

        $attributeProperty = $container.PSObject.Properties[$AttributeName]
        if ($null -ne $attributeProperty) {
            return $true
        }
    }

    return $false
}

function Get-AttributeContainerKeys {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document
    )

    $keys = @()
    $attributeContainers = @("CustomAttributes", "Attributes", "DocumentAttributes", "EnvironmentAttributes")

    foreach ($containerName in $attributeContainers) {
        $container = Get-DocumentPropertyValue -Document $Document -PropertyNames @($containerName)
        if ($null -eq $container) {
            continue
        }

        if ($null -ne $container.Keys) {
            $keys += @($container.Keys | ForEach-Object { [string]$_ })
            continue
        }

        $keys += @($container.PSObject.Properties | ForEach-Object { $_.Name })
    }

    return @($keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Test-AttributeExistsOnce {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Documents,

        [string]$FolderId,

        [string]$FolderPath,

        [Parameter(Mandatory = $true)]
        [string]$AttributeName
    )

    if (-not $ValidarAtributoUmaVez) {
        Write-Log "Validacao unica do atributo desativada." "WARN"
        return $true
    }

    if ($Documents.Count -eq 0) {
        Write-Log "Sem documentos para validar o atributo '$AttributeName'." "WARN"
        return $true
    }

    $sampleDocument = $null
    foreach ($document in $Documents) {
        $sampleName = Get-DocumentPropertyValue -Document $document -PropertyNames @("Name", "DocumentName")
        if (-not [string]::IsNullOrWhiteSpace([string]$sampleName)) {
            $sampleDocument = $document
            break
        }
    }

    if ($null -eq $sampleDocument) {
        Write-Log "Nenhum documento com Name preenchido foi encontrado para validar o atributo '$AttributeName'." "WARN"
        return $true
    }

    $sampleDocumentName = Get-DocumentPropertyValue -Document $sampleDocument -PropertyNames @("Name", "DocumentName")

    Write-Log "Validando existencia do atributo '$AttributeName' em um documento de amostra: $sampleDocumentName" "INFO"

    $documentsWithAttributes = @()
    try {
        if (-not [string]::IsNullOrWhiteSpace($FolderId)) {
            $documentsWithAttributes = @(Get-PWDocumentsBySearch -FolderID ([int]$FolderId) -JustThisFolder -DocumentName ([string]$sampleDocumentName) -GetAttributes -PopulatePath -ErrorAction Stop)
        }
        else {
            $documentsWithAttributes = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName ([string]$sampleDocumentName) -GetAttributes -PopulatePath -ErrorAction Stop)
        }
    }
    catch {
        Write-Log "Falha na validacao unica do atributo. O lote continuara e eventuais falhas serao tratadas por documento. Erro: $($_.Exception.Message)" "WARN"
        return $true
    }

    if ($documentsWithAttributes.Count -eq 0) {
        Write-Log "Nao foi possivel recarregar o documento de amostra com atributos. O lote continuara com tratamento individual." "WARN"
        return $true
    }

    if (-not (Test-AttributeExists -Document $documentsWithAttributes[0] -AttributeName $AttributeName)) {
        throw "O atributo '$AttributeName' nao foi encontrado no ambiente desta pasta. Processamento interrompido antes do lote."
    }

    Write-Log "Atributo '$AttributeName' validado com sucesso no documento de amostra." "OK"
    return $true
}

function ConvertTo-SimpleName {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($character in $normalized.ToCharArray()) {
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($character)
        }
    }

    return (($builder.ToString() -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
}

function Get-DocumentWithAttributesByName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DocumentName,

        [string]$FolderId,

        [string]$FolderPath
    )

    if (-not [string]::IsNullOrWhiteSpace($FolderId)) {
        return @(Get-PWDocumentsBySearch -FolderID ([int]$FolderId) -JustThisFolder -DocumentName $DocumentName -GetAttributes -PopulatePath -ErrorAction Stop)
    }

    return @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -GetAttributes -PopulatePath -ErrorAction Stop)
}

function Resolve-NumeroDocumentoAttributeColumnName {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Documents,

        [object]$Folder,

        [string]$FolderId,

        [string]$FolderPath,

        [Parameter(Mandatory = $true)]
        [string]$DisplayAttributeName,

        [string]$ConfiguredColumnName,

        [string]$ConfiguredEnvironmentName
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredColumnName)) {
        Write-Log "Usando nome interno configurado para o atributo '$DisplayAttributeName': $ConfiguredColumnName" "INFO"
        return $ConfiguredColumnName
    }

    if ($Documents.Count -eq 0) {
        return $DisplayAttributeName
    }

    $sampleDocumentName = ""
    foreach ($document in $Documents) {
        $candidateName = Get-DocumentPropertyValue -Document $document -PropertyNames @("Name", "DocumentName")
        if (-not [string]::IsNullOrWhiteSpace([string]$candidateName)) {
            $sampleDocumentName = [string]$candidateName
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($sampleDocumentName)) {
        Write-Log "Nao foi possivel obter um documento de amostra para descobrir o nome interno do atributo. Sera usado '$DisplayAttributeName'." "WARN"
        return $DisplayAttributeName
    }

    Write-Log "Descobrindo nome interno do atributo '$DisplayAttributeName' em um documento de amostra: $sampleDocumentName" "INFO"

    try {
        $sampleDocuments = @(Get-DocumentWithAttributesByName -DocumentName $sampleDocumentName -FolderId $FolderId -FolderPath $FolderPath)
        $sampleDocument = $null
        if ($sampleDocuments.Count -gt 0) {
            $sampleDocument = $sampleDocuments[0]
        }

        $environmentName = $ConfiguredEnvironmentName
        if ([string]::IsNullOrWhiteSpace($environmentName)) {
            $environmentName = Get-EnvironmentNameFromObject -InputObject $Folder
        }
        if ([string]::IsNullOrWhiteSpace($environmentName) -and $null -ne $sampleDocument) {
            $environmentName = Get-EnvironmentNameFromObject -InputObject $sampleDocument
        }

        if ([string]::IsNullOrWhiteSpace($environmentName)) {
            Write-Log "Nao foi possivel identificar automaticamente o Environment Name da pasta/documento." "WARN"
            $environmentName = (Read-Host "Informe o Environment Name da pasta ou pressione Enter para pular esta etapa").Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($environmentName)) {
            Write-Log "Consultando colunas do ambiente ProjectWise: $environmentName" "INFO"
            $environmentColumns = @(Get-PWEnvironmentColumns -EnvironmentName $environmentName -ErrorAction Stop)

            if ($environmentColumns.Count -gt 0) {
                $displaySimpleName = ConvertTo-SimpleName -Text $DisplayAttributeName
                $expectedSimpleNames = @(
                    $displaySimpleName,
                    (ConvertTo-SimpleName -Text "NumeroDocumento"),
                    (ConvertTo-SimpleName -Text "Numero Documento"),
                    (ConvertTo-SimpleName -Text "N Documento"),
                    (ConvertTo-SimpleName -Text "NDocumento"),
                    (ConvertTo-SimpleName -Text "NoDocumento"),
                    (ConvertTo-SimpleName -Text "NumDocumento"),
                    (ConvertTo-SimpleName -Text "NrDocumento")
                )

                foreach ($column in $environmentColumns) {
                    $allTextValues = @(
                        $column.PSObject.Properties |
                            Where-Object { $null -ne $_.Value -and -not [string]::IsNullOrWhiteSpace([string]$_.Value) } |
                            ForEach-Object { [string]$_.Value }
                    )

                    $matched = $false
                    foreach ($textValue in $allTextValues) {
                        if ($expectedSimpleNames -contains (ConvertTo-SimpleName -Text $textValue)) {
                            $matched = $true
                            break
                        }
                    }

                    if ($matched) {
                        $internalName = Get-ObjectPropertyValue -InputObject $column -PropertyNames @(
                            "ColumnName",
                            "Name",
                            "AttributeName",
                            "EnvironmentAttributeName",
                            "Column",
                            "ColumnLabel",
                            "DatabaseColumn",
                            "DbColumnName"
                        )

                        if (-not [string]::IsNullOrWhiteSpace([string]$internalName)) {
                            Write-Log "Nome interno encontrado no ambiente '$environmentName': $internalName" "OK"
                            return ([string]$internalName).Trim()
                        }
                    }
                }

                $likelyColumns = @(
                    $environmentColumns |
                        Where-Object {
                            $values = @($_.PSObject.Properties | ForEach-Object { [string]$_.Value })
                            ($values -join " ") -match '(?i)doc|documento|numero|num|nr'
                        } |
                        Select-Object -First 20
                )

                if ($likelyColumns.Count -gt 0) {
                    Write-Log "Colunas candidatas do ambiente '$environmentName':" "WARN"
                    foreach ($candidate in $likelyColumns) {
                        $candidateText = @(
                            $candidate.PSObject.Properties |
                                Where-Object { $null -ne $_.Value -and -not [string]::IsNullOrWhiteSpace([string]$_.Value) } |
                                ForEach-Object { "$($_.Name)=$($_.Value)" }
                        ) -join " | "
                        Write-Host "  $candidateText" -ForegroundColor Yellow
                    }
                }
            }
            else {
                Write-Log "Nenhuma coluna foi retornada para o ambiente '$environmentName'." "WARN"
            }
        }

        if ($sampleDocuments.Count -eq 0) {
            Write-Log "Nao foi possivel recarregar o documento de amostra com atributos. Sera usado '$DisplayAttributeName'." "WARN"
            return $DisplayAttributeName
        }

        $attributeKeys = @(Get-AttributeContainerKeys -Document $sampleDocument)
        if ($attributeKeys.Count -gt 0) {
            $expectedAttributeKeys = @(
                $DisplayAttributeName,
                "NumeroDocumento",
                "Numero Documento",
                "N Documento",
                "NDocumento",
                "NoDocumento",
                "NumDocumento",
                "NrDocumento",
                "NomeCompleto",
                "Label"
            )

            $expectedAttributeSimpleKeys = @($expectedAttributeKeys | ForEach-Object { ConvertTo-SimpleName -Text $_ })
            $attributeKeyMatch = @(
                $attributeKeys |
                    Where-Object { $expectedAttributeSimpleKeys -contains (ConvertTo-SimpleName -Text $_) } |
                    Select-Object -First 1
            )

            if ($attributeKeyMatch.Count -gt 0) {
                Write-Log "Nome interno encontrado nos atributos carregados do documento: $($attributeKeyMatch[0])" "OK"
                return $attributeKeyMatch[0]
            }

            $likelyAttributeKeys = @(
                $attributeKeys |
                    Where-Object { $_ -match '(?i)doc|documento|numero|num|nr|nome|label' } |
                    Select-Object -First 25
            )

            if ($likelyAttributeKeys.Count -gt 0) {
                Write-Log "Chaves candidatas em CustomAttributes/Attributes: $($likelyAttributeKeys -join ', ')" "WARN"
            }
        }

        $propertyNames = @($sampleDocument.PSObject.Properties | ForEach-Object { $_.Name })
        $exactMatch = @($propertyNames | Where-Object { $_ -ceq $DisplayAttributeName } | Select-Object -First 1)
        if ($exactMatch.Count -gt 0) {
            Write-Log "Nome interno encontrado por correspondencia exata: $($exactMatch[0])" "OK"
            return $exactMatch[0]
        }

        $expectedNames = @(
            $DisplayAttributeName,
            "NumeroDocumento",
            "Numero Documento",
            "N Documento",
            "NDocumento",
            "NoDocumento",
            "NumDocumento",
            "NrDocumento"
        )

        $expectedSimpleNames = @($expectedNames | ForEach-Object { ConvertTo-SimpleName -Text $_ })
        $normalizedMatch = @(
            $propertyNames |
                Where-Object { $expectedSimpleNames -contains (ConvertTo-SimpleName -Text $_) } |
                Select-Object -First 1
        )

        if ($normalizedMatch.Count -gt 0) {
            Write-Log "Nome interno encontrado para '$DisplayAttributeName': $($normalizedMatch[0])" "OK"
            return $normalizedMatch[0]
        }

        $likelyColumns = @(
            $propertyNames |
                Where-Object { $_ -match '(?i)doc|documento|numero|num|nr' } |
                Select-Object -First 25
        )

        if ($likelyColumns.Count -gt 0) {
            Write-Log "Nao consegui identificar automaticamente. Colunas candidatas encontradas: $($likelyColumns -join ', ')" "WARN"
        }
        else {
            Write-Log "Nao consegui identificar automaticamente o nome interno do atributo no documento de amostra." "WARN"
        }

        $typedColumnName = (Read-Host "Informe o nome interno da coluna para '$DisplayAttributeName' ou pressione Enter para usar o nome exibido").Trim()
        if (-not [string]::IsNullOrWhiteSpace($typedColumnName)) {
            Write-Log "Nome interno informado manualmente: $typedColumnName" "OK"
            return $typedColumnName
        }

        Write-Log "Sera usado o nome exibido '$DisplayAttributeName'. Se a atualizacao falhar, configure `$NumeroDocumentoAttributeColumnName no inicio do script." "WARN"
        return $DisplayAttributeName
    }
    catch {
        Write-Log "Falha ao descobrir o nome interno do atributo. Sera usado '$DisplayAttributeName'. Erro: $($_.Exception.Message)" "WARN"
        return $DisplayAttributeName
    }
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

function Update-NumeroDocumentoAttribute {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,

        [Parameter(Mandatory = $true)]
        [string]$DocumentName,

        [Parameter(Mandatory = $true)]
        [string]$AttributeColumnName,

        [Parameter(Mandatory = $true)]
        [string]$AttributeDisplayName,

        [Parameter(Mandatory = $true)]
        [string]$Value,

        [string]$FolderId,

        [string]$FolderPath
    )

    $attributesToUpdate = @{
        $AttributeColumnName = [string]$Value
    }

    Write-Log "Tentando atualizar atributo. Documento='$DocumentName' | Coluna='$AttributeColumnName' | Valor='$Value'." "INFO"
    $updateWarnings = $null
    $updateResult = Update-PWDocumentAttributes -InputDocuments @($Document) -Attributes $attributesToUpdate -ReturnBoolean -WarningAction SilentlyContinue -WarningVariable updateWarnings -ErrorAction Stop
    if ($updateWarnings) {
        Write-Log "Avisos retornados pelo Update-PWDocumentAttributes: $($updateWarnings -join ' | ')" "WARN"
    }

    Write-Log "Retorno da primeira tentativa Update-PWDocumentAttributes: $updateResult" "INFO"

    if ($updateResult -ne $false) {
        return $true
    }

    if ($ModoBuscaRapida) {
        Write-Log "Primeira tentativa retornou False. Recarregando '$DocumentName' com atributos e tentando novamente." "WARN"
        $documentsWithAttributes = @(Get-DocumentWithAttributesByName -DocumentName $DocumentName -FolderId $FolderId -FolderPath $FolderPath)

        if ($documentsWithAttributes.Count -gt 0) {
            Write-Log "Documento recarregado com atributos. Tentando atualizar novamente '$DocumentName'." "INFO"
            $retryWarnings = $null
            $updateResult = Update-PWDocumentAttributes -InputDocuments @($documentsWithAttributes[0]) -Attributes $attributesToUpdate -ReturnBoolean -WarningAction SilentlyContinue -WarningVariable retryWarnings -ErrorAction Stop
            if ($retryWarnings) {
                Write-Log "Avisos retornados na segunda tentativa: $($retryWarnings -join ' | ')" "WARN"
            }

            Write-Log "Retorno da segunda tentativa Update-PWDocumentAttributes: $updateResult" "INFO"
            if ($updateResult -ne $false) {
                return $true
            }
        }
        else {
            Write-Log "Nao foi possivel recarregar '$DocumentName' com atributos para segunda tentativa." "WARN"
        }
    }

    throw "Update-PWDocumentAttributes retornou False usando a coluna '$AttributeColumnName' para o atributo '$AttributeDisplayName'. Verifique se esse e o nome interno correto da coluna no ambiente ProjectWise."
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

function Get-NumeroDocumentoFromName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DocumentName
    )

    # Remove a extensao do arquivo, quando existir. Ex.: .dwg, .pdf, .docx.
    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName.Trim())

    if ([string]::IsNullOrWhiteSpace($nameWithoutExtension)) {
        return $null
    }

    # Captura o codigo ate o ultimo bloco numerico apos hifen antes da revisao.
    # Exemplo: ERM-493RJ-112+815-EXE-DE-L3-001-R01a -> ERM-493RJ-112+815-EXE-DE-L3-001
    # O bloco numerico pode variar: 001, 002, 123, 1001 etc.
    $match = [regex]::Match($nameWithoutExtension, "^(?<NumeroDocumento>.+-\d+)(?:-[A-Za-z].*|[A-Za-z].*)?$")

    if ($match.Success) {
        return $match.Groups["NumeroDocumento"].Value
    }

    return $null
}

try {
    Write-Log "Iniciando preenchimento do atributo '$NumeroDocumentoAttributeName'." "INFO"
    Write-Log "Arquivo de log desta execucao: $LogFilePath" "INFO"
    Write-Log "Script: $PSCommandPath" "INFO"
    Write-Log "PowerShell: $($PSVersionTable.PSVersion) | ApartmentState: $([System.Threading.Thread]::CurrentThread.GetApartmentState()) | Usuario: $env:USERNAME" "INFO"

    $executionMode = Resolve-ExecutionMode -DryRunWasRequested $DryRunParam.IsPresent -ExecuteWasRequested $ExecuteParam.IsPresent -DefaultDryRun $DryRun
    $DryRun = $executionMode.DryRun
    if ($executionMode.SelectedByPrompt) {
        if ($DryRun) {
            Write-Log "Modo selecionado no console: SIMULACAO. Nenhum documento sera alterado." "DRYRUN"
        }
        else {
            Write-Log "Modo selecionado no console: EXECUCAO REAL. Documentos poderao ser alterados no ProjectWise." "WARN"
        }
    }

    Write-Log "Parametros: ProjectPathParam='$ProjectPathParam' | DryRunParam=$($DryRunParam.IsPresent) | ExecuteParam=$($ExecuteParam.IsPresent) | LimiteDocumentosParam=$LimiteDocumentosParam" "INFO"
    Write-Log "Configuracao efetiva: ProjectPath='$ProjectPath' | ImportFolderName='$ImportFolderName' | DryRun=$DryRun | LimiteDocumentos=$LimiteDocumentos | ModoBuscaRapida=$ModoBuscaRapida | DatasourceName='$DatasourceName' | UsarBentleyIMS=$UsarBentleyIMS" "INFO"
    Write-Log "Atributo alvo: Exibicao='$NumeroDocumentoAttributeName' | ColunaConfigurada='$NumeroDocumentoAttributeColumnName' | EnvironmentConfigurado='$ProjectWiseEnvironmentName'" "INFO"

    Write-Log "Importando modulo PWPS_DAB." "INFO"
    Import-Module PWPS_DAB -ErrorAction Stop
    $pwpsModule = Get-Module PWPS_DAB
    if ($null -ne $pwpsModule) {
        Write-Log "Modulo PWPS_DAB carregado. Versao=$($pwpsModule.Version) | Caminho=$($pwpsModule.Path)" "OK"
    }

    Connect-ProjectWiseDatasource

    $targetFolder = $null
    $targetFolderId = ""
    $targetFolderPath = ""

    if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
        $selectedProject = Select-ProjectFromConsole
        $targetFolder = Select-DocumentFolderFromProject -ProjectFolder $selectedProject
        $targetFolderId = Get-FolderId -Folder $targetFolder
        $targetFolderPath = Get-FolderPathValue -Folder $targetFolder

        if ([string]::IsNullOrWhiteSpace($targetFolderId)) {
            throw "Nao foi possivel identificar o ID da pasta selecionada."
        }

        Write-Log "Pasta selecionada para buscar documentos: $(Get-FolderLabel -Folder $targetFolder)" "OK"
    }
    else {
        $targetFolderPath = Join-PWPath -ParentPath $ProjectPath -ChildPath $ImportFolderName
        Write-Log "Caminho do projeto informado: $ProjectPath" "INFO"
        Write-Log "Validando subpasta '$ImportFolderName': $targetFolderPath" "INFO"

        $targetFolder = Get-PWFolders -FolderPath $targetFolderPath -JustOne -ErrorAction Stop
        if ($null -eq $targetFolder) {
            throw "A subpasta '$ImportFolderName' nao foi encontrada no projeto informado."
        }

        $targetFolderId = Get-FolderId -Folder $targetFolder
        Write-Log "Subpasta '$ImportFolderName' localizada com sucesso." "OK"
    }

    if ($DryRun) {
        Write-Log "Modo simulacao ativo. Nenhum documento sera alterado." "DRYRUN"
    }

    # Busca somente documentos existentes diretamente na pasta selecionada.
    # Em modo rapido, os atributos nao sao carregados para todos os documentos.
    $targetFolderLabel = if ($null -ne $targetFolder) { Get-FolderLabel -Folder $targetFolder } else { $targetFolderPath }
    $documents = @(Get-DocumentsFromSelectedFolder -FolderId $targetFolderId -FolderPath $targetFolderPath -FolderLabel $targetFolderLabel)

    $totalFound = $documents.Count

    if ($totalFound -gt 0 -and $LimiteDocumentosParam -lt 0) {
        $LimiteDocumentos = Read-ProcessingLimit -TotalDocuments $totalFound
    }

    if ($LimiteDocumentos -gt 0 -and $documents.Count -gt $LimiteDocumentos) {
        Write-Log "Limite de processamento ativo: somente os primeiros $LimiteDocumentos documentos serao processados." "WARN"
        $documents = @($documents | Select-Object -First $LimiteDocumentos)
    }

    $totalToProcess = $documents.Count
    $successCount = 0
    $errorCount = 0
    $ignoredCount = 0
    $unitMismatchCount = 0
    $processedCount = 0

    Write-Log "Documentos encontrados: $totalFound" "INFO"
    Write-Log "Documentos selecionados para processamento: $totalToProcess" "INFO"

    if ($ModoBuscaRapida -and $ValidarAtributoUmaVez -and $totalToProcess -gt 0) {
        Test-AttributeExistsOnce -Documents $documents -FolderId $targetFolderId -FolderPath $targetFolderPath -AttributeName $NumeroDocumentoAttributeName | Out-Null
    }

    $numeroDocumentoAttributeColumn = $NumeroDocumentoAttributeName
    if ($totalToProcess -gt 0) {
        $numeroDocumentoAttributeColumn = Resolve-NumeroDocumentoAttributeColumnName -Documents $documents -Folder $targetFolder -FolderId $targetFolderId -FolderPath $targetFolderPath -DisplayAttributeName $NumeroDocumentoAttributeName -ConfiguredColumnName $NumeroDocumentoAttributeColumnName -ConfiguredEnvironmentName $ProjectWiseEnvironmentName
        Write-Log "Atualizacao usara a coluna interna '$numeroDocumentoAttributeColumn' para preencher '$NumeroDocumentoAttributeName'." "INFO"
    }

    foreach ($document in $documents) {
        $processedCount++
        $displayName = Get-DocumentDisplayName -Document $document
        $percentComplete = if ($totalToProcess -gt 0) {
            [int](($processedCount / $totalToProcess) * 100)
        }
        else {
            100
        }

        Write-Progress -Activity "Ajuste de atributos ProjectWise" -Status ("Processando {0} de {1}: {2}" -f $processedCount, $totalToProcess, $displayName) -PercentComplete $percentComplete

        try {
            Write-Log ("Processando documento {0}/{1}: {2}" -f $processedCount, $totalToProcess, $displayName) "INFO"

            # Le o campo Name da aba General.
            $documentName = Get-DocumentPropertyValue -Document $document -PropertyNames @("Name", "DocumentName")
            if ([string]::IsNullOrWhiteSpace([string]$documentName)) {
                $ignoredCount++
                Write-Log "Ignorado: campo Name vazio." "WARN"
                continue
            }

            # Extrai do Name somente o codigo do documento, removendo extensao e revisao.
            $numeroDocumento = Get-NumeroDocumentoFromName -DocumentName ([string]$documentName)
            if ([string]::IsNullOrWhiteSpace([string]$numeroDocumento)) {
                $ignoredCount++
                Write-Log "Ignorado: nao foi possivel extrair o numero do documento a partir do Name '$documentName'." "WARN"
                continue
            }

            Write-Log "Name original: '$documentName' | Valor para '$NumeroDocumentoAttributeName': '$numeroDocumento'" "INFO"

            # Evita alterar documentos bloqueados/check-out.
            if (Test-DocumentIsCheckedOut -Document $document) {
                $ignoredCount++
                Write-Log "Ignorado: documento bloqueado ou em check-out." "WARN"
                continue
            }

            # Confirma que o atributo existe quando os atributos foram carregados na busca inicial.
            if (-not $ModoBuscaRapida -and -not (Test-AttributeExists -Document $document -AttributeName $NumeroDocumentoAttributeName)) {
                $ignoredCount++
                Write-Log "Ignorado: atributo '$NumeroDocumentoAttributeName' nao encontrado." "WARN"
                continue
            }

            if ($DryRun) {
                $successCount++
                Write-Log "Simulacao: atributo '$NumeroDocumentoAttributeName' seria preenchido pela coluna '$numeroDocumentoAttributeColumn' com '$numeroDocumento'." "DRYRUN"
                continue
            }

            # Atualiza e salva os atributos do documento no ProjectWise.
            Update-NumeroDocumentoAttribute -Document $document -DocumentName ([string]$documentName) -AttributeColumnName $numeroDocumentoAttributeColumn -AttributeDisplayName $NumeroDocumentoAttributeName -Value ([string]$numeroDocumento) -FolderId $targetFolderId -FolderPath $targetFolderPath | Out-Null

            $successCount++
            Write-Log "Atualizado com sucesso: '$NumeroDocumentoAttributeName' = '$numeroDocumento'." "OK"
        }
        catch {
            $errorMessage = $_.Exception.Message

            if ($errorMessage -like "*(01) Unidade do documento diferente da unidade da pasta*") {
                $unitMismatchCount++
                Write-ExceptionLog -ErrorRecord $_ -Context "Falha de validacao ProjectWise em '$displayName': unidade do documento diferente da unidade da pasta"
                continue
            }

            $errorCount++
            Write-ExceptionLog -ErrorRecord $_ -Context "Falha ao processar '$displayName'"
        }
    }

    Write-Progress -Activity "Ajuste de atributos ProjectWise" -Completed

    Write-Host ""
    Write-Log "Resumo final" "INFO"
    Write-Log "Documentos encontrados: $totalFound" "INFO"
    Write-Log "Documentos processados nesta execucao: $totalToProcess" "INFO"
    Write-Log "Atualizados com sucesso: $successCount" "OK"
    Write-Log "Ignorados: $ignoredCount" "WARN"
    Write-Log "Com divergencia de unidade documento/pasta: $unitMismatchCount" "ERROR"
    Write-Log "Com erro: $errorCount" "ERROR"

    if ($DryRun) {
        Write-Log "Execucao finalizada em modo simulacao. Para alterar documentos, execute novamente e escolha E no modo de execucao." "DRYRUN"
        Write-Log "Opcionalmente, execute com -ExecuteParam para pular a pergunta e forcar a atualizacao real." "DRYRUN"
    }
}
catch {
    Write-ExceptionLog -ErrorRecord $_ -Context "Erro fatal"
    exit 1
}
