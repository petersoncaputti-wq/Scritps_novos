<#
.SYNOPSIS
Exporta dados dos projetos do ProjectWise no layout "Modelo de criacao de projeto".

.DESCRIPTION
O script busca Work Areas/Rich Projects no ProjectWise e gera um CSV com as colunas
usadas para criacao/validacao de projetos:

Nome Concessao; Sigla Concessao; Nome Projeto; Descricao; Projetista;
Gestor Engenharia; Assistente Engenharia; Poder Concedente; TH

Mapeamento usado na aba Propriedades do projeto:
- Nome Concessao = NomeConcessao / NomeConcessão
- Sigla Concessao = Nova Sigla
- Projetista = Nome Projetista
- Gestor Engenharia = Gestor da Engenharia
- Assistente Engenharia = Assistente
- Poder Concedente = P. Concedente
- TH = TH

.EXAMPLE
.\07 - Exportar Modelo de Criacao de Projetos ProjectWise.ps1 -UseGui

.EXAMPLE
.\07 - Exportar Modelo de Criacao de Projetos ProjectWise.ps1 -DatasourceName "Servidor:Datasource" -BentleyIMS -NonAdminLogin

.EXAMPLE
.\07 - Exportar Modelo de Criacao de Projetos ProjectWise.ps1 -DatasourceName "Servidor:Datasource" -FolderPath "Projetos\Ecovias Araguaia" -OutputPath ".\exports\projetos.csv"

.EXAMPLE
.\07 - Exportar Modelo de Criacao de Projetos ProjectWise.ps1 -UseGui -InteractiveSelection
#>

[CmdletBinding()]
param(
    [string]$DatasourceName,
    [string]$FolderPath,
    [string]$ConcessionsRootPath = 'Engenharia',
    [string[]]$EngineeringRootNames = @('Engenharia', 'Engineering'),
    [string[]]$ProjectsFolderNames = @('Projetos', 'Projeto', 'Projects', 'Project'),
    [string]$ProjectTypeName,
    [string]$OutputPath,
    [string]$UserName,
    [securestring]$Password,
    [switch]$UseGui,
    [switch]$BentleyIMS,
    [switch]$NonAdminLogin,
    [switch]$OnlyConnectedProjects,
    [switch]$OnlyNonConnectedProjects,
    [switch]$JustOne,
    [switch]$InteractiveSelection,
    [switch]$KeepSessionOpen,
    [switch]$IncludeDiagnostics,
    [switch]$NaoPausar
)

$previousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
Import-Module pwps_dab -DisableNameChecking -WarningAction SilentlyContinue
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ScriptBasePath = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    (Get-Location).Path
}
else {
    $PSScriptRoot
}

$ExportsPath = Join-Path $ScriptBasePath 'exports'
$LogsPath = Join-Path $ScriptBasePath 'Logs'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ExportsPath ("modelo-criacao-projeto-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

if (-not (Test-Path -LiteralPath $LogsPath)) {
    New-Item -ItemType Directory -Path $LogsPath -Force | Out-Null
}

$LogPath = Join-Path $LogsPath ("ExportarModeloCriacaoProjetos_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

if ([string]::IsNullOrWhiteSpace($DatasourceName) -and -not $UseGui) {
    $UseGui = $true
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $line = "{0} | {1} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpperInvariant(), $Message
    try {
        [System.IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine)
    }
    catch {
    }
}

function Write-Status {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    Write-Host $Message
    Write-Log -Message $Message -Level $Level
}

function ConvertTo-NormalizedKey {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormD)
    $builder = [System.Text.StringBuilder]::new()

    foreach ($char in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark -and [char]::IsLetterOrDigit($char)) {
            [void]$builder.Append([char]::ToLowerInvariant($char))
        }
    }

    return $builder.ToString()
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value -and "$($property.Value)".Trim().Length -gt 0) {
            return $property.Value
        }
    }

    return ''
}

function Get-ProjectPropertyMap {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Project
    )

    $map = @{}

    foreach ($property in $Project.PSObject.Properties) {
        if ($null -ne $property.Name -and $null -ne $property.Value) {
            $key = ConvertTo-NormalizedKey $property.Name
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not $map.ContainsKey($key)) {
                $map[$key] = $property.Value
            }
        }
    }

    $bags = @(
        'ProjectProperties',
        'Properties',
        'CustomProperties',
        'ProjectPropertyValues'
    )

    foreach ($bagName in $bags) {
        $bagProperty = $Project.PSObject.Properties[$bagName]
        if ($null -eq $bagProperty -or $null -eq $bagProperty.Value) {
            continue
        }

        $bag = $bagProperty.Value

        if ($bag -is [System.Collections.IDictionary]) {
            foreach ($key in $bag.Keys) {
                if ($null -ne $key) {
                    $map[(ConvertTo-NormalizedKey "$key")] = $bag[$key]
                }
            }

            continue
        }

        foreach ($item in @($bag)) {
            $name = Get-ObjectPropertyValue -InputObject $item -Names @(
                'Name',
                'PropertyName',
                'ColumnName',
                'AttributeName',
                'Label'
            )
            $value = Get-ObjectPropertyValue -InputObject $item -Names @(
                'Value',
                'PropertyValue',
                'ColumnValue',
                'AttributeValue',
                'DisplayValue'
            )

            if (-not [string]::IsNullOrWhiteSpace("$name")) {
                $map[(ConvertTo-NormalizedKey "$name")] = $value
            }
        }
    }

    return $map
}

function Get-ProjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$PropertyMap,

        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $key = ConvertTo-NormalizedKey $name
        if ($PropertyMap.ContainsKey($key) -and $null -ne $PropertyMap[$key]) {
            $value = "$($PropertyMap[$key])".Trim()
            if ($value.Length -gt 0) {
                return $PropertyMap[$key]
            }
        }
    }

    return ''
}

function Get-ProjectPath {
    param([object]$Project)

    $path = Get-ObjectPropertyValue -InputObject $Project -Names @('FullPath', 'Path', 'FolderPath')
    if (-not [string]::IsNullOrWhiteSpace("$path")) {
        return $path
    }

    try {
        if ($null -ne $Project.PSObject.Methods['GetFullPath']) {
            $fullPath = $Project.GetFullPath()
            if ($null -ne $fullPath) {
                return $fullPath
            }
        }
    }
    catch {
        return ''
    }

    return ''
}

function Get-PathSegment {
    param(
        [string]$Path,
        [int]$Index
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    $segments = @($Path -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($segments.Count -le $Index) {
        return ''
    }

    return $segments[$Index]
}

function Get-ConcessionAcronym {
    param([string]$ConcessionName)

    $known = @{
        'ecoviasaraguaia' = 'ECA'
        'ecoviasdocerrado' = 'ECO050'
        'ecoviasdosimigrantes' = 'ECOVIAS'
        'ecopistas' = 'ECP'
        'ecoponte' = 'ECPONTE'
        'ecosul' = 'ECS'
        'ecoriominas' = 'ECRIO'
        'ecorodovias' = 'ECO'
    }

    $key = ConvertTo-NormalizedKey $ConcessionName
    if ($known.ContainsKey($key)) {
        return $known[$key]
    }

    return ''
}

function Get-DisplayName {
    param([object]$Item)

    return Get-ObjectPropertyValue -InputObject $Item -Names @('Name', 'FolderName', 'ProjectName')
}

function Get-DisplayPath {
    param([object]$Item)

    $path = Get-ProjectPath -Project $Item
    if (-not [string]::IsNullOrWhiteSpace("$path")) {
        return $path
    }

    return Get-DisplayName -Item $Item
}

function Get-FolderId {
    param([object]$Folder)

    return Get-ObjectPropertyValue -InputObject $Folder -Names @('ProjectID', 'FolderID', 'Id', 'ID')
}

function Get-PWRootFoldersSafe {
    $folders = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    $folders = @($folders | Where-Object { $null -ne $_ })

    if ($folders.Count -eq 0) {
        throw 'Nenhuma pasta foi retornada na raiz do datasource.'
    }

    return $folders
}

function Get-PWChildFoldersSafe {
    param([string]$FolderId)

    $folders = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
    $folders = @($folders | Where-Object { $null -ne $_ })

    if ($folders.Count -eq 0) {
        throw "Nenhuma subpasta foi retornada para o FolderID $FolderId."
    }

    return $folders
}

function Find-FolderByPossibleNames {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Folders,

        [Parameter(Mandatory = $true)]
        [string[]]$PossibleNames,

        [string]$Description = 'Pasta'
    )

    if ($Folders.Count -eq 0) {
        throw "$Description nao encontrada. Nenhuma pasta foi retornada para pesquisa."
    }

    $normalizedNames = @($PossibleNames | ForEach-Object { ConvertTo-NormalizedKey $_ })

    $found = $Folders | Where-Object {
        $name = ConvertTo-NormalizedKey (Get-DisplayName -Item $_)
        $normalizedNames -contains $name
    } | Select-Object -First 1

    if (-not $found) {
        $found = $Folders | Where-Object {
            $name = ConvertTo-NormalizedKey (Get-DisplayName -Item $_)
            foreach ($possible in $normalizedNames) {
                if ($name -like "*$possible*") {
                    return $true
                }
            }

            return $false
        } | Select-Object -First 1
    }

    if (-not $found) {
        $options = @($Folders | ForEach-Object { Get-DisplayName -Item $_ }) -join ', '
        throw "$Description nao encontrada. Nomes esperados: $($PossibleNames -join ' / '). Encontradas: $options"
    }

    return $found
}

function Get-RichProjectFromFolder {
    param([object]$Folder)

    $folderPath = Get-DisplayPath -Item $Folder
    if ([string]::IsNullOrWhiteSpace("$folderPath")) {
        return $Folder
    }

    try {
        $folderWithProperties = Get-PWFolders -FolderPath $folderPath -JustOne -ErrorAction Stop | Get-PWFolderPathAndProperties -ErrorAction Stop
        if ($null -ne $folderWithProperties) {
            return $folderWithProperties
        }
    }
    catch {
        Write-Status -Message "Aviso: nao foi possivel ler propriedades de '$folderPath' via Get-PWFolderPathAndProperties. Tentando Get-PWRichProjects." -Level 'WARN'
    }

    try {
        $richProject = Get-PWRichProjects -FolderPath $folderPath -JustOne -PopulatePaths -PopulateProjectProperties -ErrorAction Stop
        if ($null -ne $richProject) {
            return $richProject
        }
    }
    catch {
        Write-Status -Message "Aviso: nao foi possivel ler propriedades de '$folderPath' via Get-PWRichProjects. Usando dados basicos da pasta." -Level 'WARN'
    }

    return $Folder
}

function Select-SingleItem {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if ($Items.Count -eq 0) {
        throw "Nenhum item encontrado para selecao: $Title"
    }

    Write-Host ''
    Write-Host $Title
    Write-Host ('-' * $Title.Length)

    for ($index = 0; $index -lt $Items.Count; $index++) {
        $number = $index + 1
        $name = Get-DisplayName -Item $Items[$index]
        Write-Host "$number. $name"
    }

    do {
        $answer = Read-Host 'Digite o numero desejado'
        $selectedNumber = 0
        $isValid = [int]::TryParse($answer, [ref]$selectedNumber) -and $selectedNumber -ge 1 -and $selectedNumber -le $Items.Count

        if (-not $isValid) {
            Write-Host 'Selecao invalida. Tente novamente.'
        }
    } while (-not $isValid)

    return $Items[$selectedNumber - 1]
}

function Select-MultipleItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if ($Items.Count -eq 0) {
        throw "Nenhum item encontrado para selecao: $Title"
    }

    Write-Host ''
    Write-Host $Title
    Write-Host ('-' * $Title.Length)

    for ($index = 0; $index -lt $Items.Count; $index++) {
        $number = $index + 1
        $name = Get-DisplayName -Item $Items[$index]
        Write-Host "$number. $name"
    }

    Write-Host ''
    Write-Host 'Digite T para exportar todos, ou informe numeros separados por virgula. Exemplo: 1,3,5'

    do {
        $answer = (Read-Host 'Selecao').Trim()

        if ($answer -match '^(?i:t|todos|all)$') {
            return $Items
        }

        $selectedItems = New-Object System.Collections.Generic.List[object]
        $isValid = $true

        foreach ($part in ($answer -split ',')) {
            $selectedNumber = 0
            $trimmed = $part.Trim()

            if (-not [int]::TryParse($trimmed, [ref]$selectedNumber) -or $selectedNumber -lt 1 -or $selectedNumber -gt $Items.Count) {
                $isValid = $false
                break
            }

            $selectedItems.Add($Items[$selectedNumber - 1])
        }

        if ($isValid -and $selectedItems.Count -gt 0) {
            return @($selectedItems.ToArray())
        }

        Write-Host 'Selecao invalida. Tente novamente.'
    } while ($true)
}

if ($OnlyConnectedProjects -and $OnlyNonConnectedProjects) {
    throw 'Use apenas um filtro: -OnlyConnectedProjects ou -OnlyNonConnectedProjects.'
}

$loginWasCreated = $false

try {
    Write-Status -Message 'Iniciando exportacao do modelo de criacao de projetos.'
    Write-Log -Message "Arquivo de saida definido: $OutputPath"
    Write-Log -Message "Diagnostics habilitado: $IncludeDiagnostics"

    $isLoggedIn = $false
    try {
        $isLoggedIn = [bool](Get-PWLoginStatus)
    }
    catch {
        $isLoggedIn = $false
    }

    if (-not $isLoggedIn) {
        $loginParams = @{
            DoNotCreateWorkingDirectory = $true
        }

        if ($UseGui) {
            $loginParams.UseGui = $true
        }
        else {
            if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
                throw 'Informe -DatasourceName "Servidor:Datasource" ou use -UseGui para abrir a tela de login.'
            }

            $loginParams.DatasourceName = $DatasourceName

            if (-not [string]::IsNullOrWhiteSpace($UserName)) {
                $loginParams.UserName = $UserName

                if ($null -eq $Password) {
                    $Password = Read-Host -Prompt 'Senha do ProjectWise' -AsSecureString
                }

                $loginParams.Password = $Password
            }

            if ($BentleyIMS) {
                $loginParams.BentleyIMS = $true
            }
        }

        if ($NonAdminLogin) {
            $loginParams.NonAdminLogin = $true
        }

        $loginOk = New-PWLogin @loginParams
        if (-not $loginOk) {
            throw 'Nao foi possivel fazer login no ProjectWise.'
        }

        $loginWasCreated = $true
    }

    $selectedProjectFolders = $null

    if ($InteractiveSelection -or [string]::IsNullOrWhiteSpace($FolderPath)) {
        Write-Status -Message 'Buscando pasta de engenharia na raiz do datasource...'
        $rootFolders = Get-PWRootFoldersSafe
        $engineeringNames = @($ConcessionsRootPath) + @($EngineeringRootNames)
        $engineeringFolder = Find-FolderByPossibleNames -Folders $rootFolders -PossibleNames $engineeringNames -Description 'Pasta de engenharia'
        $engineeringFolderId = Get-FolderId -Folder $engineeringFolder

        if ([string]::IsNullOrWhiteSpace("$engineeringFolderId")) {
            throw 'Nao foi possivel obter o ID da pasta de engenharia.'
        }

        Write-Status -Message "Buscando concessoes em: $(Get-DisplayName -Item $engineeringFolder)"
        $concessions = @(Get-PWChildFoldersSafe -FolderId $engineeringFolderId | Sort-Object { Get-DisplayName -Item $_ })
        $selectedConcession = Select-SingleItem -Items $concessions -Title 'Selecione a concessao'
        $selectedConcessionId = Get-FolderId -Folder $selectedConcession

        if ([string]::IsNullOrWhiteSpace("$selectedConcessionId")) {
            throw 'Nao foi possivel obter o ID da concessao selecionada.'
        }

        Write-Status -Message "Concessao selecionada: $(Get-DisplayName -Item $selectedConcession)"
        $concessionChildren = Get-PWChildFoldersSafe -FolderId $selectedConcessionId
        $projectsFolder = Find-FolderByPossibleNames -Folders $concessionChildren -PossibleNames $ProjectsFolderNames -Description 'Pasta de projetos da concessao'
        $projectsFolderId = Get-FolderId -Folder $projectsFolder

        if ([string]::IsNullOrWhiteSpace("$projectsFolderId")) {
            throw 'Nao foi possivel obter o ID da pasta de projetos da concessao.'
        }

        Write-Status -Message "Buscando projetos em: $(Get-DisplayName -Item $projectsFolder)"
        $projectFolders = @(Get-PWChildFoldersSafe -FolderId $projectsFolderId | Sort-Object { Get-DisplayName -Item $_ })
        $selectedProjectFolders = @(Select-MultipleItems -Items $projectFolders -Title 'Selecione os projetos para exportar')
    }

    $searchParams = @{
        PopulatePaths = $true
        PopulateProjectProperties = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $searchParams.FolderPath = $FolderPath
    }

    if (-not [string]::IsNullOrWhiteSpace($ProjectTypeName)) {
        $searchParams.ProjectTypeName = $ProjectTypeName
    }

    if ($OnlyConnectedProjects) {
        $searchParams.OnlyConnectedProjects = $true
    }

    if ($OnlyNonConnectedProjects) {
        $searchParams.OnlyNonConnectedProjects = $true
    }

    if ($JustOne) {
        if ([string]::IsNullOrWhiteSpace($FolderPath)) {
            throw 'Para usar -JustOne, informe também -FolderPath com o caminho completo do projeto.'
        }

        $searchParams.JustOne = $true
    }

    if ($selectedProjectFolders) {
        Write-Status -Message 'Lendo propriedades dos projetos selecionados...'
        $projects = @($selectedProjectFolders | ForEach-Object { Get-RichProjectFromFolder -Folder $_ })
    }
    else {
        Write-Status -Message 'Buscando projetos no ProjectWise...'
        $projects = @(Get-PWRichProjects @searchParams)

        if ($InteractiveSelection) {
            $projects = @(Select-MultipleItems -Items $projects -Title 'Selecione os projetos para exportar')
        }
    }

    $rows = foreach ($project in $projects) {
        $properties = Get-ProjectPropertyMap -Project $project
        $projectPath = Get-ProjectPath -Project $project
        $nomeConcessao = Get-ProjectPropertyValue -PropertyMap $properties -Names @(
            'ProjectNomeConcessao',
            'NomeConcessao',
            'Nome Concessao',
            'Concessao',
            'Concessionaria',
            'Unidade',
            'Empresa',
            'PROJECT_Nome_Concessao',
            'PROJECT_Concessao'
        )
        if ([string]::IsNullOrWhiteSpace("$nomeConcessao")) {
            $nomeConcessao = Get-PathSegment -Path $projectPath -Index 1
        }

        $siglaConcessao = Get-ProjectPropertyValue -PropertyMap $properties -Names @(
            'ProjectNovaSiglaANTT',
            'ProjectSiglaDaConcesso',
            'Nova Sigla',
            'Sigla Concessao',
            'SiglaConcessao',
            'Sigla',
            'CodConcessao',
            'CodigoConcessao',
            'PROJECT_Sigla_Concessao',
            'PROJECT_Sigla'
        )
        if ([string]::IsNullOrWhiteSpace("$siglaConcessao")) {
            $siglaConcessao = Get-ConcessionAcronym -ConcessionName $nomeConcessao
        }

        $row = [ordered]@{
            'Nome Concessao' = $nomeConcessao
            'Sigla Concessao' = $siglaConcessao
            'Nome Projeto' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('Nome Projeto', 'NomeProjeto', 'ProjectName', 'Name', 'FolderName', 'PROJECT_Project_Name', 'PROJECT_Nome_Projeto')
            'Descricao' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('Descricao', 'Description', 'Descripition', 'ProjectDescription', 'Desc', 'PROJECT_Descricao')
            'Projetista' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('ProjectNomeProjetista', 'Nome Projetista', 'Projetista', 'Designer', 'Empresa Projetista', 'EmpresaProjetista', 'PROJECT_Projetista')
            'Gestor Engenharia' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('ProjectGestorDaEngenharia', 'Gestor da Engenharia', 'Gestor Engenharia', 'GestorEngenharia', 'Gestor', 'Coordenador Engenharia', 'PROJECT_Gestor_Engenharia')
            'Assistente Engenharia' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('ProjectAssistente', 'Assistente', 'Assistente Engenharia', 'AssistenteEngenharia', 'PROJECT_Assistente_Engenharia')
            'Poder Concedente' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('ProjectPoderConcedente', 'P. Concedente', 'Poder Concedente', 'PoderConcedente', 'Agencia Reguladora', 'Concedente', 'PROJECT_Poder_Concedente')
            'TH' = Get-ProjectPropertyValue -PropertyMap $properties -Names @('ProjectTH', 'TH', 'T H', 'Termo Homologacao', 'TermoHomologacao', 'PROJECT_TH')
        }

        if ([string]::IsNullOrWhiteSpace("$($row['Nome Projeto'])")) {
            $row['Nome Projeto'] = Get-ObjectPropertyValue -InputObject $project -Names @('Name', 'ProjectName', 'FolderName')
        }

        if ([string]::IsNullOrWhiteSpace("$($row['Descricao'])")) {
            $row['Descricao'] = Get-ObjectPropertyValue -InputObject $project -Names @('Description', 'Descripition', 'ProjectDescription')
        }

        if ($IncludeDiagnostics) {
            $row['PW_Caminho'] = $projectPath
            $row['PW_IDProjeto'] = Get-ObjectPropertyValue -InputObject $project -Names @('ProjectID', 'ProjectId', 'ID', 'FolderID', 'FolderId')
            $row['PW_GUID'] = Get-ObjectPropertyValue -InputObject $project -Names @('GUID', 'Guid', 'ProjectGUID', 'ProjectGuid')
            $row['PW_TipoProjeto'] = Get-ObjectPropertyValue -InputObject $project -Names @('ProjectTypeName', 'ProjectType', 'RichProjectTypeName')
            $row['PW_PropriedadesEncontradas'] = @($properties.Keys | Sort-Object) -join ', '
        }

        [pscustomobject]$row
    }

    $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent

    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }

    $rows | Export-Csv -Path $resolvedOutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    Write-Status -Message "Exportacao concluida: $resolvedOutputPath"
    Write-Status -Message "Projetos exportados: $($rows.Count)"
    Write-Status -Message "Log gerado: $LogPath"
}
catch {
    Write-Log -Message $_.Exception.Message -Level 'ERROR'
    Write-Host ''
    Write-Host 'Ocorreu um erro durante a exportacao.' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    throw
}
finally {
    if ($loginWasCreated -and -not $KeepSessionOpen) {
        Undo-PWLogin
        Write-Log -Message 'Sessao ProjectWise encerrada.'
    }

    if (-not $NaoPausar) {
        Write-Host ''
        Read-Host 'Pressione Enter para encerrar'
    }
}
