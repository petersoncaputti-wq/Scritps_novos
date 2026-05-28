<#
.SYNOPSIS
    Reanalisa e movimenta documentos ProjectWise que ficaram sem validacao ou sem pasta destino.

.DESCRIPTION
    O script trabalha somente dentro da pasta escolhida pelo usuario e suas subpastas.
    Em modo -Simular ou -WhatIf, nao cria pastas, nao copia/move documentos e nao altera states.

.EXAMPLE
    .\ReavaliarMovimentacaoDocumentos.ps1 -Simular

.EXAMPLE
    .\ReavaliarMovimentacaoDocumentos.ps1 -PastaAnalise 'ENGENHARIA\Concessao\Projetos\Projeto X' -Simular
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$PastaAnalise,
    [switch]$Simular,
    [switch]$MoverDocumentoOrigem,
    [switch]$SomentePastaSelecionada,
    [switch]$AutoCriarPastasDestino,
    [int]$LimiteBuscaInicial = 0,
    [int]$LimiteDocumentos = -1,
    [string]$LogPath = ''
)

$ErrorActionPreference = 'Stop'

$ScriptDirectory = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Split-Path -Path $PSCommandPath -Parent
}
else {
    (Get-Location).Path
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $ScriptDirectory ("Logs\ReavaliarMovimentacaoDocumentos_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

#-------------------------------------------------------
# Configuracoes gerais
#-------------------------------------------------------
# Deixe vazio para usar o mesmo comportamento do Ajuste_de_Atributos.ps1.
$DatasourceName = ''
$UsarBentleyIMS = $true

$StatesParaReanalise = @(
    'Nao Validado pelo Sistema',
    'Validado pelo Sistema - Pasta destino nao localizada'
)

$StateOrigemCopiadoComSucesso = 'Validado pelo Sistema - Copiado p/ Disciplina'
$StateDestinoDocumentoNovo    = 'Em analise do Assistente'
$StateSuperado                = 'Superado'
$WorkflowEngenharia           = 'Workflow - Engenharia'
$WorkflowComentarios          = 'Workflow - Comentarios'
$EnvironmentXFDF              = 'dmsXFDF'

$NomePastaAreaTrabalho = '1 - Area de Trabalho'
$NomePastaSuperados    = 'Superados'
$NomePastaXFDF         = 'xfdf__'

$PastaAdministracao = "Area de Administra$([char]0x00E7)$([char]0x00E3)o"
$CaminhoDisciplinasANTT = "Ecorodovias\$PastaAdministracao\Registros\Disciplinas ANTT"

$AtributoErros = 'Erros'

$NomesPossiveisEngenhariaRaiz = @('Engenharia', 'Engineering')
$NomesPossiveisPastaProjetos  = @('Projetos', 'Projeto', 'Projects', 'Project')
$ProfundidadeProjetos = 0

#-------------------------------------------------------
# Caches
#-------------------------------------------------------
$script:CachePastas = @{}
$script:PastasDestinoCriadas = @{}
$script:CacheDocumentosPasta = @{}
$script:CacheDisciplinasANTT = @{}
$script:CacheCmdlets = @{}
$script:LoginCriadoPeloScript = $false

$script:Contador = [ordered]@{
    Encontrados               = 0
    Processados               = 0
    DestinoCalculado          = 0
    PastaBaseNaoLocalizada    = 0
    DisciplinaNaoLocalizada   = 0
    TipoDocumentoVazio        = 0
    IgnoradosErroValidacao    = 0
    PastasCriadas             = 0
    DocumentosCopiados        = 0
    DocumentosMovidos         = 0
    JaExistiamNoDestino       = 0
    MovidosParaSuperados      = 0
    EstadosAtualizados        = 0
    Erros                     = 0
}

#-------------------------------------------------------
# Funcoes utilitarias
#-------------------------------------------------------
function Escrever-Log {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Mensagem,
        [ValidateSet('INFO', 'AVISO', 'ERRO', 'SIMULACAO', 'OK')][string]$Nivel = 'INFO'
    )

    $linha = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Nivel, $Mensagem
    Write-Host $linha

    $pastaLog = Split-Path -Path $LogPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($pastaLog) -and -not (Test-Path -LiteralPath $pastaLog)) {
        New-Item -ItemType Directory -Path $pastaLog -Force | Out-Null
    }

    Add-Content -Path $LogPath -Value $linha -Encoding UTF8
}

function Get-ValorAtributo {
    param($Documento, [Parameter(Mandatory = $true)][string]$Nome)

    if ($null -eq $Documento -or $null -eq $Documento.Attributes -or $null -eq $Documento.Attributes[0]) {
        return $null
    }

    return $Documento.Attributes[0].$Nome
}

function Obter-NomeDocumento {
    param($Documento)

    if ($Documento.FileName) { return $Documento.FileName }
    if ($Documento.Name)     { return $Documento.Name }
    return '<sem nome>'
}

function Normalizar-Texto {
    param([object]$Valor)

    if ($null -eq $Valor) { return '' }
    return ([string]$Valor).Trim()
}

function Testar-TextoVazio {
    param([object]$Valor)
    return [string]::IsNullOrWhiteSpace((Normalizar-Texto $Valor))
}

function Get-CmdletCached {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not $script:CacheCmdlets.ContainsKey($Name)) {
        $script:CacheCmdlets[$Name] = Get-Command -Name $Name -ErrorAction SilentlyContinue
    }

    return $script:CacheCmdlets[$Name]
}

function Converter-PWUriParaFolderPath {
    param([Parameter(Mandatory = $true)][string]$Caminho)

    $valor = $Caminho.Trim().TrimEnd('\')
    if ($valor -notmatch '^pw:\\\\') { return $valor }

    $semPrefixo = $valor -replace '^pw:\\\\[^\\]+\\Documents\\', ''
    return $semPrefixo.TrimEnd('\')
}

function Connect-ProjectWiseDatasource {
    try {
        $currentDatasource = Get-PWCurrentDatasource -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$currentDatasource)) {
            Escrever-Log -Nivel 'OK' -Mensagem "Sessao ProjectWise ativa: $currentDatasource"
            return
        }
    }
    catch {
        Escrever-Log -Nivel 'AVISO' -Mensagem 'Nenhuma sessao ProjectWise ativa encontrada. Sera realizado login.'
    }

    if ($UsarBentleyIMS) {
        if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
            Escrever-Log -Mensagem 'Efetuando login via Bentley IMS.'
            New-PWLogin -BentleyIMS -ErrorAction Stop | Out-Null
            $script:LoginCriadoPeloScript = $true
        }
        else {
            Escrever-Log -Mensagem "Efetuando login via Bentley IMS no datasource '$DatasourceName'."
            New-PWLogin -DatasourceName $DatasourceName -BentleyIMS -ErrorAction Stop | Out-Null
            $script:LoginCriadoPeloScript = $true
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
            Escrever-Log -Mensagem 'Efetuando login pelo seletor padrao do ProjectWise.'
            New-PWLogin -ErrorAction Stop | Out-Null
            $script:LoginCriadoPeloScript = $true
        }
        else {
            Escrever-Log -Mensagem "Efetuando login no datasource '$DatasourceName'."
            New-PWLogin -DatasourceName $DatasourceName -UseGui -ErrorAction Stop | Out-Null
            $script:LoginCriadoPeloScript = $true
        }
    }

    $currentDatasource = Get-PWCurrentDatasource -ErrorAction Stop
    Escrever-Log -Nivel 'OK' -Mensagem "Login realizado com sucesso: $currentDatasource"
}

function Get-FolderPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Folder,
        [Parameter(Mandatory = $true)][string[]]$PropertyNames
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

    return ''
}

function Get-FolderId {
    param([Parameter(Mandatory = $true)]$Folder)
    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @('ProjectID', 'ProjectId', 'FolderID', 'FolderId', 'Id', 'ID')
}

function Get-FolderPathValue {
    param([Parameter(Mandatory = $true)]$Folder)
    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @('FullPath', 'Fullpath', 'FolderPath', 'Path', 'ProjectPath')
}

function Resolve-FolderPathById {
    param([Parameter(Mandatory = $true)][string]$FolderId)

    if ([string]::IsNullOrWhiteSpace($FolderId)) { return '' }

    try {
        $folder = Get-PWFolders -FolderID $FolderId -PopulatePaths -JustOne -ErrorAction Stop
        $path = Get-FolderPathValue -Folder $folder
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            return $path.TrimEnd('\')
        }
    }
    catch {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Nao foi possivel resolver caminho por Get-PWFolders para FolderID $FolderId. Erro: $($_.Exception.Message)"
    }

    try {
        $idNumerico = [int]$FolderId
        $sql = "SELECT REPLACE(dbo.GetVaultPath($idNumerico, 0), '/', '\') AS FullPath"
        $resultado = Select-PWSQL $sql
        if ($resultado -and $resultado.Rows.Count -gt 0) {
            $pathSql = ([string]$resultado.Rows[0].FullPath).TrimEnd('\')
            if (-not [string]::IsNullOrWhiteSpace($pathSql)) {
                return $pathSql
            }
        }
    }
    catch {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Nao foi possivel resolver caminho por SQL para FolderID $FolderId. Erro: $($_.Exception.Message)"
    }

    return ''
}

function Get-FolderNameValue {
    param([Parameter(Mandatory = $true)]$Folder)
    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @('Name', 'FolderName', 'ProjectName', 'ObjectName')
}

function Get-FolderDescriptionValue {
    param([Parameter(Mandatory = $true)]$Folder)
    return Get-FolderPropertyValue -Folder $Folder -PropertyNames @('Description', 'Descricao', 'ProjectDescription', 'FolderDescription')
}

function Get-FolderLabel {
    param([Parameter(Mandatory = $true)]$Folder)

    $description = Get-FolderDescriptionValue -Folder $Folder
    if (-not [string]::IsNullOrWhiteSpace($description)) {
        return $description
    }

    return Get-FolderNameValue -Folder $Folder
}

function Sort-FoldersByLabel {
    param([object[]]$Folders)
    return @($Folders | Sort-Object -Property @{ Expression = { (Get-FolderLabel -Folder $_).ToLowerInvariant() } })
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)

    if ($params -contains 'Root') {
        return @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $null -ne $_ })
    }

    return @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $null -ne $_ })
}

function Get-PWChildFoldersSafe {
    param(
        [Parameter(Mandatory = $true)][string]$FolderId,
        [string]$Context = ''
    )

    try {
        if (-not [string]::IsNullOrWhiteSpace($Context)) {
            Escrever-Log -Mensagem "Listando subpastas de: $Context"
        }

        return @(
            Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop |
                Where-Object {
                    $null -ne $_ -and
                    -not [string]::IsNullOrWhiteSpace((Get-FolderId -Folder $_)) -and
                    -not [string]::IsNullOrWhiteSpace((Get-FolderLabel -Folder $_))
                }
        )
    }
    catch {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Falha ao listar subpastas de '$Context'. Erro: $($_.Exception.Message)"
        return @()
    }
}

function Find-FolderByPossibleNames {
    param(
        [Parameter(Mandatory = $true)][object[]]$Folders,
        [Parameter(Mandatory = $true)][string[]]$PossibleNames,
        [Parameter(Mandatory = $true)][string]$Description
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

    $available = @($Folders | ForEach-Object { Get-FolderLabel -Folder $_ }) -join ', '
    throw "Nao encontrei $Description. Pastas disponiveis: $available"
}

function Read-NumberSelection {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][int]$Minimum,
        [Parameter(Mandatory = $true)][int]$Maximum
    )

    while ($true) {
        $inputValue = (Read-Host $Message).Trim()

        if ($inputValue -match '^\d+$') {
            $number = [int]$inputValue
            if ($number -ge $Minimum -and $number -le $Maximum) {
                return $number
            }
        }

        Escrever-Log -Nivel 'AVISO' -Mensagem "Informe um numero entre $Minimum e $Maximum."
    }
}

function Select-FolderFromList {
    param(
        [Parameter(Mandatory = $true)][object[]]$Folders,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Prompt
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan

    for ($i = 0; $i -lt $Folders.Count; $i++) {
        '{0:000}. {1}' -f ($i + 1), (Get-FolderLabel -Folder $Folders[$i]) | Write-Host
    }

    $selection = Read-NumberSelection -Message $Prompt -Minimum 1 -Maximum $Folders.Count
    return $Folders[$selection - 1]
}

function Get-ProjectsFromProjectsFolder {
    param(
        [Parameter(Mandatory = $true)]$ProjectsFolder,
        [int]$MaximumDepth = 0
    )

    $projectsFolderId = Get-FolderId -Folder $ProjectsFolder
    if ([string]::IsNullOrWhiteSpace($projectsFolderId)) {
        throw 'Nao foi possivel identificar o ID da pasta de projetos.'
    }

    $result = @()
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue([pscustomobject]@{
        FolderId = $projectsFolderId
        Level    = 0
        Label    = Get-FolderLabel -Folder $ProjectsFolder
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
                $queue.Enqueue([pscustomobject]@{
                    FolderId = $childId
                    Level    = $current.Level + 1
                    Label    = Get-FolderLabel -Folder $child
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
    Escrever-Log -Mensagem 'Carregando pastas raiz do ProjectWise.'
    $rootFolders = Get-PWRootFoldersSafe
    $engineeringFolder = Find-FolderByPossibleNames -Folders $rootFolders -PossibleNames $NomesPossiveisEngenhariaRaiz -Description 'a pasta raiz de engenharia'

    $engineeringFolderId = Get-FolderId -Folder $engineeringFolder
    if ([string]::IsNullOrWhiteSpace($engineeringFolderId)) {
        throw 'Nao foi possivel identificar o ID da pasta de engenharia.'
    }

    $concessions = Sort-FoldersByLabel -Folders (Get-PWChildFoldersSafe -FolderId $engineeringFolderId -Context 'Engenharia')
    if ($concessions.Count -eq 0) {
        throw 'Nenhuma concessao encontrada dentro da pasta Engenharia.'
    }

    $selectedConcession = Select-FolderFromList -Folders $concessions -Title 'Concessoes encontradas:' -Prompt 'Selecione a concessao'
    $selectedConcessionLabel = Get-FolderLabel -Folder $selectedConcession
    Escrever-Log -Nivel 'OK' -Mensagem "Concessao selecionada: $selectedConcessionLabel"

    $selectedConcessionId = Get-FolderId -Folder $selectedConcession
    if ([string]::IsNullOrWhiteSpace($selectedConcessionId)) {
        throw 'Nao foi possivel identificar o ID da concessao selecionada.'
    }

    $concessionChildren = Get-PWChildFoldersSafe -FolderId $selectedConcessionId -Context $selectedConcessionLabel
    $projectsFolder = Find-FolderByPossibleNames -Folders $concessionChildren -PossibleNames $NomesPossiveisPastaProjetos -Description "a pasta de projetos da concessao '$selectedConcessionLabel'"

    $projects = Get-ProjectsFromProjectsFolder -ProjectsFolder $projectsFolder -MaximumDepth $ProfundidadeProjetos
    if ($projects.Count -eq 0) {
        throw "Nenhum projeto encontrado na concessao '$selectedConcessionLabel'."
    }

    $selectedProject = Select-FolderFromList -Folders $projects -Title "Projetos encontrados em '$selectedConcessionLabel':" -Prompt 'Selecione o projeto'
    Escrever-Log -Nivel 'OK' -Mensagem "Projeto selecionado: $(Get-FolderLabel -Folder $selectedProject)"

    return $selectedProject
}

function Select-DocumentFolderFromProject {
    param([Parameter(Mandatory = $true)]$ProjectFolder)

    $currentFolder = $ProjectFolder
    $currentTrail = @(Get-FolderLabel -Folder $ProjectFolder)

    while ($true) {
        $currentFolderId = Get-FolderId -Folder $currentFolder
        if ([string]::IsNullOrWhiteSpace($currentFolderId)) {
            throw 'Nao foi possivel identificar o ID da pasta atual.'
        }

        $children = Sort-FoldersByLabel -Folders (Get-PWChildFoldersSafe -FolderId $currentFolderId -Context ($currentTrail -join '\'))

        Write-Host ''
        Write-Host "Pasta atual: $($currentTrail -join '\')" -ForegroundColor Cyan
        Write-Host '000. Usar esta pasta para buscar documentos'

        for ($i = 0; $i -lt $children.Count; $i++) {
            '{0:000}. {1}' -f ($i + 1), (Get-FolderLabel -Folder $children[$i]) | Write-Host
        }

        if ($children.Count -eq 0) {
            Escrever-Log -Nivel 'AVISO' -Mensagem 'A pasta atual nao possui subpastas. Ela sera usada para buscar documentos.'
            return $currentFolder
        }

        $selection = Read-NumberSelection -Message 'Selecione uma subpasta ou 0 para usar a pasta atual' -Minimum 0 -Maximum $children.Count
        if ($selection -eq 0) {
            return $currentFolder
        }

        $currentFolder = $children[$selection - 1]
        $currentTrail += (Get-FolderLabel -Folder $currentFolder)
    }
}

function Get-PWFoldersCached {
    param([Parameter(Mandatory = $true)][string]$FolderPath)

    $path = $FolderPath.TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }

    $key = $path.ToLowerInvariant()
    if (-not $script:CachePastas.ContainsKey($key)) {
        try {
            $script:CachePastas[$key] = Get-PWFolders -FolderPath $path -PopulatePaths -JustOne -ErrorAction Stop
        }
        catch {
            $script:CachePastas[$key] = $null
        }
    }

    return $script:CachePastas[$key]
}

function Registrar-Erro {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$Mensagem
    )

    $nome = Obter-NomeDocumento -Documento $Documento
    Escrever-Log -Nivel 'ERRO' -Mensagem ("{0}: {1}" -f $nome, $Mensagem)

    if ($Simular) {
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem ("Atualizaria atributo {0}: {1}" -f $AtributoErros, $Mensagem)
        return
    }

    try {
        $null = Update-PWDocumentAttributes -InputDocuments $Documento -Attributes @{ $AtributoErros = $Mensagem } -WarningAction SilentlyContinue
    }
    catch {
        Escrever-Log -Nivel 'ERRO' -Mensagem ("Falha ao registrar erro no atributo {0}: {1}" -f $AtributoErros, $_.Exception.Message)
    }
}

function Testar-ErroValidacaoBloqueante {
    param(
        [string]$StateAtual,
        [string]$MensagemErro
    )

    if ($StateAtual -ne 'Nao Validado pelo Sistema') {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($MensagemErro)) {
        return $false
    }

    $erro = $MensagemErro.ToLowerInvariant()

    $padroesBloqueantes = @(
        'revisao',
        'versao',
        'taxonomia',
        'nao foram emitidos todos os arquivos',
        'documento nao previsto',
        'documento não previsto'
    )

    foreach ($padrao in $padroesBloqueantes) {
        if ($erro.Contains($padrao)) {
            return $true
        }
    }

    return $false
}

function Selecionar-PastaAnalise {
    param([string]$PastaInformada)

    if (-not [string]::IsNullOrWhiteSpace($PastaInformada)) {
        return (Converter-PWUriParaFolderPath -Caminho $PastaInformada)
    }

    $selectedProject = Select-ProjectFromConsole
    $selectedFolder = Select-DocumentFolderFromProject -ProjectFolder $selectedProject
    $selectedFolderPath = Get-FolderPathValue -Folder $selectedFolder

    if ([string]::IsNullOrWhiteSpace($selectedFolderPath)) {
        $folderId = Get-FolderId -Folder $selectedFolder
        if (-not [string]::IsNullOrWhiteSpace($folderId)) {
            $selectedFolderPath = Resolve-FolderPathById -FolderId $folderId
        }
    }

    if ([string]::IsNullOrWhiteSpace($selectedFolderPath)) {
        throw 'Nao foi possivel identificar o caminho da pasta selecionada.'
    }

    Escrever-Log -Nivel 'OK' -Mensagem "Pasta selecionada para reanalise: $selectedFolderPath"
    return $selectedFolderPath.TrimEnd('\')
}

function Obter-DocumentosParaReanalise {
    param([Parameter(Mandatory = $true)][string]$PastaRaiz)

    if ($LimiteBuscaInicial -gt 0 -and -not $SomentePastaSelecionada) {
        Escrever-Log -Nivel 'AVISO' -Mensagem 'LimiteBuscaInicial so limita a busca real quando usado junto com -SomentePastaSelecionada. A busca em subpastas seguira pelo metodo normal.'
    }

    if ($LimiteBuscaInicial -gt 0 -and $SomentePastaSelecionada) {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Busca inicial limitada aos primeiros $LimiteBuscaInicial documentos da pasta selecionada."

        $pasta = Get-PWFoldersCached -FolderPath $PastaRaiz
        if (-not $pasta -or [string]::IsNullOrWhiteSpace([string]$pasta.ProjectID)) {
            throw "Nao foi possivel identificar o ProjectID da pasta para aplicar LimiteBuscaInicial: $PastaRaiz"
        }

        $top = [math]::Max(1, $LimiteBuscaInicial)
        $projectId = [int]$pasta.ProjectID
        $sql = @"
SELECT TOP ($top)
    D.o_itemname AS DocumentName
FROM dms_doc AS D
WHERE D.o_projectno = $projectId
ORDER BY D.o_itemname
"@

        $linhas = Select-PWSQL -SQLSelectStatement $sql
        $nomes = @()
        if ($linhas -and $linhas.Rows) {
            $nomes = @($linhas.Rows | ForEach-Object { ([string]$_.DocumentName).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }

        Escrever-Log -Mensagem ("Documentos candidatos carregados pela busca limitada: {0}" -f $nomes.Count)

        $documentosLimitados = New-Object System.Collections.Generic.List[object]
        foreach ($nomeDocumento in $nomes) {
            try {
                $doc = Get-PWDocumentsBySearch `
                    -FolderPath $PastaRaiz `
                    -JustThisFolder `
                    -DocumentName $nomeDocumento `
                    -GetAttributes `
                    -WarningAction SilentlyContinue

                foreach ($item in @($doc)) {
                    if ($item) { $documentosLimitados.Add($item) }
                }
            }
            catch {
                Escrever-Log -Nivel 'AVISO' -Mensagem "Falha ao carregar documento '$nomeDocumento' pela busca limitada: $($_.Exception.Message)"
            }
        }

        return @($documentosLimitados | Where-Object {
            $StatesParaReanalise -contains $_.WorkflowState
        })
    }

    if ($SomentePastaSelecionada) {
        Escrever-Log -Mensagem "Buscando documentos somente na pasta selecionada, sem subpastas: $PastaRaiz"

        $documentos = @(Get-PWDocumentsBySearch `
            -FolderPath $PastaRaiz `
            -JustThisFolder `
            -GetAttributes `
            -WarningAction SilentlyContinue)
    }
    else {
        Escrever-Log -Mensagem "Buscando documentos na pasta selecionada e subpastas: $PastaRaiz"

        $documentos = @(Get-PWDocumentsBySearch `
            -FolderPath $PastaRaiz `
            -GetAttributes `
            -WarningAction SilentlyContinue)
    }

    return @($documentos | Where-Object {
        $StatesParaReanalise -contains $_.WorkflowState
    })
}

function Confirmar-ProcessamentoDocumentos {
    param(
        [Parameter(Mandatory = $true)][object[]]$Documentos,
        [Parameter(Mandatory = $true)][string]$PastaRaiz
    )

    $total = $Documentos.Count
    Escrever-Log -Mensagem ("Documentos encontrados para reanalise: {0}" -f $total)

    if ($total -eq 0) {
        Escrever-Log -Nivel 'AVISO' -Mensagem 'Nenhum documento encontrado com os states configurados para reanalise.'
        return $false
    }

    Write-Host ''
    Write-Host 'Resumo antes do processamento' -ForegroundColor Cyan
    Write-Host "Pasta analisada: $PastaRaiz" -ForegroundColor Cyan
    Write-Host "Total para reanalise: $total" -ForegroundColor Cyan

    $porState = $Documentos |
        Group-Object -Property WorkflowState |
        Sort-Object Name

    foreach ($grupo in $porState) {
        Write-Host ('- {0}: {1}' -f $grupo.Name, $grupo.Count) -ForegroundColor Cyan
    }

    if ($Simular) {
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem 'Modo simulacao ativo. O processamento sera executado sem alterar documentos/pastas/states.'
        return $true
    }

    $confirmacao = (Read-Host 'Confirmar inicio da reanalise/movimentacao? Digite S para confirmar').Trim()
    if ($confirmacao -notin @('S', 's', 'SIM', 'Sim', 'sim')) {
        Escrever-Log -Nivel 'AVISO' -Mensagem 'Processamento cancelado pelo usuario antes da reanalise.'
        return $false
    }

    return $true
}

function Selecionar-LimiteDocumentos {
    param(
        [Parameter(Mandatory = $true)][int]$TotalDocumentos
    )

    if ($LimiteDocumentos -ge 0) {
        return $LimiteDocumentos
    }

    Write-Host ''
    Write-Host 'Informe quantos documentos deseja processar nesta execucao.' -ForegroundColor Cyan
    Write-Host 'Use 0 para processar todos.' -ForegroundColor Cyan

    while ($true) {
        $inputValue = (Read-Host 'Quantidade para processar').Trim()

        if ($inputValue -match '^\d+$') {
            $limite = [int]$inputValue
            if ($limite -eq 0 -or ($limite -ge 1 -and $limite -le $TotalDocumentos)) {
                return $limite
            }
        }

        Escrever-Log -Nivel 'AVISO' -Mensagem "Informe 0 ou um numero entre 1 e $TotalDocumentos."
    }
}

function Carregar-DisciplinasANTT {
    param([string]$CaminhoLista = $CaminhoDisciplinasANTT)

    if ($script:CacheDisciplinasANTT.Count -gt 0) { return $script:CacheDisciplinasANTT }

    $folderPath = Converter-PWUriParaFolderPath -Caminho $CaminhoLista
    Escrever-Log -Mensagem "Carregando lista oficial de disciplinas ANTT: $folderPath"

    $registros = @(Get-PWDocumentsBySearch -FolderPath $folderPath -GetAttributes -WarningAction SilentlyContinue)

    if ($registros.Count -eq 0) {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Nenhum registro encontrado na pasta de disciplinas ANTT. Tentando fallback pelo ambiente dmsRegistro."

        try {
            $registros = @(Get-PWDocumentsBySearch `
                -Environment 'dmsRegistro' `
                -Attributes @{ TipoRegistro = 'Disciplinas ANTT' } `
                -GetAttributes `
                -WarningAction SilentlyContinue)
        }
        catch {
            Escrever-Log -Nivel 'AVISO' -Mensagem "Fallback pelo ambiente dmsRegistro falhou: $($_.Exception.Message)"
            $registros = @()
        }
    }

    foreach ($registro in $registros) {
        $codigo = Normalizar-Texto (Get-ValorAtributo -Documento $registro -Nome 'Codigo')
        if ([string]::IsNullOrWhiteSpace($codigo)) {
            $codigo = Normalizar-Texto (Get-ValorAtributo -Documento $registro -Nome 'Disciplina')
        }

        if ([string]::IsNullOrWhiteSpace($codigo)) { continue }

        $nomeOficial = Normalizar-Texto (Get-ValorAtributo -Documento $registro -Nome 'Relacionamento')
        if ([string]::IsNullOrWhiteSpace($nomeOficial)) { $nomeOficial = Normalizar-Texto (Get-ValorAtributo -Documento $registro -Nome 'DisciplinaPai') }
        if ([string]::IsNullOrWhiteSpace($nomeOficial)) { $nomeOficial = Normalizar-Texto (Get-ValorAtributo -Documento $registro -Nome 'Nome') }
        if ([string]::IsNullOrWhiteSpace($nomeOficial)) { $nomeOficial = Normalizar-Texto $registro.Description }
        if ([string]::IsNullOrWhiteSpace($nomeOficial)) { $nomeOficial = Normalizar-Texto $registro.Name }

        if (-not [string]::IsNullOrWhiteSpace($nomeOficial)) {
            $script:CacheDisciplinasANTT[$codigo.ToUpperInvariant()] = $nomeOficial
        }
    }

    Escrever-Log -Mensagem ("Disciplinas ANTT carregadas: {0}" -f $script:CacheDisciplinasANTT.Count)
    return $script:CacheDisciplinasANTT
}

function Resolver-NomeDisciplina {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)]$MapaDisciplinas,
        [string]$PoderConcedente
    )

    $codigo = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'Disciplina')
    $disciplinaPai = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'DisciplinaPai')

    $resultado = [ordered]@{
        Codigo               = $codigo
        Nome                 = $null
        EncontradaLista      = $false
        UsouFallbackAtributo = $false
        Erro                 = $null
    }

    if ($PoderConcedente -eq 'ANTT' -and -not [string]::IsNullOrWhiteSpace($codigo)) {
        $key = $codigo.ToUpperInvariant()
        if ($MapaDisciplinas.ContainsKey($key)) {
            $resultado.Nome = $MapaDisciplinas[$key]
            $resultado.EncontradaLista = $true
            return [pscustomobject]$resultado
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($disciplinaPai)) {
        $resultado.Nome = $disciplinaPai
        $resultado.UsouFallbackAtributo = $true
        return [pscustomobject]$resultado
    }

    if ($PoderConcedente -eq 'ANTT') {
        $resultado.Erro = 'Disciplina nao localizada na lista oficial e atributo Disciplina Pai vazio.'
    }
    else {
        $resultado.Erro = 'Disciplina Pai vazio. Nao foi possivel montar pasta destino.'
    }

    return [pscustomobject]$resultado
}

function Resolver-FaseDestino {
    param([Parameter(Mandatory = $true)]$Documento)

    $fasePai = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'FasePai')
    if (-not [string]::IsNullOrWhiteSpace($fasePai)) { return $fasePai }

    $faseProjetoPai = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'FaseProjetoPai')
    if (-not [string]::IsNullOrWhiteSpace($faseProjetoPai)) { return $faseProjetoPai }

    $faseProjeto = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'FaseProjeto')
    $fase = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'Fase')

    switch ($faseProjeto) {
        'FUN' { return '01 - Funcional' }
        'ANT' { return '02 - Anteprojeto' }
        'EXE' { return '03 - Executivo' }
    }

    switch ($fase) {
        'FUN' { return '01 - Funcional' }
        'ANT' { return '02 - Anteprojeto' }
        'EXE' { return '03 - Executivo' }
    }

    return ''
}

function Resolver-TipoDocumentoDestino {
    param([Parameter(Mandatory = $true)]$Documento)

    $tipoPai = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'TipoDocumentoPai')
    if (-not [string]::IsNullOrWhiteSpace($tipoPai)) { return $tipoPai }

    $tipoDocumento = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'TipoDocumento')
    switch ($tipoDocumento) {
        'DE' { return 'Desenhos' }
        'MC' { return 'Memoria de Calculo' }
        'MD' { return 'Memorial Descritivo' }
        'RT' { return 'Relatorios Tecnicos' }
        default { return '' }
    }
}

function Obter-PastaRaizProjeto {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$PastaAnaliseSelecionada
    )

    try {
        $richProject = Get-PWRichProjectForDocument -InputDocument $Documento -ErrorAction Stop
        if ($richProject.FullPath) { return $richProject.FullPath.TrimEnd('\') }
    }
    catch {
        Escrever-Log -Nivel 'AVISO' -Mensagem ("Nao foi possivel obter rich project para {0}: {1}" -f (Obter-NomeDocumento $Documento), $_.Exception.Message)
    }

    $marcador = "\$NomePastaAreaTrabalho\"
    $folder = Normalizar-Texto $Documento.FolderPath
    if ($folder -and $folder.Contains($marcador)) {
        return $folder.Substring(0, $folder.IndexOf($marcador)).TrimEnd('\')
    }

    if ($PastaAnaliseSelecionada.Contains($marcador)) {
        return $PastaAnaliseSelecionada.Substring(0, $PastaAnaliseSelecionada.IndexOf($marcador)).TrimEnd('\')
    }

    return $PastaAnaliseSelecionada.TrimEnd('\')
}

function Montar-CaminhoDestino {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$PastaAnaliseSelecionada,
        [Parameter(Mandatory = $true)]$MapaDisciplinas
    )

    $poderConcedente = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'PoderConcedente')
    $faseDestino = Resolver-FaseDestino -Documento $Documento
    $volume = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'Volume')
    $disciplina = Resolver-NomeDisciplina -Documento $Documento -MapaDisciplinas $MapaDisciplinas -PoderConcedente $poderConcedente
    $tipoDestino = Resolver-TipoDocumentoDestino -Documento $Documento
    $pastaRaizProjeto = Obter-PastaRaizProjeto -Documento $Documento -PastaAnaliseSelecionada $PastaAnaliseSelecionada

    $resultado = [ordered]@{
        PoderConcedente       = $poderConcedente
        FaseDestino           = $faseDestino
        Volume                = $volume
        DisciplinaCodigo      = $disciplina.Codigo
        DisciplinaNome        = $disciplina.Nome
        DisciplinaLista       = $disciplina.EncontradaLista
        DisciplinaFallback    = $disciplina.UsouFallbackAtributo
        TipoDocumentoCodigo   = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'TipoDocumento')
        TipoDocumentoDestino  = $tipoDestino
        PastaRaizProjeto      = $pastaRaizProjeto
        PastaBase             = ''
        PastaDestino          = ''
        PastaBaseLocalizada   = $false
        Erros                 = New-Object System.Collections.Generic.List[string]
    }

    if ([string]::IsNullOrWhiteSpace($faseDestino)) { $resultado.Erros.Add('Fase Pai vazio. Nao foi possivel montar pasta destino.') }
    if ([string]::IsNullOrWhiteSpace($volume)) { $resultado.Erros.Add('Volume vazio. Nao foi possivel montar pasta destino.') }
    if ($disciplina.Erro) { $resultado.Erros.Add($disciplina.Erro) }
    if ([string]::IsNullOrWhiteSpace($tipoDestino)) { $resultado.Erros.Add('Tipo Documento Pai vazio. Nao foi possivel montar pasta destino.') }

    if ($resultado.Erros.Count -gt 0) {
        return [pscustomobject]$resultado
    }

    $resultado.PastaBase = @(
        $pastaRaizProjeto,
        $NomePastaAreaTrabalho,
        $faseDestino,
        $volume
    ) -join '\'

    $resultado.PastaDestino = @(
        $resultado.PastaBase,
        $disciplina.Nome,
        $tipoDestino
    ) -join '\'

    $resultado.PastaBaseLocalizada = [bool](Get-PWFoldersCached -FolderPath $resultado.PastaBase)
    if (-not $resultado.PastaBaseLocalizada) {
        $resultado.Erros.Add("Pasta base nao localizada: $($resultado.PastaBase)")
    }

    return [pscustomobject]$resultado
}

function Garantir-PastaDestino {
    param(
        [Parameter(Mandatory = $true)]$DestinoInfo
    )

    if (-not $DestinoInfo.PastaBaseLocalizada) {
        return $false
    }

    $pastaDisciplina = Join-Path $DestinoInfo.PastaBase $DestinoInfo.DisciplinaNome
    $pastaTipo = Join-Path $pastaDisciplina $DestinoInfo.TipoDocumentoDestino

    $keyDestino = $pastaTipo.TrimEnd('\').ToLowerInvariant()
    if ($script:PastasDestinoCriadas.ContainsKey($keyDestino)) {
        Escrever-Log -Mensagem "Pasta destino ja garantida em cache: $pastaTipo"
        return $true
    }

    $pastaDisciplinaExiste = [bool](Get-PWFoldersCached -FolderPath $pastaDisciplina)
    $pastaTipoExiste = [bool](Get-PWFoldersCached -FolderPath $pastaTipo)

    Escrever-Log -Mensagem ("Pasta disciplina existe: {0}" -f $(if ($pastaDisciplinaExiste) { 'Sim' } else { 'Nao' }))
    Escrever-Log -Mensagem ("Pasta tipo documento existe: {0}" -f $(if ($pastaTipoExiste) { 'Sim' } else { 'Nao' }))

    if ($pastaDisciplinaExiste -and $pastaTipoExiste) {
        $script:PastasDestinoCriadas[$keyDestino] = $true
        return $true
    }

    if ($Simular) {
        if (-not $pastaDisciplinaExiste) { Escrever-Log -Nivel 'SIMULACAO' -Mensagem "Criaria pasta $($DestinoInfo.DisciplinaNome)" }
        if (-not $pastaTipoExiste) { Escrever-Log -Nivel 'SIMULACAO' -Mensagem "Criaria pasta $($DestinoInfo.TipoDocumentoDestino)" }
        $script:PastasDestinoCriadas[$keyDestino] = $true
        return $true
    }

    Escrever-Log -Nivel 'AVISO' -Mensagem "Criacao automatica autorizada para pasta destino: $pastaTipo"

    try {
        if (-not $pastaDisciplinaExiste) {
            $null = New-PWFolder -FolderPath $pastaDisciplina -Workflow $WorkflowEngenharia -ErrorAction Stop
            $script:Contador.PastasCriadas++
            $script:CachePastas[$pastaDisciplina.ToLowerInvariant()] = Get-PWFolders -FolderPath $pastaDisciplina -PopulatePaths -JustOne
        }

        if (-not $pastaTipoExiste) {
            $null = New-PWFolder -FolderPath $pastaTipo -Workflow $WorkflowEngenharia -ErrorAction Stop
            $script:Contador.PastasCriadas++
            $script:CachePastas[$pastaTipo.ToLowerInvariant()] = Get-PWFolders -FolderPath $pastaTipo -PopulatePaths -JustOne
        }

        $script:PastasDestinoCriadas[$keyDestino] = $true
        return $true
    }
    catch {
        throw "Falha ao garantir pasta destino ${pastaTipo}: $($_.Exception.Message)"
    }
}

function Obter-DocumentosRelacionados {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$PastaAnaliseSelecionada
    )

    $numeroDocumento = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente')
    $sufixo = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'Sufixo')
    $pastaOrigem = Normalizar-Texto $Documento.FolderPath

    $relacionados = New-Object System.Collections.Generic.List[object]

    if (-not [string]::IsNullOrWhiteSpace($numeroDocumento)) {
        $candidatos = @(Get-PWDocumentsBySearch -FolderPath $PastaAnaliseSelecionada -Attributes @{ NumeroPoderConcedente = $numeroDocumento } -GetAttributes -WarningAction SilentlyContinue)
        foreach ($candidato in $candidatos) {
            $sufixoCand = Normalizar-Texto (Get-ValorAtributo -Documento $candidato -Nome 'Sufixo')
            if ([string]::IsNullOrWhiteSpace($sufixo) -or $sufixoCand -eq $sufixo -or [string]::IsNullOrWhiteSpace($sufixoCand)) {
                $relacionados.Add($candidato)
            }
        }
    }

    if ($pastaOrigem) {
        $nomeBase = [System.IO.Path]::GetFileNameWithoutExtension((Obter-NomeDocumento -Documento $Documento))
        foreach ($ext in @('.pdf', '.xfdf')) {
            try {
                $porNome = Get-PWDocumentsBySearch -FolderPath $pastaOrigem -JustThisFolder -FileName ($nomeBase + $ext) -GetAttributes -WarningAction SilentlyContinue
                foreach ($item in @($porNome)) { $relacionados.Add($item) }
            }
            catch {
                # Parametro FileName pode variar por versao; a busca por atributo acima ja cobre o fluxo principal.
            }
        }

        $pastaXFDF = Join-Path $pastaOrigem $NomePastaXFDF
        if (Get-PWFoldersCached -FolderPath $pastaXFDF) {
            try {
                $xfdfs = @(Get-PWDocumentsBySearch -FolderPath $pastaXFDF -GetAttributes -WarningAction SilentlyContinue)
                foreach ($xfdf in $xfdfs) {
                    if ($xfdf.FileName -like "$nomeBase*.xfdf" -or $xfdf.Name -like "$nomeBase*.xfdf") {
                        $relacionados.Add($xfdf)
                    }
                }
            }
            catch {
                Escrever-Log -Nivel 'AVISO' -Mensagem ("Falha ao buscar XFDF em {0}: {1}" -f $pastaXFDF, $_.Exception.Message)
            }
        }
    }

    return @($relacionados | Where-Object { $_ } | Sort-Object ProjectID, DocumentID, FileName -Unique)
}

function BuscaDocumentoDuplicadoNoDestino {
    param($Documento, [Parameter(Mandatory = $true)][string]$PastaDestino)

    $numeroDocumento = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente')
    $sequencialEmissao = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'SequencialEmissao')

    if ($numeroDocumento) {
        $encontrados = @(Get-PWDocumentsBySearch -FolderPath $PastaDestino -JustThisFolder -Attributes @{ NumeroPoderConcedente = $numeroDocumento } -GetAttributes -WarningAction SilentlyContinue)
        $duplicado = $encontrados |
            Where-Object {
                $_.Name -eq $Documento.Name -or
                $_.FileName -eq $Documento.FileName -or
                ($sequencialEmissao -and (Normalizar-Texto (Get-ValorAtributo -Documento $_ -Nome 'SequencialEmissao')) -eq $sequencialEmissao)
            } |
            Select-Object -First 1

        if ($duplicado) { return $duplicado }
    }

    try {
        if ($Documento.FileName) {
            $porArquivo = Get-PWDocumentsBySearch -FolderPath $PastaDestino -JustThisFolder -FileName $Documento.FileName -GetAttributes -WarningAction SilentlyContinue
            if ($porArquivo) { return @($porArquivo)[0] }
        }
    }
    catch {
    }

    return $null
}

function ObtemDocumentosAnterioresNaPastaDestino {
    param($PastaDestino, $NumeroDocumento, $SequencialEmissao)

    if ([string]::IsNullOrWhiteSpace($PastaDestino) -or [string]::IsNullOrWhiteSpace($NumeroDocumento) -or [string]::IsNullOrWhiteSpace($SequencialEmissao)) {
        return @()
    }

    $documentos = @(Get-PWDocumentsBySearch -FolderPath $PastaDestino -JustThisFolder -Attributes @{ NumeroPoderConcedente = $NumeroDocumento } -GetAttributes -WarningAction SilentlyContinue)

    return @($documentos | Where-Object {
        $seq = Get-ValorAtributo -Documento $_ -Nome 'SequencialEmissao'
        ($seq -as [int]) -lt ($SequencialEmissao -as [int])
    })
}

function Tratar-Superados {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$PastaDestino
    )

    $numeroDocumento = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente')
    $sequencialEmissao = Normalizar-Texto (Get-ValorAtributo -Documento $Documento -Nome 'SequencialEmissao')
    $documentosAnteriores = @(ObtemDocumentosAnterioresNaPastaDestino -PastaDestino $PastaDestino -NumeroDocumento $numeroDocumento -SequencialEmissao $sequencialEmissao)

    if ($documentosAnteriores.Count -eq 0) {
        Escrever-Log -Mensagem 'Superados: nenhum documento anterior encontrado.'
        return
    }

    $pastaSuperados = Join-Path $PastaDestino $NomePastaSuperados
    Escrever-Log -Mensagem ("Superados: {0} documento(s) anterior(es) para mover para {1}" -f $documentosAnteriores.Count, $pastaSuperados)

    if ($Simular) {
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem "Moveria documentos anteriores para Superados e aplicaria state $StateSuperado."
        return
    }

    $null = Get-PWFoldersCached -FolderPath $pastaSuperados
    $movidos = Move-PWDocumentsToFolder -InputDocument $documentosAnteriores -TargetFolderPath $pastaSuperados -ErrorAction Stop
    if ($movidos) {
        Set-PWDocumentState -InputDocuments $movidos -State $StateSuperado -Force
        $script:Contador.MovidosParaSuperados += @($movidos).Count
    }
}

function Tratar-XFDF {
    param(
        [Parameter(Mandatory = $true)]$DocumentoXFDF,
        [Parameter(Mandatory = $true)][string]$PastaDestinoDocumento
    )

    $pastaDestinoXFDF = Join-Path $PastaDestinoDocumento $NomePastaXFDF

    if ($Simular) {
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem ("Copiaria/moveria XFDF {0} para {1}" -f (Obter-NomeDocumento $DocumentoXFDF), $pastaDestinoXFDF)
        return $null
    }

    if (-not (Get-PWFoldersCached -FolderPath $pastaDestinoXFDF)) {
        $null = New-PWFolder -FolderPath $pastaDestinoXFDF -Environment $EnvironmentXFDF -Workflow $WorkflowComentarios -ErrorAction Stop
        $script:CachePastas[$pastaDestinoXFDF.ToLowerInvariant()] = Get-PWFolders -FolderPath $pastaDestinoXFDF -PopulatePaths -JustOne
    }

    return CopiarOuMover-Documento -Documento $DocumentoXFDF -PastaDestino $pastaDestinoXFDF
}

function CopiarOuMover-Documento {
    param(
        [Parameter(Mandatory = $true)]$Documento,
        [Parameter(Mandatory = $true)][string]$PastaDestino
    )

    $nome = Obter-NomeDocumento -Documento $Documento

    if ($Simular) {
        $acao = if ($MoverDocumentoOrigem) { 'Moveria' } else { 'Copiaria' }
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem ("{0} {1} para {2}" -f $acao, $nome, $PastaDestino)
        return $Documento
    }

    $duplicado = BuscaDocumentoDuplicadoNoDestino -Documento $Documento -PastaDestino $PastaDestino
    if ($duplicado) {
        Escrever-Log -Nivel 'AVISO' -Mensagem ("Documento ja existe no destino: {0}" -f $nome)
        $script:Contador.JaExistiamNoDestino++
        return $duplicado
    }

    try {
        if ($MoverDocumentoOrigem) {
            $movido = Move-PWDocumentsToFolder -InputDocument $Documento -TargetFolderPath $PastaDestino -ErrorAction Stop
            $script:Contador.DocumentosMovidos += @($movido).Count
            return @($movido)[0]
        }

        $copiado = Copy-PWDocumentsToFolder -InputDocument $Documento -TargetFolderPath $PastaDestino -ErrorAction Stop -WarningAction Stop
        $script:Contador.DocumentosCopiados += @($copiado).Count
        return @($copiado)[0]
    }
    catch {
        throw "Falha ao copiar/mover $nome para ${PastaDestino}: $($_.Exception.Message)"
    }
}

function Atualizar-StateFinal {
    param(
        [Parameter(Mandatory = $true)]$DocumentoOrigem,
        $DocumentoDestino
    )

    if ($Simular) {
        Escrever-Log -Nivel 'SIMULACAO' -Mensagem "Atualizaria state da origem para $StateOrigemCopiadoComSucesso."
        if ($DocumentoDestino) {
            Escrever-Log -Nivel 'SIMULACAO' -Mensagem "Atualizaria state do documento no destino para $StateDestinoDocumentoNovo."
        }
        return
    }

    if ($DocumentoDestino) {
        Set-PWDocumentState -InputDocuments $DocumentoDestino -State $StateDestinoDocumentoNovo -Force
        $script:Contador.EstadosAtualizados++
    }

    if (-not $MoverDocumentoOrigem) {
        Set-PWDocumentState -InputDocuments $DocumentoOrigem -State $StateOrigemCopiadoComSucesso -Force
        $script:Contador.EstadosAtualizados++
    }
}

function Exibir-ResumoFinal {
    Escrever-Log -Mensagem ''
    Escrever-Log -Mensagem '============================================================'
    Escrever-Log -Mensagem 'Resumo da reanalise/movimentacao assistida'
    Escrever-Log -Mensagem '============================================================'

    foreach ($item in $script:Contador.GetEnumerator()) {
        Escrever-Log -Mensagem ('{0}: {1}' -f $item.Key, $item.Value)
    }

    Escrever-Log -Mensagem '============================================================'
    Escrever-Log -Mensagem "Log detalhado: $LogPath"
}

#-------------------------------------------------------
# Inicio da execucao
#-------------------------------------------------------
$deveSimular = $Simular -or $WhatIfPreference
if ($deveSimular -and -not $Simular) { $Simular = $true }

try {
    Import-Module pwps_dab -ErrorAction Stop

    Escrever-Log -Mensagem "Iniciando reanalise. Simular: $Simular"

    Connect-ProjectWiseDatasource

    $pastaSelecionada = Selecionar-PastaAnalise -PastaInformada $PastaAnalise
    if ([string]::IsNullOrWhiteSpace($pastaSelecionada)) {
        throw 'Nenhuma pasta de analise foi informada.'
    }

    if (-not (Get-PWFoldersCached -FolderPath $pastaSelecionada)) {
        throw "Pasta de analise nao localizada: $pastaSelecionada"
    }

    $documentos = @(Obter-DocumentosParaReanalise -PastaRaiz $pastaSelecionada)

    $script:Contador.Encontrados = $documentos.Count

    if (-not (Confirmar-ProcessamentoDocumentos -Documentos $documentos -PastaRaiz $pastaSelecionada)) {
        return
    }

    $limiteSelecionado = Selecionar-LimiteDocumentos -TotalDocumentos $documentos.Count
    if ($limiteSelecionado -gt 0 -and $documentos.Count -gt $limiteSelecionado) {
        Escrever-Log -Nivel 'AVISO' -Mensagem "Limite de teste ativo: somente os primeiros $limiteSelecionado documentos serao processados."
        $documentos = @($documentos | Select-Object -First $limiteSelecionado)
    }

    $mapaDisciplinasANTT = Carregar-DisciplinasANTT

    for ($i = 0; $i -lt $documentos.Count; $i++) {
        $documento = $documentos[$i]
        $script:Contador.Processados++

        $nome = Obter-NomeDocumento -Documento $documento
        $percentual = if ($documentos.Count -gt 0) { [math]::Round((($i + 1) / $documentos.Count) * 100, 2) } else { 100 }

        Write-Progress -Activity 'Reanalise/movimentacao assistida' -Status ("Processando {0}/{1} - {2}" -f ($i + 1), $documentos.Count, $nome) -PercentComplete $percentual

        Escrever-Log -Mensagem ''
        Escrever-Log -Mensagem ("Contador: {0}/{1}" -f ($i + 1), $documentos.Count)
        Escrever-Log -Mensagem "Documento: $nome"
        Escrever-Log -Mensagem "State atual: $($documento.WorkflowState)"

        try {
            $erroAtual = Normalizar-Texto (Get-ValorAtributo -Documento $documento -Nome $AtributoErros)
            if (-not [string]::IsNullOrWhiteSpace($erroAtual)) {
                Escrever-Log -Nivel 'AVISO' -Mensagem "Erros atual: $erroAtual"
            }

            if (Testar-ErroValidacaoBloqueante -StateAtual $documento.WorkflowState -MensagemErro $erroAtual) {
                $script:Contador.IgnoradosErroValidacao++
                Escrever-Log -Nivel 'AVISO' -Mensagem 'Documento ignorado: o atributo Erros indica pendencia de validacao, nao problema de pasta destino/movimentacao.'
                continue
            }

            $destino = Montar-CaminhoDestino -Documento $documento -PastaAnaliseSelecionada $pastaSelecionada -MapaDisciplinas $mapaDisciplinasANTT

            Escrever-Log -Mensagem "Poder Concedente: $($destino.PoderConcedente)"
            Escrever-Log -Mensagem "Fase destino: $($destino.FaseDestino)"
            Escrever-Log -Mensagem "Volume: $($destino.Volume)"
            Escrever-Log -Mensagem ("Disciplina: {0} -> {1}" -f $destino.DisciplinaCodigo, $destino.DisciplinaNome)
            Escrever-Log -Mensagem ("Disciplina encontrada na lista oficial: {0}" -f $(if ($destino.DisciplinaLista) { 'Sim' } else { 'Nao' }))
            Escrever-Log -Mensagem ("Tipo Documento: {0} -> {1}" -f $destino.TipoDocumentoCodigo, $destino.TipoDocumentoDestino)

            if ($destino.PastaDestino) {
                Escrever-Log -Mensagem "Destino calculado: $($destino.PastaDestino)"
                Escrever-Log -Mensagem ("Pasta base localizada: {0}" -f $(if ($destino.PastaBaseLocalizada) { 'Sim' } else { 'Nao' }))
            }

            if ($destino.Erros.Count -gt 0) {
                foreach ($erro in $destino.Erros) { Escrever-Log -Nivel 'ERRO' -Mensagem $erro }

                if (-not $destino.PastaBaseLocalizada) { $script:Contador.PastaBaseNaoLocalizada++ }
                if ($destino.Erros -contains 'Disciplina nao localizada na lista oficial e atributo Disciplina Pai vazio.') { $script:Contador.DisciplinaNaoLocalizada++ }
                if ($destino.Erros -contains 'Tipo Documento Pai vazio. Nao foi possivel montar pasta destino.') { $script:Contador.TipoDocumentoVazio++ }

                Registrar-Erro -Documento $documento -Mensagem (($destino.Erros | ForEach-Object { [string]$_ }) -join ' | ')
                continue
            }

            $script:Contador.DestinoCalculado++

            $pastaGarantida = Garantir-PastaDestino -DestinoInfo $destino
            if (-not $pastaGarantida) {
                Registrar-Erro -Documento $documento -Mensagem "Pasta destino nao garantida: $($destino.PastaDestino)"
                continue
            }

            $relacionados = @(Obter-DocumentosRelacionados -Documento $documento -PastaAnaliseSelecionada $pastaSelecionada)
            $nativos = @($relacionados | Where-Object { $_.FileName -notmatch '\.(pdf|xfdf)$' })
            $pdfs = @($relacionados | Where-Object { $_.FileName -match '\.pdf$' })
            $xfdfs = @($relacionados | Where-Object { $_.FileName -match '\.xfdf$' })

            Escrever-Log -Mensagem ("Nativo encontrado: {0}" -f $(if ($nativos.Count -gt 0) { 'Sim' } else { 'Nao' }))
            Escrever-Log -Mensagem ("PDF encontrado: {0}" -f $(if ($pdfs.Count -gt 0) { 'Sim' } else { 'Nao' }))
            Escrever-Log -Mensagem ("XFDF encontrado: {0}" -f $(if ($xfdfs.Count -gt 0) { 'Sim' } else { 'Nao' }))

            Tratar-Superados -Documento $documento -PastaDestino $destino.PastaDestino

            $documentosDestino = New-Object System.Collections.Generic.List[object]
            foreach ($relacionado in $relacionados) {
                if ($relacionado.FileName -match '\.xfdf$') {
                    $copiadoXFDF = Tratar-XFDF -DocumentoXFDF $relacionado -PastaDestinoDocumento $destino.PastaDestino
                    if ($copiadoXFDF) { $documentosDestino.Add($copiadoXFDF) }
                    continue
                }

                $copiado = CopiarOuMover-Documento -Documento $relacionado -PastaDestino $destino.PastaDestino
                if ($copiado) { $documentosDestino.Add($copiado) }
            }

            if ($relacionados.Count -eq 0) {
                $copiadoPrincipal = CopiarOuMover-Documento -Documento $documento -PastaDestino $destino.PastaDestino
                if ($copiadoPrincipal) { $documentosDestino.Add($copiadoPrincipal) }
            }

            $destinoPrincipal = $documentosDestino |
                Where-Object { $_.FileName -eq $documento.FileName -or $_.Name -eq $documento.Name } |
                Select-Object -First 1

            Atualizar-StateFinal -DocumentoOrigem $documento -DocumentoDestino $destinoPrincipal
            Escrever-Log -Nivel 'OK' -Mensagem 'Resultado final: documento reanalisado e fluxo de movimentacao concluido.'
        }
        catch {
            $script:Contador.Erros++
            $mensagem = "Falha ao processar: $($_.Exception.Message)"
            Registrar-Erro -Documento $documento -Mensagem $mensagem
        }
    }
}
finally {
    Write-Progress -Activity 'Reanalise/movimentacao assistida' -Completed

    try {
        Exibir-ResumoFinal
    }
    catch {
        Write-Warning $_.Exception.Message
    }

    if ($script:LoginCriadoPeloScript) {
        try {
            Undo-PWLogin | Out-Null
        }
        catch {
            Write-Warning ('Falha ao encerrar login no ProjectWise: {0}' -f $_.Exception.Message)
        }
    }
}
