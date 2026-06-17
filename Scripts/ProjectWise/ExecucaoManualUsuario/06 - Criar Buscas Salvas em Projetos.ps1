<#
.SYNOPSIS
    Cria buscas salvas nos projetos selecionados do ProjectWise.

.DESCRIPTION
    Script de execucao manual. Conecta no ProjectWise, permite selecionar uma
    concessao e um ou mais projetos, e cria buscas salvas vinculadas a cada
    projeto selecionado.

    Por padrao, cria as buscas salvas:
    - Em análise da Assistente
    - Em análise da Engenharia
    - Em análise Especifica
    - Enviado para a Unidade
    - Não Validada pelo Sistema

    Tambem oferece buscas opcionais:
    - Enviado ao Poder Concedente
    - Emitido pela Engenharia

    Cada busca usa:
    - Nome do documento = %
    - State correspondente ao nome da pesquisa
    - Escopo = projeto selecionado
    - Incluir subpastas = sim

    No inicio da execucao, o usuario escolhe:
    - ENTER/P para criar somente as buscas padrao
    - T para criar todas as buscas
    - Numeros especificos, como 1,3,6 ou 1-5

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\06 - Criar Buscas Salvas em Projetos.ps1" -WhatIf

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\06 - Criar Buscas Salvas em Projetos.ps1" -ProjectPaths "ENGENHARIA\Ecovias\Projetos\ABC123"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\06 - Criar Buscas Salvas em Projetos.ps1" -EstadosPesquisa "Em analise do assistente"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$DatasourceName = "",
    [switch]$UsarBentleyIMS,
    [string[]]$ProjectPaths,
    [switch]$TodosProjetosDaConcessao,
    [switch]$ReplaceExisting,
    [switch]$NaoDesconectar,
    [switch]$NaoPausar,
    [switch]$RelancadoMTA,
    [switch]$TodasBuscas,
    [string[]]$NomesBuscas,
    [string]$NomePesquisa = "Em análise da Assistente",
    [string[]]$EstadosPesquisa = @("Em analise do assistente")
)

$ErrorActionPreference = "Stop"
$script:LoginPW = $null
$script:ExitCode = 0

$apartmentInicial = [System.Threading.Thread]::CurrentThread.GetApartmentState()
if ($apartmentInicial -ne [System.Threading.ApartmentState]::MTA -and -not $RelancadoMTA -and -not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Write-Host "PowerShell aberto em modo $apartmentInicial. Reabrindo automaticamente em modo MTA..."

    $argumentos = @(
        "-NoProfile",
        "-MTA",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        $PSCommandPath,
        "-RelancadoMTA"
    )

    foreach ($parametro in $PSBoundParameters.GetEnumerator()) {
        if ($parametro.Key -eq "RelancadoMTA") {
            continue
        }

        $valor = $parametro.Value
        if ($valor -is [System.Management.Automation.SwitchParameter]) {
            if ($valor.IsPresent) {
                $argumentos += "-$($parametro.Key)"
            }
            continue
        }

        if ($null -eq $valor) {
            continue
        }

        $argumentos += "-$($parametro.Key)"
        if ($valor -is [array]) {
            foreach ($item in $valor) {
                $argumentos += [string]$item
            }
        }
        else {
            $argumentos += [string]$valor
        }
    }

    & powershell.exe @argumentos
    exit $LASTEXITCODE
}

# ---------------------------------------------------------------------------
# BUSCAS PADRAO
# ---------------------------------------------------------------------------
# Estas configuracoes reproduzem o fluxo manual:
# 1. Busca de documento.
# 2. Campo Nome = * na interface. No cmdlet Add-PWSavedSearch o wildcard e %.
# 3. State conforme cada pesquisa.
# 4. Change/Browse = projeto selecionado.
# 5. Incluir subpastas = marcado.
# 6. Salvar como busca global no projeto selecionado.
$BuscasPadrao = @(
    @{
        Tipo = "Documento"
        Nome = $NomePesquisa
        DocumentName = "%"
        States = $EstadosPesquisa
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $false
    },
    @{
        Tipo = "Documento"
        Nome = "Em análise da Engenharia"
        DocumentName = "%"
        States = @("Em analise da Engenharia")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $false
    },
    @{
        Tipo = "Documento"
        Nome = "Em análise Especifica"
        DocumentName = "%"
        States = @("Em analise Especifica")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $false
    },
    @{
        Tipo = "Documento"
        Nome = "Enviado para a Unidade"
        DocumentName = "%"
        States = @("Enviado para Unidade")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $false
    },
    @{
        Tipo = "Documento"
        Nome = "Não Validada pelo Sistema"
        DocumentName = "%"
        States = @("Nao Validado pelo Sistema")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $false
    },
    @{
        Tipo = "Documento"
        Nome = "Enviado ao Poder Concedente"
        DocumentName = "%"
        States = @("Enviado ao Poder Concedente")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $true
    },
    @{
        Tipo = "Documento"
        Nome = "Emitido pela Engenharia"
        DocumentName = "%"
        States = @("Emitido pela Engenharia")
        SearchSubFolders = $true
        OriginalsOnly = $true
        Opcional = $true
    }
)

function Write-Step {
    param([string]$Mensagem)
    Write-Host ("{0} | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Mensagem)
    Write-Host ("{0} | OK | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Green
}

function Write-Warn {
    param([string]$Mensagem)
    Write-Host ("{0} | AVISO | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Yellow
}

function Pausar-Final {
    if ($NaoPausar) {
        return
    }

    Write-Host ""
    Read-Host "Pressione ENTER para fechar esta janela"
}

function Assert-Mta {
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartment -ne [System.Threading.ApartmentState]::MTA) {
        throw "Execute usando powershell.exe -MTA. Estado atual: $apartment"
    }
}

function Importar-ModuloProjectWise {
    foreach ($modulo in @("pwps_dab", "PWPS_DAB", "Bentley.PowerShell.ProjectWise", "ProjectWisePowerShell")) {
        if (Get-Module -Name $modulo) {
            return
        }

        if (Get-Module -ListAvailable -Name $modulo) {
            Import-Module $modulo -ErrorAction Stop
            return
        }
    }

    throw "Modulo ProjectWise/PWPS_DAB nao encontrado."
}

function Conectar-ProjectWise {
    try {
        $currentDatasource = Get-PWCurrentDatasource -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$currentDatasource)) {
            Write-Ok "Sessao ProjectWise ativa: $currentDatasource"
            return
        }
    }
    catch {
    }

    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        $params = @($cmd.Parameters.Keys)
        try {
            if (-not [string]::IsNullOrWhiteSpace($DatasourceName) -and $params -contains "DatasourceName") {
                if ($UsarBentleyIMS -and $params -contains "BentleyIMS") {
                    Write-Step "Conectando no datasource '$DatasourceName' via Bentley IMS..."
                    $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName -BentleyIMS -ErrorAction Stop
                }
                elseif ($params -contains "UseGui") {
                    Write-Step "Conectando no datasource '$DatasourceName' via GUI..."
                    $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName -UseGui -ErrorAction Stop
                }
                else {
                    Write-Step "Conectando no datasource '$DatasourceName'..."
                    $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName -ErrorAction Stop
                }
                return
            }

            if ($UsarBentleyIMS -and $params -contains "BentleyIMS") {
                Write-Step "Conectando via Bentley IMS..."
                $script:LoginPW = & $nomeCmdlet -BentleyIMS -ErrorAction Stop
                return
            }

            if ($params -contains "UseGui") {
                Write-Step "Conectando via GUI..."
                $script:LoginPW = & $nomeCmdlet -UseGui -ErrorAction Stop
                return
            }

            Write-Step "Conectando via $nomeCmdlet..."
            $script:LoginPW = & $nomeCmdlet -ErrorAction Stop
            return
        }
        catch {
            Write-Warn "Falha em $nomeCmdlet`: $($_.Exception.Message)"
        }
    }

    throw "Nao foi possivel conectar ao ProjectWise."
}

function Encerrar-SessaoProjectWise {
    if ($NaoDesconectar -or $null -eq $script:LoginPW) {
        return
    }

    foreach ($nomeCmdlet in @("Undo-PWLogin", "Close-PWConnection", "Remove-PWLogin")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        try {
            $params = @($cmd.Parameters.Keys)
            if ($params -contains "Login") {
                & $nomeCmdlet -Login $script:LoginPW -ErrorAction Stop | Out-Null
            }
            elseif ($params -contains "InputObject") {
                & $nomeCmdlet -InputObject $script:LoginPW -ErrorAction Stop | Out-Null
            }
            else {
                & $nomeCmdlet -ErrorAction Stop | Out-Null
            }
            Write-Ok "Sessao ProjectWise encerrada."
            return
        }
        catch {
        }
    }
}

function Obter-ValorSeguroPropriedade {
    param(
        [object]$Objeto,
        [string[]]$PossiveisNomes
    )

    if ($null -eq $Objeto) {
        return ""
    }

    foreach ($nome in $PossiveisNomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = $propriedade.Value.ToString().Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
}

function Obter-IdPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID")
}

function Obter-NomePasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Name", "FolderName", "ProjectName", "ObjectName")
}

function Obter-DescricaoPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Description", "Descricao", "ProjectDescription", "FolderDescription")
}

function Obter-CaminhoPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("CaminhoPW", "FullPath", "FolderPath", "Path", "ProjectPath")
}

function Obter-RotuloPasta {
    param([object]$Pasta)
    $descricao = Obter-DescricaoPasta -Pasta $Pasta
    if (-not [string]::IsNullOrWhiteSpace($descricao)) {
        return $descricao
    }

    return Obter-NomePasta -Pasta $Pasta
}

function Set-CaminhoPW {
    param(
        [object]$Pasta,
        [string]$Caminho
    )

    if ($null -eq $Pasta -or [string]::IsNullOrWhiteSpace($Caminho)) {
        return $Pasta
    }

    try {
        $Pasta | Add-Member -NotePropertyName "CaminhoPW" -NotePropertyValue $Caminho -Force
    }
    catch {
    }

    return $Pasta
}

function Join-PWPath {
    param(
        [string]$ParentPath,
        [string]$ChildName
    )

    if ([string]::IsNullOrWhiteSpace($ParentPath)) {
        return $ChildName
    }

    return ("{0}\{1}" -f $ParentPath.TrimEnd("\"), $ChildName.TrimStart("\"))
}

function Ordenar-ItensPorNome {
    param([array]$Itens)
    return @($Itens | Sort-Object -Property @{ Expression = { (Obter-RotuloPasta -Pasta $_).ToLowerInvariant() } })
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)
    if ($params -contains "Root") {
        $pastas = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $null -ne $_ })
    }
    else {
        $pastas = @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $null -ne $_ })
    }

    foreach ($pasta in $pastas) {
        Set-CaminhoPW -Pasta $pasta -Caminho (Obter-NomePasta -Pasta $pasta) | Out-Null
    }

    return $pastas
}

function Get-PWChildFoldersOrEmpty {
    param(
        [string]$FolderId,
        [string]$ParentPath = "",
        [string]$Contexto = ""
    )

    try {
        if ($Contexto) {
            Write-Step "Listando filhos: $Contexto"
        }

        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $null -ne $_ })
        foreach ($pasta in $pastas) {
            $nome = Obter-NomePasta -Pasta $pasta
            Set-CaminhoPW -Pasta $pasta -Caminho (Join-PWPath -ParentPath $ParentPath -ChildName $nome) | Out-Null
        }

        return $pastas
    }
    catch {
        Write-Warn "Falha ao listar filhos de $Contexto ($FolderId): $($_.Exception.Message)"
        return @()
    }
}

function Localizar-PastaPorPossiveisNomes {
    param(
        [array]$Pastas,
        [string[]]$NomesPossiveis
    )

    foreach ($nomeEsperado in $NomesPossiveis) {
        $encontrada = @(
            $Pastas | Where-Object {
                (Obter-NomePasta -Pasta $_) -ieq $nomeEsperado -or
                (Obter-RotuloPasta -Pasta $_) -ieq $nomeEsperado
            }
        )

        if ($encontrada.Count -gt 0) {
            return $encontrada[0]
        }
    }

    return $null
}

function Read-NumberSelection {
    param(
        [string]$Message,
        [int]$Minimum,
        [int]$Maximum,
        [switch]$AllowZero,
        [switch]$AllowAll
    )

    while ($true) {
        $entrada = (Read-Host $Message).Trim()
        if ($AllowAll -and $entrada -ieq "T") {
            return "T"
        }

        $numero = 0
        if ([int]::TryParse($entrada, [ref]$numero)) {
            if ($AllowZero -and $numero -eq 0) {
                return 0
            }

            if ($numero -ge $Minimum -and $numero -le $Maximum) {
                return $numero
            }
        }

        Write-Warn "Opcao invalida. Informe um numero entre $Minimum e $Maximum$(if ($AllowZero) { ', 0' })$(if ($AllowAll) { ' ou T' })."
    }
}

function Read-MultipleNumberSelection {
    param(
        [string]$Message,
        [int]$Maximum
    )

    while ($true) {
        $entrada = (Read-Host $Message).Trim()
        if ($entrada -ieq "T") {
            return @(1..$Maximum)
        }

        $selecionados = New-Object System.Collections.Generic.List[int]
        $partes = @($entrada -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

        foreach ($parte in $partes) {
            if ($parte -match "^\d+\-\d+$") {
                $limites = $parte -split "-"
                $inicio = [int]$limites[0]
                $fim = [int]$limites[1]
                if ($inicio -le $fim -and $inicio -ge 1 -and $fim -le $Maximum) {
                    foreach ($numero in $inicio..$fim) {
                        if (-not $selecionados.Contains($numero)) {
                            $selecionados.Add($numero)
                        }
                    }
                }
            }
            else {
                $numero = 0
                if ([int]::TryParse($parte, [ref]$numero) -and $numero -ge 1 -and $numero -le $Maximum) {
                    if (-not $selecionados.Contains($numero)) {
                        $selecionados.Add($numero)
                    }
                }
            }
        }

        if ($selecionados.Count -gt 0) {
            return @($selecionados | Sort-Object)
        }

        Write-Warn "Opcao invalida. Use exemplos como: 1,3,5 ou 2-6 ou T para todos."
    }
}

function Read-SearchSelection {
    param(
        [string]$Message,
        [int]$Maximum,
        [int[]]$DefaultSelection
    )

    while ($true) {
        $entrada = (Read-Host $Message).Trim()
        if ([string]::IsNullOrWhiteSpace($entrada) -or $entrada -ieq "P") {
            return @($DefaultSelection)
        }

        if ($entrada -ieq "T") {
            return @(1..$Maximum)
        }

        $selecionados = New-Object System.Collections.Generic.List[int]
        $partes = @($entrada -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

        foreach ($parte in $partes) {
            if ($parte -match "^\d+\-\d+$") {
                $limites = $parte -split "-"
                $inicio = [int]$limites[0]
                $fim = [int]$limites[1]
                if ($inicio -le $fim -and $inicio -ge 1 -and $fim -le $Maximum) {
                    foreach ($numero in $inicio..$fim) {
                        if (-not $selecionados.Contains($numero)) {
                            $selecionados.Add($numero)
                        }
                    }
                }
            }
            else {
                $numero = 0
                if ([int]::TryParse($parte, [ref]$numero) -and $numero -ge 1 -and $numero -le $Maximum) {
                    if (-not $selecionados.Contains($numero)) {
                        $selecionados.Add($numero)
                    }
                }
            }
        }

        if ($selecionados.Count -gt 0) {
            return @($selecionados | Sort-Object)
        }

        Write-Warn "Opcao invalida. Use P para padrao, T para todas, ou numeros como: 1,3,5 ou 2-6."
    }
}

function Select-BuscasParaCriacao {
    param([array]$Buscas)

    if ($TodasBuscas) {
        return @($Buscas)
    }

    if ($NomesBuscas -and $NomesBuscas.Count -gt 0) {
        $buscasPorNome = New-Object System.Collections.Generic.List[object]
        foreach ($nomeBuscaInformado in $NomesBuscas) {
            $encontrada = @($Buscas | Where-Object { [string]$_.Nome -ieq [string]$nomeBuscaInformado })
            if ($encontrada.Count -eq 0) {
                throw "Busca informada nao encontrada: $nomeBuscaInformado"
            }

            foreach ($buscaEncontrada in $encontrada) {
                $buscasPorNome.Add($buscaEncontrada)
            }
        }

        return @($buscasPorNome)
    }

    Write-Host ""
    Write-Host "Buscas disponiveis para criar:" -ForegroundColor White

    $indicesPadrao = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $Buscas.Count; $i++) {
        $busca = $Buscas[$i]
        $indice = $i + 1
        $tipo = if ($busca.ContainsKey("Opcional") -and [bool]$busca.Opcional) { "OPCIONAL" } else { "PADRAO" }
        if ($tipo -eq "PADRAO") {
            $indicesPadrao.Add($indice)
        }

        Write-Host ("[{0}] [{1}] {2}" -f $indice, $tipo, $busca.Nome)
    }

    $selecionados = Read-SearchSelection -Message "Selecione as buscas (ENTER/P=padrao, T=todas, ex.: 1,3,6 ou 1-5)" -Maximum $Buscas.Count -DefaultSelection @($indicesPadrao)
    return @($selecionados | ForEach-Object { $Buscas[$_ - 1] })
}

function Select-FolderFromList {
    param(
        [array]$Folders,
        [string]$Title,
        [string]$Prompt
    )

    if (-not $Folders -or $Folders.Count -eq 0) {
        throw "Nenhuma pasta encontrada para selecao."
    }

    $ordenadas = Ordenar-ItensPorNome -Itens $Folders
    Write-Host ""
    Write-Host $Title -ForegroundColor White
    for ($i = 0; $i -lt $ordenadas.Count; $i++) {
        $rotulo = Obter-RotuloPasta -Pasta $ordenadas[$i]
        $nome = Obter-NomePasta -Pasta $ordenadas[$i]
        $sufixo = if ($rotulo -ne $nome -and -not [string]::IsNullOrWhiteSpace($nome)) { " [$nome]" } else { "" }
        Write-Host ("[{0}] {1}{2}" -f ($i + 1), $rotulo, $sufixo)
    }

    $selection = Read-NumberSelection -Message $Prompt -Minimum 1 -Maximum $ordenadas.Count
    return $ordenadas[$selection - 1]
}

function Select-MultipleFoldersFromList {
    param(
        [array]$Folders,
        [string]$Title,
        [string]$Prompt
    )

    if (-not $Folders -or $Folders.Count -eq 0) {
        throw "Nenhuma pasta encontrada para selecao."
    }

    $ordenadas = Ordenar-ItensPorNome -Itens $Folders
    Write-Host ""
    Write-Host $Title -ForegroundColor White
    for ($i = 0; $i -lt $ordenadas.Count; $i++) {
        $rotulo = Obter-RotuloPasta -Pasta $ordenadas[$i]
        $nome = Obter-NomePasta -Pasta $ordenadas[$i]
        $sufixo = if ($rotulo -ne $nome -and -not [string]::IsNullOrWhiteSpace($nome)) { " [$nome]" } else { "" }
        Write-Host ("[{0}] {1}{2}" -f ($i + 1), $rotulo, $sufixo)
    }

    $selecionados = Read-MultipleNumberSelection -Message $Prompt -Maximum $ordenadas.Count
    return @($selecionados | ForEach-Object { $ordenadas[$_ - 1] })
}

function Select-ProjetosFromConsole {
    $rootFolders = Get-PWRootFoldersSafe
    $engenharia = Localizar-PastaPorPossiveisNomes -Pastas $rootFolders -NomesPossiveis @("Engenharia", "Engineering")

    if ($null -eq $engenharia) {
        Write-Warn "Pasta Engenharia nao encontrada no nivel raiz. Listando concessoes a partir da raiz."
        $baseConcessoes = $rootFolders
    }
    else {
        $engenhariaId = Obter-IdPasta -Pasta $engenharia
        $baseConcessoes = Get-PWChildFoldersOrEmpty -FolderId $engenhariaId -ParentPath (Obter-CaminhoPasta -Pasta $engenharia) -Contexto "Engenharia"
    }

    $concessao = Select-FolderFromList -Folders $baseConcessoes -Title "Concessoes encontradas:" -Prompt "Selecione a concessao"
    $concessaoId = Obter-IdPasta -Pasta $concessao
    $concessaoPath = Obter-CaminhoPasta -Pasta $concessao

    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId $concessaoId -ParentPath $concessaoPath -Contexto (Obter-RotuloPasta -Pasta $concessao)
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis @("Projetos", "Projeto", "Projects", "Project")

    if ($null -eq $pastaProjetos) {
        Write-Warn "Pasta Projetos nao encontrada. Listando projetos diretamente dentro da concessao."
        $projetos = $filhosConcessao
    }
    else {
        $projetos = Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $pastaProjetos) -ParentPath (Obter-CaminhoPasta -Pasta $pastaProjetos) -Contexto (Obter-RotuloPasta -Pasta $pastaProjetos)
    }

    if ($TodosProjetosDaConcessao) {
        return @(Ordenar-ItensPorNome -Itens $projetos)
    }

    return Select-MultipleFoldersFromList -Folders $projetos -Title "Projetos encontrados:" -Prompt "Selecione os projetos (ex.: 1,3,5 ou 2-6 ou T para todos)"
}

function Get-ProjetosPorCaminho {
    param([string[]]$Paths)

    $projetos = New-Object System.Collections.Generic.List[object]
    foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        Write-Step "Localizando projeto: $path"
        $projeto = Get-PWFolders -FolderPath $path -JustOne -PopulatePaths -ErrorAction Stop
        if ($null -eq $projeto) {
            throw "Projeto nao encontrado: $path"
        }

        Set-CaminhoPW -Pasta $projeto -Caminho $path | Out-Null
        $projetos.Add($projeto)
    }

    return @($projetos)
}

function Resolver-SearchFolder {
    param(
        [object]$Projeto,
        [hashtable]$Busca
    )

    $projectPath = Obter-CaminhoPasta -Pasta $Projeto
    if ([string]::IsNullOrWhiteSpace($projectPath)) {
        throw "Nao foi possivel determinar o caminho do projeto '$((Obter-RotuloPasta -Pasta $Projeto))'."
    }

    if ($Busca.ContainsKey("SearchFolder") -and -not [string]::IsNullOrWhiteSpace([string]$Busca.SearchFolder)) {
        return [string]$Busca.SearchFolder
    }

    if ($Busca.ContainsKey("SearchSubFolder") -and -not [string]::IsNullOrWhiteSpace([string]$Busca.SearchSubFolder)) {
        return Join-PWPath -ParentPath $projectPath -ChildName ([string]$Busca.SearchSubFolder)
    }

    return $projectPath
}

function Add-ParametroSeInformado {
    param(
        [hashtable]$Parametros,
        [hashtable]$Origem,
        [string]$Nome
    )

    if ($Origem.ContainsKey($Nome) -and $null -ne $Origem[$Nome]) {
        if ($Origem[$Nome] -is [string] -and [string]::IsNullOrWhiteSpace($Origem[$Nome])) {
            return
        }

        $Parametros[$Nome] = $Origem[$Nome]
    }
}

function Obter-NomeBuscaSalva {
    param([object]$BuscaSalva)

    return Obter-ValorSeguroPropriedade -Objeto $BuscaSalva -PossiveisNomes @(
        "SearchName",
        "SavedSearchName",
        "Name",
        "FullPath",
        "Path",
        "ObjectName"
    )
}

function Test-BuscaSalvaExistente {
    param(
        [object]$Projeto,
        [string]$NomeBusca
    )

    if ($ReplaceExisting) {
        return $false
    }

    $buscasExistentes = @(Get-PWSavedSearches -InputFolder $Projeto -ErrorAction Stop | Where-Object { $null -ne $_ })
    foreach ($buscaExistente in $buscasExistentes) {
        $nomeExistente = Obter-NomeBuscaSalva -BuscaSalva $buscaExistente
        if ([string]::IsNullOrWhiteSpace($nomeExistente)) {
            continue
        }

        if ($nomeExistente -ieq $NomeBusca -or $nomeExistente -ilike "*\$NomeBusca") {
            return $true
        }
    }

    return $false
}

function Test-AvisoBuscaJaExiste {
    param([object[]]$Avisos)

    foreach ($aviso in @($Avisos)) {
        if ($null -eq $aviso) {
            continue
        }

        $textoAviso = [string]$aviso
        if ($textoAviso -match "already exists" -or $textoAviso -match "ja existe" -or $textoAviso -match "já existe") {
            return $true
        }
    }

    return $false
}

function Nova-BuscaDocumento {
    param(
        [object]$Projeto,
        [hashtable]$Busca
    )

    $nomeBusca = [string]$Busca.Nome
    $searchFolder = Resolver-SearchFolder -Projeto $Projeto -Busca $Busca

    if (Test-BuscaSalvaExistente -Projeto $Projeto -NomeBusca $nomeBusca) {
        return "Ignorada"
    }

    $params = @{
        OwnerProject = $Projeto
        SearchName = $nomeBusca
        SearchFolder = $searchFolder
        ErrorAction = "Stop"
    }

    if ($Busca.ContainsKey("SearchSubFolders") -and [bool]$Busca.SearchSubFolders) { $params.SearchSubFolders = $true }
    if ($Busca.ContainsKey("OriginalsOnly") -and [bool]$Busca.OriginalsOnly) { $params.OriginalsOnly = $true }
    if ($Busca.ContainsKey("AllEnvironments") -and [bool]$Busca.AllEnvironments) { $params.AllEnvironments = $true }
    if ($Busca.ContainsKey("WholeDatasource") -and [bool]$Busca.WholeDatasource) { $params.WholeDatasource = $true }
    if ($Busca.ContainsKey("WholePhrase") -and [bool]$Busca.WholePhrase) { $params.WholePhrase = $true }
    if ($Busca.ContainsKey("AnyWord") -and [bool]$Busca.AnyWord) { $params.AnyWord = $true }
    if ($Busca.ContainsKey("SearchAttributes") -and [bool]$Busca.SearchAttributes) { $params.SearchAttributes = $true }
    if ($ReplaceExisting) { $params.ReplaceExisting = $true }

    foreach ($campo in @(
        "Description", "Environment", "Attributes", "Workflow", "States",
        "UpdatedAfter", "UpdatedBefore", "ViewName", "Status", "DocumentGUID",
        "SearchText", "DocumentName", "FileName"
    )) {
        Add-ParametroSeInformado -Parametros $params -Origem $Busca -Nome $campo
    }

    if ($PSCmdlet.ShouldProcess((Obter-RotuloPasta -Pasta $Projeto), "Criar busca de documento '$nomeBusca'")) {
        $avisosBusca = @()
        $params.WarningAction = "SilentlyContinue"
        $params.WarningVariable = "avisosBusca"
        Add-PWSavedSearch @params | Out-Null

        if (Test-AvisoBuscaJaExiste -Avisos $avisosBusca) {
            return "Ignorada"
        }

        return "Criada"
    }

    return "Simulada"
}

function Nova-BuscaPasta {
    param(
        [object]$Projeto,
        [hashtable]$Busca
    )

    $nomeBusca = [string]$Busca.Nome
    $searchFolder = Resolver-SearchFolder -Projeto $Projeto -Busca $Busca

    if (Test-BuscaSalvaExistente -Projeto $Projeto -NomeBusca $nomeBusca) {
        return "Ignorada"
    }

    $params = @{
        OwnerProject = $Projeto
        SearchName = $nomeBusca
        SearchFolder = $searchFolder
        ErrorAction = "Stop"
    }

    if ($Busca.ContainsKey("EmptyFolders") -and [bool]$Busca.EmptyFolders) { $params.EmptyFolders = $true }
    if ($Busca.ContainsKey("WholeDatasource") -and [bool]$Busca.WholeDatasource) { $params.WholeDatasource = $true }
    if ($ReplaceExisting) { $params.ReplaceExisting = $true }

    foreach ($campo in @(
        "ProjectType", "FolderName", "ProjectProperties", "Environment",
        "Workflow", "UpdatedAfter", "UpdatedBefore"
    )) {
        Add-ParametroSeInformado -Parametros $params -Origem $Busca -Nome $campo
    }

    if ($PSCmdlet.ShouldProcess((Obter-RotuloPasta -Pasta $Projeto), "Criar busca de pasta '$nomeBusca'")) {
        $avisosBusca = @()
        $params.WarningAction = "SilentlyContinue"
        $params.WarningVariable = "avisosBusca"
        Add-PWSavedFolderSearch @params | Out-Null

        if (Test-AvisoBuscaJaExiste -Avisos $avisosBusca) {
            return "Ignorada"
        }

        return "Criada"
    }

    return "Simulada"
}

function Validar-Buscas {
    param([array]$Buscas)

    if (-not $Buscas -or $Buscas.Count -eq 0) {
        throw "Nenhuma busca configurada em `$BuscasPadrao."
    }

    foreach ($busca in $Buscas) {
        if (-not ($busca -is [hashtable])) {
            throw "Cada busca deve ser uma hashtable."
        }

        if (-not $busca.ContainsKey("Nome") -or [string]::IsNullOrWhiteSpace([string]$busca.Nome)) {
            throw "Existe busca sem Nome configurado."
        }

        if (-not $busca.ContainsKey("Tipo") -or [string]::IsNullOrWhiteSpace([string]$busca.Tipo)) {
            throw "A busca '$($busca.Nome)' esta sem Tipo."
        }

        if (@("Documento", "Pasta") -notcontains [string]$busca.Tipo) {
            throw "Tipo invalido na busca '$($busca.Nome)': $($busca.Tipo). Use Documento ou Pasta."
        }
    }
}

try {
    Assert-Mta
    $buscasSelecionadas = Select-BuscasParaCriacao -Buscas $BuscasPadrao
    Importar-ModuloProjectWise
    Validar-Buscas -Buscas $buscasSelecionadas
    Conectar-ProjectWise

    if ($ProjectPaths -and $ProjectPaths.Count -gt 0) {
        $projetosSelecionados = Get-ProjetosPorCaminho -Paths $ProjectPaths
    }
    else {
        $projetosSelecionados = Select-ProjetosFromConsole
    }

    if (-not $projetosSelecionados -or $projetosSelecionados.Count -eq 0) {
        throw "Nenhum projeto selecionado."
    }

    Write-Step "Projetos selecionados: $($projetosSelecionados.Count)"
    Write-Step "Buscas selecionadas: $($buscasSelecionadas.Count)"

    $totalCriadas = 0
    $totalIgnoradas = 0
    $totalFalhas = 0

    foreach ($projeto in $projetosSelecionados) {
        $rotuloProjeto = Obter-RotuloPasta -Pasta $projeto
        Write-Host ""
        Write-Step "Processando projeto: $rotuloProjeto"

        foreach ($busca in $buscasSelecionadas) {
            try {
                if ([string]$busca.Tipo -ieq "Documento") {
                    $resultadoBusca = Nova-BuscaDocumento -Projeto $projeto -Busca $busca
                }
                else {
                    $resultadoBusca = Nova-BuscaPasta -Projeto $projeto -Busca $busca
                }

                if ($resultadoBusca -eq "Ignorada") {
                    $totalIgnoradas++
                    Write-Warn "Busca ja existe; nao sera criada novamente: $($busca.Nome)"
                }
                elseif ($resultadoBusca -eq "Simulada") {
                    $totalCriadas++
                    Write-Ok "Busca simulada: $($busca.Nome)"
                }
                else {
                    $totalCriadas++
                    Write-Ok "Busca criada/validada: $($busca.Nome)"
                }
            }
            catch {
                $totalFalhas++
                Write-Warn "Falha na busca '$($busca.Nome)' no projeto '$rotuloProjeto': $($_.Exception.Message)"
            }
        }
    }

    Write-Host ""
    Write-Ok "Finalizado. Criadas/Simuladas: $totalCriadas | Ignoradas por ja existir: $totalIgnoradas | Falhas: $totalFalhas"
}
catch {
    $script:ExitCode = 1
    Write-Host ""
    Write-Host "ERRO NA EXECUCAO DO SCRIPT" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red

    if ($_.InvocationInfo) {
        Write-Host ""
        Write-Host ("Linha: {0}" -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Yellow
        Write-Host $_.InvocationInfo.Line -ForegroundColor Yellow
    }
}
finally {
    Encerrar-SessaoProjectWise
    Pausar-Final
}

if ($script:ExitCode -ne 0) {
    exit $script:ExitCode
}
