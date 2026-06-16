<#
.SYNOPSIS
    Lista grupos vinculados a workflows/states em Work Areas/folders do ProjectWise.

.DESCRIPTION
    Acessa o ProjectWise via Bentley IMS, lista as concessoes dentro de Engenharia,
    permite selecionar uma concessao e gera um inventario das segurancas localizadas,
    com Tipo, Workflow, State e grupos vinculados.

    Este script nao altera o ProjectWise. Ele apenas consulta e exporta relatorio CSV.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Listar_Grupos_Workflows_States.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Listar_Grupos_Workflows_States.ps1" -DatasourceName "01SSRV305.ECSC.ECORODOVIAS.CORP:Ecorodovias-01" -ConcessaoNome "Ecovias"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Listar_Grupos_Workflows_States.ps1" -TipoFiltro group -WorkflowFiltro "Workflow - Engenharia"
#>

[CmdletBinding()]
param(
    [string]$DatasourceName = "",
    [string]$ConcessaoNome = "",
    [string]$ProjetoNome = "",
    [int]$ProfundidadeProjetos = 0,
    [ValidateSet("Workflows", "Folder", "Real")]
    [string]$ModoAccessControl = "Workflows",
    [string]$TipoFiltro = "group",
    [string]$SecurityTypeFiltro = "Document",
    [string]$WorkflowFiltro = "",
    [string]$StateFiltro = "",
    [string]$GrupoFiltro = "",
    [switch]$SomenteConcessao,
    [switch]$UsarBentleyIMS = $true,
    [switch]$NaoDesconectar
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Restart-ScriptInMtaIfNeeded {
    param(
        [hashtable]$BoundParameters
    )

    $apartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartmentState -eq [System.Threading.ApartmentState]::MTA) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw "O PWPS_DAB precisa de PowerShell em modo MTA. Execute o script com: powershell.exe -MTA -File `"<caminho do script>`""
    }

    $powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $powershellExe)) {
        throw "Nao foi possivel localizar o Windows PowerShell 5.1 em '$powershellExe'."
    }

    Write-Host "[INFO] Sessao atual esta em $apartmentState. Reiniciando em PowerShell 5.1 MTA..." -ForegroundColor Yellow

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-MTA",
        "-File",
        $PSCommandPath
    )

    foreach ($key in $BoundParameters.Keys) {
        $value = $BoundParameters[$key]
        if ($value -is [switch] -or $value -is [System.Management.Automation.SwitchParameter]) {
            $arguments += "-$key`:$($value.IsPresent.ToString().ToLowerInvariant())"
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

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArquivoLog = Join-Path $PastaLogs "PW_ListarGruposWorkflowsStates_$TimeStamp.log"
$ArquivoCsv = Join-Path $PastaLogs "PW_ListarGruposWorkflowsStates_$TimeStamp.csv"
$script:LoginPW = $null
$script:SessaoJaExistia = $false
$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering", "ENGENHARIA")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")

function Write-Log {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    $linha = "{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Nivel.ToUpperInvariant(), $Mensagem
    [System.IO.File]::AppendAllText($ArquivoLog, $linha + [Environment]::NewLine)
    Write-Host $linha
}

function Write-Warn {
    param([string]$Mensagem)
    Write-Log -Mensagem $Mensagem -Nivel "WARN"
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
            Write-Log "Modulo '$modulo' ja carregado."
            return
        }

        if (Get-Module -ListAvailable -Name $modulo) {
            Import-Module $modulo -ErrorAction Stop
            $moduloCarregado = Get-Module -Name $modulo
            if ($moduloCarregado) {
                Write-Log "Modulo '$modulo' importado. Versao=$($moduloCarregado.Version) | Caminho=$($moduloCarregado.Path)"
            }
            else {
                Write-Log "Modulo '$modulo' importado."
            }
            return
        }
    }

    throw "Modulo ProjectWise/PWPS_DAB nao encontrado."
}

function Obter-DatasourceAtual {
    $cmd = Get-Command Get-PWCurrentDatasource -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return ""
    }

    try {
        return [string](Get-PWCurrentDatasource -ErrorAction Stop)
    }
    catch {
        return ""
    }
}

function Obter-UsuarioAtual {
    foreach ($nomeCmdlet in @("Get-PWCurrentUser", "Get-PWUserMe")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        try {
            return (& $nomeCmdlet -ErrorAction Stop)
        }
        catch {
            Write-Warn "Falha ao obter usuario atual via $nomeCmdlet`: $($_.Exception.Message)"
        }
    }

    return $null
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

function Obter-CaminhoPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("FullPath", "Fullpath", "FolderPath", "Path", "ProjectPath")
}

function Obter-NomePasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Name", "FolderName", "ProjectName", "ObjectName")
}

function Obter-DescricaoPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Description", "Descricao", "ProjectDescription", "FolderDescription")
}

function Obter-RotuloPasta {
    param([object]$Pasta)

    $descricao = Obter-DescricaoPasta -Pasta $Pasta
    if (-not [string]::IsNullOrWhiteSpace($descricao)) {
        return $descricao
    }

    return Obter-NomePasta -Pasta $Pasta
}

function Ordenar-PastasPorNome {
    param([array]$Pastas)
    return @($Pastas | Sort-Object -Property @{ Expression = { (Obter-RotuloPasta -Pasta $_).ToLowerInvariant() } })
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)
    if ($params -contains "Root") {
        return @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $null -ne $_ })
    }

    return @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $null -ne $_ })
}

function Get-PWChildFoldersOrEmpty {
    param(
        [string]$FolderId,
        [string]$Contexto = ""
    )

    try {
        if ($Contexto) {
            Write-Log "Listando filhos: $Contexto"
        }
        return @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $null -ne $_ })
    }
    catch {
        Write-Warn "Falha ao listar filhos de '$Contexto' ($FolderId): $($_.Exception.Message)"
        return @()
    }
}

function Localizar-PastaPorPossiveisNomes {
    param(
        [array]$Pastas,
        [string[]]$NomesPossiveis,
        [string]$Descricao
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

    $nomes = @($Pastas | ForEach-Object { Obter-RotuloPasta -Pasta $_ }) -join ", "
    throw "Nao encontrei $Descricao. Pastas disponiveis: $nomes"
}

function Localizar-PastaOpcionalPorPossiveisNomes {
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

function Ler-Numero {
    param(
        [string]$Mensagem,
        [int]$Minimo,
        [int]$Maximo
    )

    while ($true) {
        $entrada = Read-Host $Mensagem
        if ($entrada -match '^\d+$') {
            $numero = [int]$entrada
            if ($numero -ge $Minimo -and $numero -le $Maximo) {
                return $numero
            }
        }

        Write-Warn "Informe um numero entre $Minimo e $Maximo."
    }
}

function Obter-ProjetosDaPastaProjetos {
    param(
        [object]$PastaProjetos,
        [string]$NomeBase,
        [int]$ProfundidadeMaxima
    )

    $idPastaProjetos = Obter-IdPasta -Pasta $PastaProjetos
    if ([string]::IsNullOrWhiteSpace($idPastaProjetos)) {
        Write-Warn "Nao foi possivel obter ID da pasta de projetos '$NomeBase'."
        return @()
    }

    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idPastaProjetos; Nivel = 0; Nome = $NomeBase })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId -Contexto $atual.Nome
        foreach ($filho in $filhos) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) {
                continue
            }

            $resultado += $filho
            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $fila.Enqueue([PSCustomObject]@{
                    FolderId = $idFilho
                    Nivel = ($atual.Nivel + 1)
                    Nome = (Obter-RotuloPasta -Pasta $filho)
                })
            }
        }
    }

    $mapa = @{}
    foreach ($item in $resultado) {
        $id = Obter-IdPasta -Pasta $item
        if (-not [string]::IsNullOrWhiteSpace($id) -and -not $mapa.ContainsKey($id)) {
            $mapa[$id] = $item
        }
    }

    return @(Ordenar-PastasPorNome -Pastas @($mapa.Values))
}

function Obter-ConcessoesEngenharia {
    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia"
    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenharia
    if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
        throw "Nao foi possivel obter o ID da pasta Engenharia."
    }

    Write-Log "Pasta Engenharia localizada: $(Obter-RotuloPasta -Pasta $pastaEngenharia)" "OK"
    return @(Ordenar-PastasPorNome -Pastas (Get-PWChildFoldersOrEmpty -FolderId $idEngenharia -Contexto "Engenharia"))
}

function Selecionar-ConcessaoEngenharia {
    $concessoes = @(Obter-ConcessoesEngenharia)
    if ($concessoes.Count -eq 0) {
        throw "Nenhuma concessao encontrada dentro de Engenharia."
    }

    if (-not [string]::IsNullOrWhiteSpace($ConcessaoNome)) {
        $matches = @(
            $concessoes | Where-Object {
                (Obter-RotuloPasta -Pasta $_) -like "*$ConcessaoNome*" -or
                (Obter-NomePasta -Pasta $_) -like "*$ConcessaoNome*"
            }
        )

        if ($matches.Count -eq 1) {
            Write-Log "Concessao localizada pelo filtro '$ConcessaoNome': $(Obter-RotuloPasta -Pasta $matches[0])" "OK"
            return $matches[0]
        }

        if ($matches.Count -gt 1) {
            Write-Warn "Mais de uma concessao encontrada para '$ConcessaoNome'. Selecione pelo numero."
            $concessoes = $matches
        }
        else {
            throw "Nenhuma concessao encontrada contendo '$ConcessaoNome'."
        }
    }

    Write-Host ""
    Write-Host "Concessoes disponiveis em Engenharia:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $concessoes.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Obter-RotuloPasta -Pasta $concessoes[$i]) | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $concessoes.Count
    return $concessoes[$indice - 1]
}

function Selecionar-ProjetoDaConcessao {
    param([object]$Concessao)

    $nomeConcessao = Obter-RotuloPasta -Pasta $Concessao
    $idConcessao = Obter-IdPasta -Pasta $Concessao
    if ([string]::IsNullOrWhiteSpace($idConcessao)) {
        throw "Nao foi possivel obter o ID da concessao '$nomeConcessao'."
    }

    Write-Log "Carregando projetos da concessao '$nomeConcessao'."
    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId $idConcessao -Contexto $nomeConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"
    $projetos = @(Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -NomeBase "$nomeConcessao/Projetos" -ProfundidadeMaxima $ProfundidadeProjetos)

    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto encontrado em '$nomeConcessao'."
    }

    if (-not [string]::IsNullOrWhiteSpace($ProjetoNome)) {
        $matches = @(
            $projetos | Where-Object {
                (Obter-RotuloPasta -Pasta $_) -like "*$ProjetoNome*" -or
                (Obter-CaminhoPasta -Pasta $_) -like "*$ProjetoNome*"
            }
        )

        if ($matches.Count -eq 1) {
            return $matches[0]
        }

        if ($matches.Count -gt 1) {
            Write-Warn "Mais de um projeto encontrado para '$ProjetoNome'. Selecione pelo numero."
            $projetos = $matches
        }
        else {
            throw "Nenhum projeto encontrado contendo '$ProjetoNome' em '$nomeConcessao'."
        }
    }

    Write-Host ""
    Write-Host "Projetos encontrados em '$nomeConcessao':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $projetos.Count; $i++) {
        $caminho = Obter-CaminhoPasta -Pasta $projetos[$i]
        if ([string]::IsNullOrWhiteSpace($caminho)) {
            $caminho = "ID=$(Obter-IdPasta -Pasta $projetos[$i])"
        }

        "{0:000}. {1} | {2}" -f ($i + 1), (Obter-RotuloPasta -Pasta $projetos[$i]), $caminho | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero do projeto que sera processado" -Minimo 1 -Maximo $projetos.Count
    return $projetos[$indice - 1]
}

function Obter-ProjetosDaConcessao {
    param([object]$Concessao)

    $nomeConcessao = Obter-RotuloPasta -Pasta $Concessao
    $idConcessao = Obter-IdPasta -Pasta $Concessao
    if ([string]::IsNullOrWhiteSpace($idConcessao)) {
        throw "Nao foi possivel obter o ID da concessao '$nomeConcessao'."
    }

    Write-Log "Carregando projetos da concessao '$nomeConcessao'."
    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId $idConcessao -Contexto $nomeConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"

    return @(Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -NomeBase "$nomeConcessao/Projetos" -ProfundidadeMaxima $ProfundidadeProjetos)
}

function Obter-ProjetosEngenharia {
    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia"
    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenharia
    if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
        throw "Nao foi possivel obter o ID da pasta Engenharia."
    }

    Write-Log "Pasta Engenharia localizada: $(Obter-RotuloPasta -Pasta $pastaEngenharia)" "OK"
    $basesEngenharia = Ordenar-PastasPorNome -Pastas (Get-PWChildFoldersOrEmpty -FolderId $idEngenharia -Contexto "Engenharia")
    if ($basesEngenharia.Count -eq 0) {
        throw "Nenhuma pasta encontrada dentro de Engenharia."
    }

    $projetosEncontrados = @()
    foreach ($base in $basesEngenharia) {
        $nomeBase = Obter-RotuloPasta -Pasta $base
        $idBase = Obter-IdPasta -Pasta $base
        if ([string]::IsNullOrWhiteSpace($idBase)) {
            continue
        }

        $filhosBase = Get-PWChildFoldersOrEmpty -FolderId $idBase -Contexto $nomeBase
        $pastaProjetos = Localizar-PastaOpcionalPorPossiveisNomes -Pastas $filhosBase -NomesPossiveis $NomesPossiveisPastaProjetos
        if ($null -eq $pastaProjetos) {
            Write-Warn "Pasta Projetos nao encontrada em '$nomeBase'."
            continue
        }

        $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -NomeBase "$nomeBase/Projetos" -ProfundidadeMaxima $ProfundidadeProjetos
        foreach ($projeto in $projetos) {
            $projetosEncontrados += [PSCustomObject]@{
                Base = $nomeBase
                Projeto = Obter-RotuloPasta -Pasta $projeto
                Caminho = Obter-CaminhoPasta -Pasta $projeto
                Id = Obter-IdPasta -Pasta $projeto
                Objeto = $projeto
            }
        }
    }

    return @($projetosEncontrados | Sort-Object Base, Projeto)
}

function Selecionar-ProjetoEngenharia {
    $projetos = @(Obter-ProjetosEngenharia)
    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto encontrado dentro da estrutura Engenharia/*/Projetos."
    }

    if (-not [string]::IsNullOrWhiteSpace($ProjetoNome)) {
        $matches = @(
            $projetos | Where-Object {
                $_.Projeto -like "*$ProjetoNome*" -or
                $_.Caminho -like "*$ProjetoNome*"
            }
        )

        if ($matches.Count -eq 1) {
            Write-Log "Projeto localizado pelo filtro '$ProjetoNome': $($matches[0].Projeto)" "OK"
            return $matches[0]
        }

        if ($matches.Count -gt 1) {
            Write-Warn "Mais de um projeto encontrado para '$ProjetoNome'. Selecione pelo numero."
            $projetos = $matches
        }
        else {
            throw "Nenhum projeto encontrado contendo '$ProjetoNome'."
        }
    }

    Write-Host ""
    Write-Host "Projetos encontrados dentro de Engenharia:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $projetos.Count; $i++) {
        $item = $projetos[$i]
        $caminho = if ([string]::IsNullOrWhiteSpace($item.Caminho)) { "ID=$($item.Id)" } else { $item.Caminho }
        "{0:000}. {1} | {2} | {3}" -f ($i + 1), $item.Base, $item.Projeto, $caminho | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero do projeto que sera processado" -Minimo 1 -Maximo $projetos.Count
    return $projetos[$indice - 1]
}

function Obter-SegurancaPasta {
    param([object]$Pasta)

    $itens = @()

    $cmdFolderSecurity = Get-Command Get-PWFolderSecurity -ErrorAction SilentlyContinue
    if ($cmdFolderSecurity) {
        try {
            $params = @($cmdFolderSecurity.Parameters.Keys)
            if ($params -contains "InputFolder") {
                $itens += @(Get-PWFolderSecurity -InputFolder $Pasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            else {
                $itens += @(Get-PWFolderSecurity $Pasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            Write-Log "Get-PWFolderSecurity retornou $($itens.Count) item(ns)."
        }
        catch {
            Write-Warn "Falha ao consultar Get-PWFolderSecurity: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warn "Cmdlet Get-PWFolderSecurity nao encontrado."
    }

    $cmdProjectSecurity = Get-Command Get-PWProjectSecurity -ErrorAction SilentlyContinue
    if ($cmdProjectSecurity) {
        try {
            $params = @($cmdProjectSecurity.Parameters.Keys)
            $idPasta = Obter-IdPasta -Pasta $Pasta

            if ($params -contains "InputProject") {
                $itens += @(Get-PWProjectSecurity -InputProject $Pasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            elseif ($params -contains "InputFolder") {
                $itens += @(Get-PWProjectSecurity -InputFolder $Pasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            elseif ($params -contains "ProjectID" -and -not [string]::IsNullOrWhiteSpace($idPasta)) {
                $itens += @(Get-PWProjectSecurity -ProjectID $idPasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            elseif ($params -contains "FolderID" -and -not [string]::IsNullOrWhiteSpace($idPasta)) {
                $itens += @(Get-PWProjectSecurity -FolderID $idPasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            else {
                $itens += @(Get-PWProjectSecurity $Pasta -ErrorAction Stop | Where-Object { $null -ne $_ })
            }
            Write-Log "Consulta complementar Get-PWProjectSecurity concluida."
        }
        catch {
            Write-Warn "Falha ao consultar Get-PWProjectSecurity: $($_.Exception.Message)"
        }
    }

    return @($itens)
}

function Testar-ItemSegurancaDoGrupo {
    param(
        [object]$ItemSeguranca,
        [string]$Grupo
    )

    $nomesPreferenciais = @(
        "MemberName",
        "GroupName",
        "Name",
        "ObjectName",
        "UserName",
        "UserListName",
        "SecurityObject",
        "Member",
        "Trustee",
        "Identity"
    )

    foreach ($nome in $nomesPreferenciais) {
        $propriedade = $ItemSeguranca.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = [string]$propriedade.Value
            if ($valor.Trim() -ieq $Grupo) {
                return $true
            }
        }
    }

    foreach ($propriedade in $ItemSeguranca.PSObject.Properties) {
        if ($null -eq $propriedade.Value) {
            continue
        }

        if ($propriedade.Value -is [string] -and $propriedade.Value.Trim() -ieq $Grupo) {
            return $true
        }
    }

    return $false
}

function Verificar-GrupoNoProjeto {
    param(
        [object]$Projeto,
        [string]$Grupo
    )

    $itensSeguranca = @(Obter-SegurancaPasta -Pasta $Projeto)
    if ($itensSeguranca.Count -eq 0) {
        Write-Warn "Nenhum item de seguranca foi retornado para o projeto selecionado."
        return @()
    }

    return @($itensSeguranca | Where-Object { Testar-ItemSegurancaDoGrupo -ItemSeguranca $_ -Grupo $Grupo })
}

function Criar-RegistroVerificacaoProjeto {
    param(
        [object]$Projeto,
        [string]$Grupo,
        [string]$NomeConcessao
    )

    $nomeProjeto = Obter-RotuloPasta -Pasta $Projeto
    $caminhoProjeto = Obter-CaminhoPasta -Pasta $Projeto
    if ([string]::IsNullOrWhiteSpace($caminhoProjeto)) {
        $caminhoProjeto = "ID=$(Obter-IdPasta -Pasta $Projeto)"
    }

    $ocorrencias = @(Verificar-GrupoNoProjeto -Projeto $Projeto -Grupo $Grupo)
    $status = if ($ocorrencias.Count -gt 0) { "ENCONTRADO" } else { "NAO_ENCONTRADO" }
    $detalhe = if ($ocorrencias.Count -gt 0) {
        "Group encontrado na seguranca do projeto. Ocorrencias=$($ocorrencias.Count)"
    }
    else {
        "Group nao apareceu na seguranca retornada para o projeto"
    }

    return [PSCustomObject]@{
        Tipo = "Projeto"
        Concessao = $NomeConcessao
        Nome = $nomeProjeto
        Caminho = $caminhoProjeto
        Grupo = $Grupo
        Status = $status
        Detalhe = $detalhe
    }
}

function Testar-FiltroTexto {
    param(
        [string]$Valor,
        [string]$Filtro
    )

    if ([string]::IsNullOrWhiteSpace($Filtro)) {
        return $true
    }

    return ($Valor -like "*$Filtro*")
}

function Converter-ItemSegurancaParaRegistro {
    param(
        [object]$ItemSeguranca,
        [object]$Pasta,
        [string]$TipoPasta,
        [string]$NomeConcessao
    )

    $tipo = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("Type", "MemberType", "AccessControlType")
    $nome = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("Name", "MemberName", "GroupName", "UserName", "UserListName", "ObjectName")
    $securityType = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("SecurityType", "Security Type")
    $workflow = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("Workflow", "WorkflowName", "WorkFlowName")
    $state = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("State", "StateName", "WorkflowState")
    $access = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("Access_Control_Settings", "AccessControlSettings", "Access", "MemberAccess")
    $inheritedFrom = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("Inheriting_From", "Inherited_From", "InheritedFrom")

    $nomePasta = Obter-RotuloPasta -Pasta $Pasta
    $caminhoPasta = Obter-CaminhoPasta -Pasta $Pasta
    if ([string]::IsNullOrWhiteSpace($caminhoPasta)) {
        $caminhoPasta = Obter-ValorSeguroPropriedade -Objeto $ItemSeguranca -PossiveisNomes @("FullPath")
    }
    if ([string]::IsNullOrWhiteSpace($caminhoPasta)) {
        $caminhoPasta = "ID=$(Obter-IdPasta -Pasta $Pasta)"
    }

    if ($script:ModoAccessControl -eq "Workflows" -and
        [string]::IsNullOrWhiteSpace($workflow) -and
        [string]::IsNullOrWhiteSpace($state)) {
        return $null
    }

    if ($script:ModoAccessControl -eq "Folder" -and
        (-not [string]::IsNullOrWhiteSpace($workflow) -or -not [string]::IsNullOrWhiteSpace($state))) {
        return $null
    }

    if (-not (Testar-FiltroTexto -Valor $tipo -Filtro $script:TipoFiltro)) {
        return $null
    }

    if (-not (Testar-FiltroTexto -Valor $securityType -Filtro $script:SecurityTypeFiltro)) {
        return $null
    }

    if (-not (Testar-FiltroTexto -Valor $workflow -Filtro $script:WorkflowFiltro)) {
        return $null
    }

    if (-not (Testar-FiltroTexto -Valor $state -Filtro $script:StateFiltro)) {
        return $null
    }

    if (-not (Testar-FiltroTexto -Valor $nome -Filtro $script:GrupoFiltro)) {
        return $null
    }

    return [PSCustomObject]@{
        Concessao = $NomeConcessao
        TipoPasta = $TipoPasta
        Pasta = $nomePasta
        Caminho = $caminhoPasta
        Tipo = $tipo
        SecurityType = $securityType
        Workflow = $workflow
        State = $state
        Grupo = $nome
        AccessControl = $access
        HerdadoDe = $inheritedFrom
        PastaObjeto = $Pasta
    }
}

function Listar-GruposWorkflowsStatesDaPasta {
    param(
        [object]$Pasta,
        [string]$TipoPasta,
        [string]$NomeConcessao
    )

    $segurancas = @(Obter-SegurancaPasta -Pasta $Pasta)
    $registros = @()

    foreach ($seguranca in $segurancas) {
        $registro = Converter-ItemSegurancaParaRegistro `
            -ItemSeguranca $seguranca `
            -Pasta $Pasta `
            -TipoPasta $TipoPasta `
            -NomeConcessao $NomeConcessao

        if ($null -ne $registro) {
            $registros += $registro
        }
    }

    return $registros
}

function Conectar-ProjectWise {
    $currentDatasource = Obter-DatasourceAtual
    if (-not [string]::IsNullOrWhiteSpace($currentDatasource)) {
        $script:SessaoJaExistia = $true
        Write-Log "Sessao ProjectWise ja ativa: $currentDatasource" "OK"
        return
    }

    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        $params = @($cmd.Parameters.Keys)

        try {
            Write-Log "Tentando conexao via '$nomeCmdlet'. Parametros: $($params -join ', ')"
            $parametrosComunsLogin = @{}
            if ($params -contains "DoNotCreateWorkingDirectory") {
                $parametrosComunsLogin.DoNotCreateWorkingDirectory = $true
            }

            if (-not [string]::IsNullOrWhiteSpace($DatasourceName) -and $params -contains "DatasourceName") {
                if ($UsarBentleyIMS -and $params -contains "BentleyIMS") {
                    $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName -BentleyIMS @parametrosComunsLogin -ErrorAction Stop
                    Write-Log "Conexao realizada via '$nomeCmdlet -DatasourceName -BentleyIMS'." "OK"
                    return
                }

                if ($params -contains "UseGui") {
                    $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName -UseGui @parametrosComunsLogin -ErrorAction Stop
                    Write-Log "Conexao realizada via '$nomeCmdlet -DatasourceName -UseGui'." "OK"
                    return
                }

                $script:LoginPW = & $nomeCmdlet -DatasourceName $DatasourceName @parametrosComunsLogin -ErrorAction Stop
                Write-Log "Conexao realizada via '$nomeCmdlet -DatasourceName'." "OK"
                return
            }

            if ($UsarBentleyIMS -and $params -contains "BentleyIMS") {
                $script:LoginPW = & $nomeCmdlet -BentleyIMS @parametrosComunsLogin -ErrorAction Stop
                Write-Log "Conexao realizada via '$nomeCmdlet -BentleyIMS'." "OK"
                return
            }

            if ($params -contains "UseGui") {
                $script:LoginPW = & $nomeCmdlet -UseGui @parametrosComunsLogin -ErrorAction Stop
                Write-Log "Conexao realizada via '$nomeCmdlet -UseGui'." "OK"
                return
            }

            $script:LoginPW = & $nomeCmdlet @parametrosComunsLogin -ErrorAction Stop
            Write-Log "Conexao realizada via '$nomeCmdlet'." "OK"
            return
        }
        catch {
            Write-Warn "Falha ao conectar via '$nomeCmdlet'. Detalhe: $($_.Exception.Message)"
        }
    }

    throw "Nao foi possivel conectar ao ProjectWise."
}

function Encerrar-SessaoProjectWise {
    if ($NaoDesconectar -or $script:SessaoJaExistia -or $null -eq $script:LoginPW) {
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

            Write-Log "Sessao ProjectWise encerrada via '$nomeCmdlet'." "OK"
            return
        }
        catch {
            Write-Warn "Falha ao encerrar sessao via '$nomeCmdlet'. Detalhe: $($_.Exception.Message)"
        }
    }
}

function Write-UiLog {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    Write-Log -Mensagem $Mensagem -Nivel $Nivel

    if ($script:LogBox) {
        $linha = "{0} | {1} | {2}" -f (Get-Date -Format "HH:mm:ss"), $Nivel.ToUpperInvariant(), $Mensagem
        $script:LogBox.AppendText($linha + [Environment]::NewLine)
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Show-UiError {
    param([string]$Mensagem)

    [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-UiInfo {
    param([string]$Mensagem)

    [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        "Informacao",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Ask-UiYesNo {
    param(
        [string]$Mensagem,
        [string]$Titulo = "Confirmacao"
    )

    $resposta = [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        $Titulo,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    return ($resposta -eq [System.Windows.Forms.DialogResult]::Yes)
}

function ConvertTo-PWWorkflowSecurityDataTable {
    param([array]$InputRows)

    $table = New-Object System.Data.DataTable
    $columns = @(
        "Concessao",
        "TipoPasta",
        "Pasta",
        "Caminho",
        "Tipo",
        "SecurityType",
        "Workflow",
        "State",
        "Grupo",
        "AccessControl",
        "HerdadoDe"
    )

    foreach ($column in $columns) {
        [void]$table.Columns.Add($column, [string])
    }

    foreach ($row in @($InputRows)) {
        $dataRow = $table.NewRow()
        foreach ($column in $columns) {
            $value = ""
            if ($row.PSObject.Properties[$column] -and $null -ne $row.$column) {
                $value = [string]$row.$column
            }
            $dataRow[$column] = $value
        }
        [void]$table.Rows.Add($dataRow)
    }

    Write-Output -NoEnumerate $table
}

function Get-SelectedWorkflowSecurityRows {
    $selecionados = @()

    foreach ($gridRow in @($script:GridResultados.SelectedRows)) {
        if ($gridRow.IsNewRow) {
            continue
        }

        $caminho = [string]$gridRow.Cells["Caminho"].Value
        $workflow = [string]$gridRow.Cells["Workflow"].Value
        $state = [string]$gridRow.Cells["State"].Value
        $grupo = [string]$gridRow.Cells["Grupo"].Value
        $tipoPasta = [string]$gridRow.Cells["TipoPasta"].Value

        $registro = @(
            $script:ResultadosConsulta | Where-Object {
                $_.Caminho -eq $caminho -and
                $_.Workflow -eq $workflow -and
                $_.State -eq $state -and
                $_.Grupo -eq $grupo -and
                $_.TipoPasta -eq $tipoPasta
            } | Select-Object -First 1
        )

        if ($registro.Count -gt 0) {
            $selecionados += $registro[0]
        }
    }

    return @($selecionados)
}

function Get-FolderAndDescendants {
    param([object]$RootFolder)

    $resultado = New-Object System.Collections.Generic.List[object]
    $visitados = @{}
    $fila = New-Object System.Collections.Queue

    $fila.Enqueue($RootFolder)

    while ($fila.Count -gt 0) {
        $pastaAtual = $fila.Dequeue()
        $idAtual = Obter-IdPasta -Pasta $pastaAtual

        if ([string]::IsNullOrWhiteSpace($idAtual)) {
            continue
        }

        if ($visitados.ContainsKey($idAtual)) {
            continue
        }

        $visitados[$idAtual] = $true
        $resultado.Add($pastaAtual)

        $nomeAtual = Obter-RotuloPasta -Pasta $pastaAtual
        $filhos = @(Get-PWChildFoldersOrEmpty -FolderId $idAtual -Contexto $nomeAtual)
        foreach ($filho in $filhos) {
            $fila.Enqueue($filho)
        }
    }

    return @($resultado.ToArray())
}

function Remover-GruposSelecionadosUi {
    try {
        $selecionados = @(Get-SelectedWorkflowSecurityRows)
        if ($selecionados.Count -eq 0) {
            Show-UiInfo "Selecione uma ou mais linhas na grade antes de remover."
            return
        }

        $invalidos = @(
            $selecionados | Where-Object {
                [string]::IsNullOrWhiteSpace($_.Grupo) -or
                [string]::IsNullOrWhiteSpace($_.Workflow) -or
                [string]::IsNullOrWhiteSpace($_.State)
            }
        )

        if ($invalidos.Count -gt 0) {
            Show-UiError "A remocao exige linhas com Grupo, Workflow e State preenchidos. Revise a selecao."
            return
        }

        $grupos = @($selecionados | Select-Object -ExpandProperty Grupo -Unique)
        $aplicarDescendentes = $script:CheckRemoverDescendentes.Checked
        $alcance = if ($aplicarDescendentes) {
            "concessao, subpastas e child work areas"
        }
        else {
            "somente o item selecionado"
        }

        $mensagem = "Confirma remover $($selecionados.Count) vinculo(s) de seguranca selecionado(s)?`r`n`r`n"
        $mensagem += "Grupo(s): $($grupos -join ', ')`r`n`r`n"
        $mensagem += "Alcance: $alcance`r`n`r`n"
        $mensagem += "Essa acao altera o ProjectWise."

        if (-not (Ask-UiYesNo -Mensagem $mensagem -Titulo "Remover grupo de workflows/states")) {
            return
        }

        $sucessos = 0
        $erros = 0

        foreach ($item in $selecionados) {
            $pastasParaRemover = @($item.PastaObjeto)
            if ($aplicarDescendentes) {
                Write-UiLog "Listando descendentes de '$($item.Pasta)' para aplicar remocao."
                $pastasParaRemover = @(Get-FolderAndDescendants -RootFolder $item.PastaObjeto)
                Write-UiLog "Pastas/work areas no alcance: $($pastasParaRemover.Count)."
            }

            foreach ($pastaRemocao in $pastasParaRemover) {
                $nomePastaRemocao = Obter-RotuloPasta -Pasta $pastaRemocao
                try {
                    Remove-PWFolderSecurity `
                        -InputFolder $pastaRemocao `
                        -DocumentSecurity `
                        -MemberType Group `
                        -MemberName $item.Grupo `
                        -WorkflowName $item.Workflow `
                        -StateName $item.State `
                        -ErrorAction Stop | Out-Null

                    $sucessos++
                    Write-UiLog "Removido: Grupo='$($item.Grupo)' | Workflow='$($item.Workflow)' | State='$($item.State)' | Pasta='$nomePastaRemocao'." "OK"
                }
                catch {
                    $erros++
                    Write-UiLog "Erro ao remover Grupo='$($item.Grupo)' | Workflow='$($item.Workflow)' | State='$($item.State)' | Pasta='$nomePastaRemocao': $($_.Exception.Message)" "ERROR"
                }
            }
        }

        Show-UiInfo "Remocao concluida.`r`nSucessos: $sucessos`r`nErros: $erros`r`n`r`nClique em Consultar novamente para atualizar a grade."
    }
    catch {
        Write-UiLog "Erro na remocao: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
    finally {
        Atualizar-EstadoUi
    }
}

function Atualizar-ResumoUi {
    $rows = @($script:ResultadosConsulta)
    $qtdWorkflows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Workflow) } | Select-Object -ExpandProperty Workflow -Unique).Count
    $qtdStates = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.State) } | Select-Object -ExpandProperty State -Unique).Count
    $qtdGrupos = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Grupo) } | Select-Object -ExpandProperty Grupo -Unique).Count

    $script:LabelResumo.Text = "Registros: $($rows.Count) | Workflows: $qtdWorkflows | States: $qtdStates | Grupos: $qtdGrupos"
}

function Atualizar-EstadoUi {
    $conectado = -not [string]::IsNullOrWhiteSpace((Obter-DatasourceAtual))
    $temConcessao = $script:ComboConcessoes.SelectedIndex -ge 0
    $temResultado = @($script:ResultadosConsulta).Count -gt 0

    $script:BtnConectar.Enabled = -not $conectado
    $script:BtnCarregarConcessoes.Enabled = $conectado
    $script:ComboConcessoes.Enabled = $conectado -and $script:ComboConcessoes.Items.Count -gt 0
    $script:BtnConsultar.Enabled = $conectado -and $temConcessao
    $script:BtnExportar.Enabled = $temResultado
    $script:BtnRemoverSelecionados.Enabled = $conectado -and $temResultado
}

function Conectar-ProjectWiseUi {
    try {
        Write-UiLog "Importando modulo ProjectWise."
        Assert-Mta
        Importar-ModuloProjectWise

        $script:DatasourceName = $script:TextDatasource.Text.Trim()
        $script:UsarBentleyIMS = $true

        Write-UiLog "Conectando via Bentley IMS."
        Conectar-ProjectWise

        $datasourceAtual = Obter-DatasourceAtual
        if ([string]::IsNullOrWhiteSpace($datasourceAtual)) {
            throw "Login aparentemente concluido, mas Get-PWCurrentDatasource nao retornou datasource."
        }

        $usuarioAtual = Obter-UsuarioAtual
        $nomeUsuarioAtual = Obter-ValorSeguroPropriedade -Objeto $usuarioAtual -PossiveisNomes @("Name", "UserName", "LoginName", "Email")
        $script:LabelStatus.Text = "Conectado: $datasourceAtual"

        if (-not [string]::IsNullOrWhiteSpace($nomeUsuarioAtual)) {
            Write-UiLog "Usuario atual: $nomeUsuarioAtual" "OK"
        }

        Write-UiLog "Conexao realizada com sucesso." "OK"
        Carregar-ConcessoesUi
    }
    catch {
        Write-UiLog "Erro ao conectar: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
    finally {
        Atualizar-EstadoUi
    }
}

function Carregar-ConcessoesUi {
    try {
        $script:ComboConcessoes.Items.Clear()
        $script:ConcessoesCarregadas = @(Obter-ConcessoesEngenharia)

        foreach ($concessao in $script:ConcessoesCarregadas) {
            [void]$script:ComboConcessoes.Items.Add((Obter-RotuloPasta -Pasta $concessao))
        }

        if ($script:ComboConcessoes.Items.Count -gt 0) {
            $indiceSelecionado = 0
            if (-not [string]::IsNullOrWhiteSpace($ConcessaoNome)) {
                for ($i = 0; $i -lt $script:ComboConcessoes.Items.Count; $i++) {
                    if ([string]$script:ComboConcessoes.Items[$i] -like "*$ConcessaoNome*") {
                        $indiceSelecionado = $i
                        break
                    }
                }
            }
            $script:ComboConcessoes.SelectedIndex = $indiceSelecionado
        }

        Write-UiLog "Concessoes carregadas: $($script:ComboConcessoes.Items.Count)." "OK"
    }
    catch {
        Write-UiLog "Erro ao carregar concessoes: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
    finally {
        Atualizar-EstadoUi
    }
}

function Obter-ConcessaoSelecionadaUi {
    if ($script:ComboConcessoes.SelectedIndex -lt 0) {
        throw "Selecione uma concessao."
    }

    return $script:ConcessoesCarregadas[$script:ComboConcessoes.SelectedIndex]
}

function Consultar-GruposWorkflowsStatesUi {
    try {
        $script:TipoFiltro = $script:TextTipo.Text.Trim()
        $script:SecurityTypeFiltro = $script:TextSecurityType.Text.Trim()
        $script:WorkflowFiltro = $script:TextWorkflow.Text.Trim()
        $script:StateFiltro = $script:TextState.Text.Trim()
        $script:GrupoFiltro = $script:TextGrupo.Text.Trim()
        $script:ModoAccessControl = [string]$script:ComboModoAccessControl.SelectedItem
        $script:SomenteConcessao = -not $script:CheckIncluirProjetos.Checked
        $script:ProfundidadeProjetos = [int]$script:NumericProfundidade.Value

        $concessaoSelecionada = Obter-ConcessaoSelecionadaUi
        $nomeConcessao = Obter-RotuloPasta -Pasta $concessaoSelecionada

        Write-UiLog "Consultando seguranca da concessao '$nomeConcessao'."
        Write-UiLog "Modo Access Control: '$script:ModoAccessControl'. Workflow, State e Grupo em branco trazem todos os valores encontrados."
        Write-UiLog "Filtros: Tipo='$script:TipoFiltro' | SecurityType='$script:SecurityTypeFiltro' | Workflow='$script:WorkflowFiltro' | State='$script:StateFiltro' | Grupo='$script:GrupoFiltro'."

        $pastasAlvo = @(
            [PSCustomObject]@{
                Tipo = "Concessao"
                Objeto = $concessaoSelecionada
            }
        )

        if ($script:CheckIncluirProjetos.Checked) {
            $projetos = @(Obter-ProjetosDaConcessao -Concessao $concessaoSelecionada)
            foreach ($projeto in $projetos) {
                $pastasAlvo += [PSCustomObject]@{
                    Tipo = "Projeto"
                    Objeto = $projeto
                }
            }
            Write-UiLog "Projetos incluidos na consulta: $($projetos.Count)."
        }

        $resultados = @()
        $contador = 0
        foreach ($pastaAlvo in $pastasAlvo) {
            $contador++
            $nomePasta = Obter-RotuloPasta -Pasta $pastaAlvo.Objeto
            $script:LabelStatus.Text = "Consultando $contador/$($pastasAlvo.Count): $nomePasta"
            [System.Windows.Forms.Application]::DoEvents()

            $resultados += @(Listar-GruposWorkflowsStatesDaPasta `
                -Pasta $pastaAlvo.Objeto `
                -TipoPasta $pastaAlvo.Tipo `
                -NomeConcessao $nomeConcessao)
        }

        $script:ResultadosConsulta = @($resultados | Sort-Object TipoPasta, Pasta, Workflow, State, Grupo)
        $dataTable = ConvertTo-PWWorkflowSecurityDataTable $script:ResultadosConsulta
        $script:GridResultados.DataSource = $null
        $script:GridResultados.DataSource = $dataTable
        $script:LabelStatus.Text = "Consulta concluida."
        Atualizar-ResumoUi

        Write-UiLog "Consulta concluida. Registros encontrados: $($script:ResultadosConsulta.Count)." "OK"
    }
    catch {
        Write-UiLog "Erro na consulta: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
    finally {
        Atualizar-EstadoUi
    }
}

function Exportar-ResultadosUi {
    try {
        if (-not $script:ResultadosConsulta -or $script:ResultadosConsulta.Count -eq 0) {
            Show-UiInfo "Nao ha resultados para exportar."
            return
        }

        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title = "Exportar relatorio"
        $dialog.Filter = "CSV (*.csv)|*.csv"
        $dialog.FileName = [System.IO.Path]::GetFileName($ArquivoCsv)
        $dialog.InitialDirectory = $PastaLogs

        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $script:ResultadosConsulta |
            Select-Object Concessao, TipoPasta, Pasta, Caminho, Tipo, SecurityType, Workflow, State, Grupo, AccessControl, HerdadoDe |
            Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
        Write-UiLog "Relatorio CSV exportado: $($dialog.FileName)" "OK"
        Show-UiInfo "Relatorio exportado com sucesso."
    }
    catch {
        Write-UiLog "Erro ao exportar: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
}

$script:ConcessoesCarregadas = @()
$script:ResultadosConsulta = @()

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "ProjectWise - Listar Grupos por Workflow/State"
$form.Size = New-Object System.Drawing.Size(1180, 760)
$form.MinimumSize = New-Object System.Drawing.Size(980, 660)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"

$labelTitulo = New-Object System.Windows.Forms.Label
$labelTitulo.Text = "ProjectWise - grupos vinculados a workflows e states"
$labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$labelTitulo.Location = New-Object System.Drawing.Point(14, 12)
$labelTitulo.AutoSize = $true
$form.Controls.Add($labelTitulo)

$groupConexao = New-Object System.Windows.Forms.GroupBox
$groupConexao.Text = "Conexao"
$groupConexao.Location = New-Object System.Drawing.Point(14, 45)
$groupConexao.Size = New-Object System.Drawing.Size(1135, 82)
$form.Controls.Add($groupConexao)

$labelDatasource = New-Object System.Windows.Forms.Label
$labelDatasource.Text = "Datasource"
$labelDatasource.Location = New-Object System.Drawing.Point(15, 25)
$labelDatasource.AutoSize = $true
$groupConexao.Controls.Add($labelDatasource)

$script:TextDatasource = New-Object System.Windows.Forms.TextBox
$script:TextDatasource.Location = New-Object System.Drawing.Point(90, 22)
$script:TextDatasource.Size = New-Object System.Drawing.Size(575, 24)
$script:TextDatasource.Text = $DatasourceName
$groupConexao.Controls.Add($script:TextDatasource)

$script:BtnConectar = New-Object System.Windows.Forms.Button
$script:BtnConectar.Text = "Conectar Bentley IMS"
$script:BtnConectar.Location = New-Object System.Drawing.Point(680, 20)
$script:BtnConectar.Size = New-Object System.Drawing.Size(150, 28)
$groupConexao.Controls.Add($script:BtnConectar)

$script:BtnCarregarConcessoes = New-Object System.Windows.Forms.Button
$script:BtnCarregarConcessoes.Text = "Atualizar concessoes"
$script:BtnCarregarConcessoes.Location = New-Object System.Drawing.Point(840, 20)
$script:BtnCarregarConcessoes.Size = New-Object System.Drawing.Size(145, 28)
$groupConexao.Controls.Add($script:BtnCarregarConcessoes)

$script:LabelStatus = New-Object System.Windows.Forms.Label
$script:LabelStatus.Text = "Nao conectado"
$script:LabelStatus.Location = New-Object System.Drawing.Point(15, 55)
$script:LabelStatus.AutoSize = $true
$groupConexao.Controls.Add($script:LabelStatus)

$groupFiltros = New-Object System.Windows.Forms.GroupBox
$groupFiltros.Text = "Concessao e filtros opcionais"
$groupFiltros.Location = New-Object System.Drawing.Point(14, 135)
$groupFiltros.Size = New-Object System.Drawing.Size(1135, 118)
$form.Controls.Add($groupFiltros)

$labelConcessao = New-Object System.Windows.Forms.Label
$labelConcessao.Text = "Concessao"
$labelConcessao.Location = New-Object System.Drawing.Point(15, 27)
$labelConcessao.AutoSize = $true
$groupFiltros.Controls.Add($labelConcessao)

$script:ComboConcessoes = New-Object System.Windows.Forms.ComboBox
$script:ComboConcessoes.Location = New-Object System.Drawing.Point(90, 24)
$script:ComboConcessoes.Size = New-Object System.Drawing.Size(340, 24)
$script:ComboConcessoes.DropDownStyle = "DropDownList"
$groupFiltros.Controls.Add($script:ComboConcessoes)

$script:CheckIncluirProjetos = New-Object System.Windows.Forms.CheckBox
$script:CheckIncluirProjetos.Text = "Incluir projetos tambem"
$script:CheckIncluirProjetos.Checked = $false
$script:CheckIncluirProjetos.Location = New-Object System.Drawing.Point(450, 25)
$script:CheckIncluirProjetos.AutoSize = $true
$groupFiltros.Controls.Add($script:CheckIncluirProjetos)

$labelProfundidade = New-Object System.Windows.Forms.Label
$labelProfundidade.Text = "Profundidade"
$labelProfundidade.Location = New-Object System.Drawing.Point(650, 27)
$labelProfundidade.AutoSize = $true
$groupFiltros.Controls.Add($labelProfundidade)

$script:NumericProfundidade = New-Object System.Windows.Forms.NumericUpDown
$script:NumericProfundidade.Location = New-Object System.Drawing.Point(735, 24)
$script:NumericProfundidade.Size = New-Object System.Drawing.Size(55, 24)
$script:NumericProfundidade.Minimum = 0
$script:NumericProfundidade.Maximum = 5
$script:NumericProfundidade.Value = $ProfundidadeProjetos
$groupFiltros.Controls.Add($script:NumericProfundidade)

$labelModo = New-Object System.Windows.Forms.Label
$labelModo.Text = "Opcao PW"
$labelModo.Location = New-Object System.Drawing.Point(810, 27)
$labelModo.AutoSize = $true
$groupFiltros.Controls.Add($labelModo)

$script:ComboModoAccessControl = New-Object System.Windows.Forms.ComboBox
$script:ComboModoAccessControl.Location = New-Object System.Drawing.Point(880, 24)
$script:ComboModoAccessControl.Size = New-Object System.Drawing.Size(95, 24)
$script:ComboModoAccessControl.DropDownStyle = "DropDownList"
[void]$script:ComboModoAccessControl.Items.Add("Workflows")
[void]$script:ComboModoAccessControl.Items.Add("Folder")
[void]$script:ComboModoAccessControl.Items.Add("Real")
$script:ComboModoAccessControl.SelectedItem = $ModoAccessControl
if ($null -eq $script:ComboModoAccessControl.SelectedItem) {
    $script:ComboModoAccessControl.SelectedIndex = 0
}
$groupFiltros.Controls.Add($script:ComboModoAccessControl)

$labelTipo = New-Object System.Windows.Forms.Label
$labelTipo.Text = "Tipo"
$labelTipo.Location = New-Object System.Drawing.Point(15, 70)
$labelTipo.AutoSize = $true
$groupFiltros.Controls.Add($labelTipo)

$script:TextTipo = New-Object System.Windows.Forms.TextBox
$script:TextTipo.Location = New-Object System.Drawing.Point(55, 67)
$script:TextTipo.Size = New-Object System.Drawing.Size(95, 24)
$script:TextTipo.Text = $TipoFiltro
$groupFiltros.Controls.Add($script:TextTipo)

$labelSecurityType = New-Object System.Windows.Forms.Label
$labelSecurityType.Text = "Security"
$labelSecurityType.Location = New-Object System.Drawing.Point(165, 70)
$labelSecurityType.AutoSize = $true
$groupFiltros.Controls.Add($labelSecurityType)

$script:TextSecurityType = New-Object System.Windows.Forms.TextBox
$script:TextSecurityType.Location = New-Object System.Drawing.Point(225, 67)
$script:TextSecurityType.Size = New-Object System.Drawing.Size(105, 24)
$script:TextSecurityType.Text = $SecurityTypeFiltro
$groupFiltros.Controls.Add($script:TextSecurityType)

$labelWorkflow = New-Object System.Windows.Forms.Label
$labelWorkflow.Text = "Workflow"
$labelWorkflow.Location = New-Object System.Drawing.Point(345, 70)
$labelWorkflow.AutoSize = $true
$groupFiltros.Controls.Add($labelWorkflow)

$script:TextWorkflow = New-Object System.Windows.Forms.TextBox
$script:TextWorkflow.Location = New-Object System.Drawing.Point(415, 67)
$script:TextWorkflow.Size = New-Object System.Drawing.Size(210, 24)
$script:TextWorkflow.Text = $WorkflowFiltro
$groupFiltros.Controls.Add($script:TextWorkflow)

$labelState = New-Object System.Windows.Forms.Label
$labelState.Text = "State"
$labelState.Location = New-Object System.Drawing.Point(640, 70)
$labelState.AutoSize = $true
$groupFiltros.Controls.Add($labelState)

$script:TextState = New-Object System.Windows.Forms.TextBox
$script:TextState.Location = New-Object System.Drawing.Point(685, 67)
$script:TextState.Size = New-Object System.Drawing.Size(180, 24)
$script:TextState.Text = $StateFiltro
$groupFiltros.Controls.Add($script:TextState)

$labelGrupo = New-Object System.Windows.Forms.Label
$labelGrupo.Text = "Grupo"
$labelGrupo.Location = New-Object System.Drawing.Point(880, 70)
$labelGrupo.AutoSize = $true
$groupFiltros.Controls.Add($labelGrupo)

$script:TextGrupo = New-Object System.Windows.Forms.TextBox
$script:TextGrupo.Location = New-Object System.Drawing.Point(930, 67)
$script:TextGrupo.Size = New-Object System.Drawing.Size(190, 24)
$script:TextGrupo.Text = $GrupoFiltro
$groupFiltros.Controls.Add($script:TextGrupo)

$script:BtnConsultar = New-Object System.Windows.Forms.Button
$script:BtnConsultar.Text = "Consultar"
$script:BtnConsultar.Location = New-Object System.Drawing.Point(995, 22)
$script:BtnConsultar.Size = New-Object System.Drawing.Size(140, 28)
$groupFiltros.Controls.Add($script:BtnConsultar)

$labelFluxo = New-Object System.Windows.Forms.Label
$labelFluxo.Text = "Em Workflows, a busca retorna os workflows/states da concessao."
$labelFluxo.Location = New-Object System.Drawing.Point(15, 94)
$labelFluxo.Size = New-Object System.Drawing.Size(520, 18)
$groupFiltros.Controls.Add($labelFluxo)

$groupResultados = New-Object System.Windows.Forms.GroupBox
$groupResultados.Text = "Resultado"
$groupResultados.Location = New-Object System.Drawing.Point(14, 260)
$groupResultados.Size = New-Object System.Drawing.Size(1135, 330)
$form.Controls.Add($groupResultados)

$script:LabelResumo = New-Object System.Windows.Forms.Label
$script:LabelResumo.Text = "Registros: 0 | Workflows: 0 | States: 0 | Grupos: 0"
$script:LabelResumo.Location = New-Object System.Drawing.Point(15, 24)
$script:LabelResumo.AutoSize = $true
$groupResultados.Controls.Add($script:LabelResumo)

$script:CheckRemoverDescendentes = New-Object System.Windows.Forms.CheckBox
$script:CheckRemoverDescendentes.Text = "Remover tambem de subpastas e child work areas"
$script:CheckRemoverDescendentes.Checked = $true
$script:CheckRemoverDescendentes.Location = New-Object System.Drawing.Point(520, 22)
$script:CheckRemoverDescendentes.Size = New-Object System.Drawing.Size(360, 24)
$groupResultados.Controls.Add($script:CheckRemoverDescendentes)

$script:BtnExportar = New-Object System.Windows.Forms.Button
$script:BtnExportar.Text = "Exportar CSV"
$script:BtnExportar.Location = New-Object System.Drawing.Point(895, 18)
$script:BtnExportar.Size = New-Object System.Drawing.Size(105, 28)
$groupResultados.Controls.Add($script:BtnExportar)

$script:BtnRemoverSelecionados = New-Object System.Windows.Forms.Button
$script:BtnRemoverSelecionados.Text = "Remover selecionados"
$script:BtnRemoverSelecionados.Location = New-Object System.Drawing.Point(1008, 18)
$script:BtnRemoverSelecionados.Size = New-Object System.Drawing.Size(112, 28)
$groupResultados.Controls.Add($script:BtnRemoverSelecionados)

$script:GridResultados = New-Object System.Windows.Forms.DataGridView
$script:GridResultados.Location = New-Object System.Drawing.Point(15, 52)
$script:GridResultados.Size = New-Object System.Drawing.Size(1105, 260)
$script:GridResultados.ReadOnly = $true
$script:GridResultados.AllowUserToAddRows = $false
$script:GridResultados.AllowUserToDeleteRows = $false
$script:GridResultados.AutoSizeColumnsMode = "DisplayedCells"
$script:GridResultados.SelectionMode = "FullRowSelect"
$script:GridResultados.MultiSelect = $true
$script:GridResultados.RowHeadersVisible = $false
$script:GridResultados.BackgroundColor = [System.Drawing.Color]::White
$script:GridResultados.BorderStyle = "FixedSingle"
$groupResultados.Controls.Add($script:GridResultados)

$groupLog = New-Object System.Windows.Forms.GroupBox
$groupLog.Text = "Log"
$groupLog.Location = New-Object System.Drawing.Point(14, 598)
$groupLog.Size = New-Object System.Drawing.Size(1135, 82)
$form.Controls.Add($groupLog)

$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Location = New-Object System.Drawing.Point(15, 22)
$script:LogBox.Size = New-Object System.Drawing.Size(1105, 46)
$script:LogBox.Multiline = $true
$script:LogBox.ReadOnly = $true
$script:LogBox.ScrollBars = "Vertical"
$groupLog.Controls.Add($script:LogBox)

$btnFechar = New-Object System.Windows.Forms.Button
$btnFechar.Text = "Fechar"
$btnFechar.Location = New-Object System.Drawing.Point(1050, 686)
$btnFechar.Size = New-Object System.Drawing.Size(95, 28)
$form.Controls.Add($btnFechar)

$groupConexao.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$groupFiltros.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$groupResultados.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$groupLog.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$btnFechar.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$script:TextDatasource.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnConectar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnCarregarConcessoes.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnConsultar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:GridResultados.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:CheckRemoverDescendentes.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnExportar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnRemoverSelecionados.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:LogBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$script:BtnConectar.Add_Click({ Conectar-ProjectWiseUi })
$script:BtnCarregarConcessoes.Add_Click({ Carregar-ConcessoesUi })
$script:BtnConsultar.Add_Click({ Consultar-GruposWorkflowsStatesUi })
$script:BtnExportar.Add_Click({ Exportar-ResultadosUi })
$script:BtnRemoverSelecionados.Add_Click({ Remover-GruposSelecionadosUi })
$btnFechar.Add_Click({ $form.Close() })

$form.Add_Shown({
    Atualizar-EstadoUi
    if (-not [string]::IsNullOrWhiteSpace($DatasourceName)) {
        Conectar-ProjectWiseUi
    }
})

$form.Add_FormClosing({
    try {
        Encerrar-SessaoProjectWise
    }
    catch {
        Write-Log "Falha ao encerrar sessao pela interface: $($_.Exception.Message)" "WARN"
    }
})

Atualizar-EstadoUi
[void]$form.ShowDialog()
