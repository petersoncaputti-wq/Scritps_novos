<#
.SYNOPSIS
    Inclui grupo na seguranca de Work Areas/folders e workflows do ProjectWise.

.DESCRIPTION
    Versao inicial para validar apenas a conexao com o ProjectWise.
    As etapas de leitura de folders, workflow security e aplicacao de permissoes
    serao adicionadas apos a validacao do login.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Incluir_Grupo_Seguranca_WorkArea_Workflow.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Incluir_Grupo_Seguranca_WorkArea_Workflow.ps1" -DatasourceName "01SSRV305.ECSC.ECORODOVIAS.CORP:Ecorodovias-01"
#>

[CmdletBinding()]
param(
    [string]$DatasourceName = "",
    [string]$NomeGrupo = "",
    [string]$ConcessaoNome = "",
    [string]$ProjetoNome = "",
    [int]$ProfundidadeProjetos = 0,
    [string[]]$PermissaoPasta = @("r"),
    [switch]$Executar,
    [switch]$UsarBentleyIMS = $true,
    [switch]$NaoDesconectar
)

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
$ArquivoLog = Join-Path $PastaLogs "PW_IncluirGrupoSegurancaWorkAreaWorkflow_$TimeStamp.log"
$ArquivoCsv = Join-Path $PastaLogs "PW_IncluirGrupoSegurancaWorkAreaWorkflow_$TimeStamp.csv"
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

function Obter-NomeGrupo {
    param([object]$Grupo)

    if ($null -eq $Grupo) {
        return ""
    }

    if ($Grupo -is [string]) {
        return $Grupo.Trim()
    }

    return Obter-ValorSeguroPropriedade -Objeto $Grupo -PossiveisNomes @("Name", "GroupName", "ObjectName", "MemberName", "FullName")
}

function Obter-TodosGrupos {
    if (-not (Get-Command Get-PWGroups -ErrorAction SilentlyContinue)) {
        throw "Cmdlet Get-PWGroups nao encontrado; nao foi possivel validar o grupo no PW Administrator."
    }

    $grupos = @(Get-PWGroups -ErrorAction Stop | Where-Object { $null -ne $_ })
    if ($grupos.Count -eq 0) {
        throw "Nenhum Group retornado pelo ProjectWise."
    }

    return @($grupos | Sort-Object @{ Expression = { (Obter-NomeGrupo -Grupo $_).ToLowerInvariant() } })
}

function Resolver-GrupoPorTexto {
    param([string]$Texto)

    while ([string]::IsNullOrWhiteSpace($Texto)) {
        $Texto = (Read-Host "Nome do Group que sera incluido nos projetos").Trim()
    }

    $todosGrupos = Obter-TodosGrupos
    $exatos = @(
        $todosGrupos | Where-Object {
            (Obter-NomeGrupo -Grupo $_) -ieq $Texto
        }
    )

    if ($exatos.Count -eq 1) {
        $nomeExato = Obter-NomeGrupo -Grupo $exatos[0]
        Write-Log "Group validado por nome exato: $nomeExato" "OK"
        return $nomeExato
    }

    $termos = @($Texto -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $encontrados = @(
        $todosGrupos | Where-Object {
            $nome = Obter-NomeGrupo -Grupo $_
            if ([string]::IsNullOrWhiteSpace($nome)) {
                return $false
            }

            foreach ($termo in $termos) {
                if ($nome -notlike "*$termo*") {
                    return $false
                }
            }

            return $true
        }
    )

    if ($encontrados.Count -eq 0) {
        throw "Nenhum Group encontrado no PW Administrator contendo o texto '$Texto'."
    }

    if ($encontrados.Count -eq 1) {
        $nomeEncontrado = Obter-NomeGrupo -Grupo $encontrados[0]
        Write-Log "Group validado por busca textual '$Texto': $nomeEncontrado" "OK"
        return $nomeEncontrado
    }

    Write-Host ""
    Write-Host "Groups encontrados para '$Texto':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $encontrados.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Obter-NomeGrupo -Grupo $encontrados[$i]) | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero do Group" -Minimo 1 -Maximo $encontrados.Count
    $nomeSelecionado = Obter-NomeGrupo -Grupo $encontrados[$indice - 1]
    Write-Log "Group selecionado e validado: $nomeSelecionado" "OK"
    return $nomeSelecionado
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

    $indice = Ler-Numero -Mensagem "Numero do projeto que recebera o grupo" -Minimo 1 -Maximo $projetos.Count
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

    $indice = Ler-Numero -Mensagem "Numero do projeto que recebera o grupo" -Minimo 1 -Maximo $projetos.Count
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

function Aplicar-GrupoNaConcessao {
    param(
        [object]$Concessao,
        [string]$Grupo
    )

    if (-not (Get-Command Update-PWFolderSecurity -ErrorAction SilentlyContinue)) {
        throw "Cmdlet Update-PWFolderSecurity nao encontrado; nao foi possivel aplicar o grupo."
    }

    $nomeConcessao = Obter-RotuloPasta -Pasta $Concessao
    $caminhoConcessao = Obter-CaminhoPasta -Pasta $Concessao
    if ([string]::IsNullOrWhiteSpace($caminhoConcessao)) {
        $caminhoConcessao = $nomeConcessao
    }

    if (-not $Executar) {
        Write-Warn "Modo simulacao ativo. O Group '$Grupo' seria incluido na concessao '$nomeConcessao'. Use -Executar para gravar."
        return [PSCustomObject]@{
            Tipo = "Concessao"
            Nome = $nomeConcessao
            Caminho = $caminhoConcessao
            Grupo = $Grupo
            Status = "SIMULADO"
            Detalhe = "FolderSecurity seria aplicado na concessao"
        }
    }

    try {
        Update-PWFolderSecurity -InputFolder $Concessao -FolderSecurity -MemberType Group -MemberName $Grupo -MemberAccess $PermissaoPasta -ErrorAction Stop | Out-Null
        Write-Log "Group '$Grupo' incluido na FolderSecurity da concessao '$nomeConcessao'." "OK"

        return [PSCustomObject]@{
            Tipo = "Concessao"
            Nome = $nomeConcessao
            Caminho = $caminhoConcessao
            Grupo = $Grupo
            Status = "SUCESSO"
            Detalhe = "FolderSecurity aplicado na concessao"
        }
    }
    catch {
        Write-Log "Falha ao aplicar Group '$Grupo' na concessao '$nomeConcessao': $($_.Exception.Message)" "ERROR"
        return [PSCustomObject]@{
            Tipo = "Concessao"
            Nome = $nomeConcessao
            Caminho = $caminhoConcessao
            Grupo = $Grupo
            Status = "ERRO"
            Detalhe = $_.Exception.Message
        }
    }
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

try {
    Write-Log "Inicio da validacao de login."
    Assert-Mta
    Importar-ModuloProjectWise
    Conectar-ProjectWise

    $datasourceAtual = Obter-DatasourceAtual
    if ([string]::IsNullOrWhiteSpace($datasourceAtual)) {
        throw "Login aparentemente concluido, mas Get-PWCurrentDatasource nao retornou datasource."
    }

    Write-Log "Datasource conectado: $datasourceAtual" "OK"

    $usuarioAtual = Obter-UsuarioAtual
    $nomeUsuarioAtual = Obter-ValorSeguroPropriedade -Objeto $usuarioAtual -PossiveisNomes @("Name", "UserName", "LoginName", "Email")
    if (-not [string]::IsNullOrWhiteSpace($nomeUsuarioAtual)) {
        Write-Log "Usuario atual: $nomeUsuarioAtual" "OK"
    }
    else {
        Write-Warn "Nao foi possivel identificar o usuario atual, mas o datasource esta conectado."
    }

    Write-Log "Validacao de login concluida com sucesso." "OK"

    $NomeGrupo = Resolver-GrupoPorTexto -Texto $NomeGrupo
    Write-Host ""
    Write-Host "Group validado:" -ForegroundColor Yellow
    Write-Host "Group   : $NomeGrupo"

    $concessaoSelecionada = Selecionar-ConcessaoEngenharia
    $nomeConcessaoSelecionada = Obter-RotuloPasta -Pasta $concessaoSelecionada
    Write-Log "Concessao selecionada: $nomeConcessaoSelecionada" "OK"

    Write-Host ""
    Write-Host "Concessao selecionada:" -ForegroundColor Yellow
    Write-Host "Concessao : $nomeConcessaoSelecionada"
    Write-Host "Caminho   : $(Obter-CaminhoPasta -Pasta $concessaoSelecionada)"
    Write-Host "ID        : $(Obter-IdPasta -Pasta $concessaoSelecionada)"

    $resultados = @()
    $ocorrenciasConcessao = @(Verificar-GrupoNoProjeto -Projeto $concessaoSelecionada -Grupo $NomeGrupo)
    Write-Host ""
    if ($ocorrenciasConcessao.Count -gt 0) {
        Write-Host "Verificacao de seguranca:" -ForegroundColor Yellow
        Write-Host "O Group '$NomeGrupo' ja aparece na seguranca da concessao. Ocorrencias: $($ocorrenciasConcessao.Count)"
        Write-Log "Group '$NomeGrupo' ja encontrado na seguranca da concessao '$nomeConcessaoSelecionada'. Ocorrencias=$($ocorrenciasConcessao.Count)" "WARN"

        $resultados += [PSCustomObject]@{
            Tipo = "Concessao"
            Concessao = $nomeConcessaoSelecionada
            Nome = $nomeConcessaoSelecionada
            Caminho = Obter-CaminhoPasta -Pasta $concessaoSelecionada
            Grupo = $NomeGrupo
            Status = "JA_EXISTE"
            Detalhe = "Group ja encontrado na seguranca da concessao. Ocorrencias=$($ocorrenciasConcessao.Count)"
        }
    }
    else {
        Write-Host "Verificacao de seguranca:" -ForegroundColor Yellow
        Write-Host "O Group '$NomeGrupo' ainda nao foi encontrado na seguranca da concessao."
        Write-Log "Group '$NomeGrupo' nao encontrado na seguranca da concessao '$nomeConcessaoSelecionada'." "OK"
        $resultados += @(Aplicar-GrupoNaConcessao -Concessao $concessaoSelecionada -Grupo $NomeGrupo)
    }

    $projetosConcessao = @(Obter-ProjetosDaConcessao -Concessao $concessaoSelecionada)
    Write-Host ""
    Write-Host "Projetos encontrados para validacao: $($projetosConcessao.Count)" -ForegroundColor Cyan

    foreach ($projeto in $projetosConcessao) {
        $registro = Criar-RegistroVerificacaoProjeto -Projeto $projeto -Grupo $NomeGrupo -NomeConcessao $nomeConcessaoSelecionada
        $resultados += $registro
        Write-Host ("{0} | {1} | {2}" -f $registro.Status, $registro.Nome, $registro.Caminho)
    }

    $resultados | Export-Csv -Path $ArquivoCsv -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Relatorio CSV: $ArquivoCsv" -ForegroundColor Green
    Write-Log "Relatorio CSV gerado: $ArquivoCsv" "OK"
}
catch {
    Write-Log "Erro na execucao: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    Encerrar-SessaoProjectWise
}
