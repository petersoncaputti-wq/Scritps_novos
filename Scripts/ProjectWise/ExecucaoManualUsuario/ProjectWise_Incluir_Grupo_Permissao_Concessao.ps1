<#
.SYNOPSIS
    Inclui um Group como permissao nos projetos de uma concessao no ProjectWise.

.DESCRIPTION
    O script conecta no ProjectWise, localiza a pasta Engenharia, permite escolher
    uma concessao e aplica um Group como objeto de seguranca nos projetos daquela
    concessao. Por padrao ele apenas simula; use -Executar para gravar.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Incluir_Grupo_Permissao_Concessao.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Incluir_Grupo_Permissao_Concessao.ps1" -ConcessaoNome "Ecovias Capixaba" -NomeGrupo "Engenharia GPR" -TipoSeguranca Ambos -TodosProjetos -Executar
#>

[CmdletBinding()]
param(
    [string]$ConcessaoNome = "",
    [string]$NomeGrupo = "",
    [string]$TextoGrupo = "",
    [ValidateSet("Folder", "Document", "Ambos")]
    [string]$TipoSeguranca = "Folder",
    [string[]]$PermissaoPasta = @("r"),
    [string[]]$PermissaoDocumento = @("r", "fr"),
    [int]$ProfundidadeProjetos = 0,
    [switch]$TodosProjetos,
    [switch]$IncluirSubpastas,
    [switch]$Executar,
    [switch]$NaoDesconectar
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArquivoLog = Join-Path $PastaLogs "PW_IncluirGrupoPermissaoConcessao_$TimeStamp.log"
$ArquivoCsv = Join-Path $PastaLogs "PW_IncluirGrupoPermissaoConcessao_$TimeStamp.csv"

$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")
$script:LoginPW = $null

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
            Write-Log "Modulo '$modulo' importado."
            return
        }
    }

    throw "Modulo ProjectWise/PWPS_DAB nao encontrado."
}

function Conectar-ProjectWise {
    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        $params = @($cmd.Parameters.Keys)
        try {
            Write-Log "Tentando conexao via $nomeCmdlet..."
            if ($params -contains "BentleyIMS") {
                $script:LoginPW = & $nomeCmdlet -BentleyIMS -ErrorAction Stop
                Write-Log "Conexao realizada via $nomeCmdlet -BentleyIMS."
                return
            }
            if ($params -contains "UseGui") {
                $script:LoginPW = & $nomeCmdlet -UseGui -ErrorAction Stop
                Write-Log "Conexao realizada via $nomeCmdlet -UseGui."
                return
            }

            $script:LoginPW = & $nomeCmdlet -ErrorAction Stop
            Write-Log "Conexao realizada via $nomeCmdlet."
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
                & $nomeCmdlet -Login $script:LoginPW -ErrorAction Stop
            }
            elseif ($params -contains "InputObject") {
                & $nomeCmdlet -InputObject $script:LoginPW -ErrorAction Stop
            }
            else {
                & $nomeCmdlet -ErrorAction Stop
            }

            Write-Log "Sessao ProjectWise encerrada via $nomeCmdlet."
            return
        }
        catch {
            Write-Warn "Falha ao encerrar sessao via $nomeCmdlet`: $($_.Exception.Message)"
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

function Ordenar-ItensPorNome {
    param([array]$Itens)
    return @($Itens | Sort-Object -Property @{ Expression = { (Obter-RotuloPasta -Pasta $_).ToLowerInvariant() } })
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)
    if ($params -contains "Root") {
        return @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $_ -ne $null })
    }
    return @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $_ -ne $null })
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
        return @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $_ -ne $null })
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

function Confirmar-SimNao {
    param(
        [string]$Mensagem,
        [bool]$Padrao = $false
    )

    $sufixo = if ($Padrao) { "S/n" } else { "s/N" }

    while ($true) {
        $resposta = (Read-Host "$Mensagem [$sufixo]").Trim().ToLowerInvariant()

        if ([string]::IsNullOrWhiteSpace($resposta)) {
            return $Padrao
        }

        if ($resposta -in @("s", "sim", "y", "yes")) {
            return $true
        }

        if ($resposta -in @("n", "nao", "não", "no")) {
            return $false
        }

        Write-Warn "Responda S ou N."
    }
}

function Ler-SelecaoProjetos {
    param([int]$Total)

    if ($TodosProjetos) {
        return @(1..$Total)
    }

    Write-Host ""
    Write-Host "Digite numeros separados por virgula, intervalos como 1-5, ou 'todos'." -ForegroundColor Cyan
    while ($true) {
        $entrada = (Read-Host "Projetos").Trim().ToLowerInvariant()
        if ($entrada -in @("todos", "tudo", "all")) {
            return @(1..$Total)
        }

        $indices = New-Object System.Collections.Generic.List[int]
        $partes = $entrada -split ","
        $valido = $true

        foreach ($parteOriginal in $partes) {
            $parte = $parteOriginal.Trim()
            if ([string]::IsNullOrWhiteSpace($parte)) {
                continue
            }

            if ($parte -match '^(\d+)-(\d+)$') {
                $inicio = [int]$Matches[1]
                $fim = [int]$Matches[2]
                if ($inicio -gt $fim) {
                    $valido = $false
                    break
                }
                foreach ($numero in $inicio..$fim) {
                    if ($numero -lt 1 -or $numero -gt $Total) {
                        $valido = $false
                        break
                    }
                    if (-not $indices.Contains($numero)) {
                        $indices.Add($numero)
                    }
                }
                continue
            }

            if ($parte -match '^\d+$') {
                $numero = [int]$parte
                if ($numero -lt 1 -or $numero -gt $Total) {
                    $valido = $false
                    break
                }
                if (-not $indices.Contains($numero)) {
                    $indices.Add($numero)
                }
                continue
            }

            $valido = $false
            break
        }

        if ($valido -and $indices.Count -gt 0) {
            return @($indices)
        }

        Write-Warn "Selecao invalida."
    }
}

function Obter-ProjetosDaPastaProjetos {
    param(
        [object]$PastaProjetos,
        [int]$ProfundidadeMaxima
    )

    $idPastaProjetos = Obter-IdPasta -Pasta $PastaProjetos
    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idPastaProjetos; Nivel = 0; Nome = "Projetos" })

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
                $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nivel = ($atual.Nivel + 1); Nome = (Obter-RotuloPasta -Pasta $filho) })
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

    return @(Ordenar-ItensPorNome -Itens @($mapa.Values))
}

function Obter-SubpastasRecursivas {
    param([object]$PastaRaiz)

    $resultado = @()
    $idRaiz = Obter-IdPasta -Pasta $PastaRaiz
    if ([string]::IsNullOrWhiteSpace($idRaiz)) {
        return @()
    }

    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idRaiz; Nome = (Obter-RotuloPasta -Pasta $PastaRaiz) })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        foreach ($filho in (Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId -Contexto $atual.Nome)) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) {
                continue
            }
            $resultado += $filho
            $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nome = (Obter-RotuloPasta -Pasta $filho) })
        }
    }

    return @($resultado)
}

function Obter-Concessao {
    param([array]$Concessoes)

    if (-not [string]::IsNullOrWhiteSpace($ConcessaoNome)) {
        $matches = @(
            $Concessoes | Where-Object {
                (Obter-RotuloPasta -Pasta $_) -like "*$ConcessaoNome*" -or
                (Obter-NomePasta -Pasta $_) -like "*$ConcessaoNome*"
            }
        )

        if ($matches.Count -eq 1) {
            return $matches[0]
        }
        if ($matches.Count -gt 1) {
            Write-Warn "Mais de uma concessao encontrada para '$ConcessaoNome'. Selecione pelo numero."
            $Concessoes = $matches
        }
        else {
            throw "Nenhuma concessao encontrada contendo '$ConcessaoNome'."
        }
    }

    Write-Host ""
    Write-Host "Concessoes disponiveis:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Concessoes.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Obter-RotuloPasta -Pasta $Concessoes[$i]) | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $Concessoes.Count
    return $Concessoes[$indice - 1]
}

function Obter-NomeGrupo {
    param([object]$Grupo)

    if ($null -eq $Grupo) {
        return ""
    }

    if ($Grupo -is [string]) {
        return $Grupo.Trim()
    }

    return Obter-ValorSeguroPropriedade -Objeto $Grupo -PossiveisNomes @("Name", "GroupName", "ObjectName", "MemberName")
}

function Obter-TodosGrupos {
    if (-not (Get-Command Get-PWGroups -ErrorAction SilentlyContinue)) {
        throw "Cmdlet Get-PWGroups nao encontrado; nao foi possivel buscar grupos."
    }

    $grupos = @(Get-PWGroups -ErrorAction Stop | Where-Object { $_ -ne $null })
    if ($grupos.Count -eq 0) {
        throw "Nenhum Group retornado pelo ProjectWise."
    }

    return @($grupos | Sort-Object @{ Expression = { (Obter-NomeGrupo -Grupo $_).ToLowerInvariant() } })
}

function Confirmar-GrupoSelecionado {
    param(
        [string]$NomeGrupoResolvido,
        [string]$TextoBusca
    )

    Write-Host ""
    Write-Host "Group encontrado/selecionado:" -ForegroundColor Yellow
    Write-Host "Busca : $TextoBusca"
    Write-Host "Group : $NomeGrupoResolvido"

    if (-not (Confirmar-SimNao -Mensagem "Confirma que este e o Group correto?" -Padrao $false)) {
        throw "Operacao cancelada pelo usuario na confirmacao do Group."
    }

    Write-Log "Group confirmado pelo usuario: $NomeGrupoResolvido"
    return $NomeGrupoResolvido
}

function Confirmar-ExecucaoReal {
    param(
        [string]$Concessao,
        [string]$Grupo,
        [int]$QuantidadeProjetos
    )

    Write-Host ""
    Write-Host "ATENCAO: execucao real solicitada." -ForegroundColor Red
    Write-Host "Esta rotina vai alterar permissoes no ProjectWise."
    Write-Host "Concessao : $Concessao"
    Write-Host "Group     : $Grupo"
    Write-Host "Projetos  : $QuantidadeProjetos"
    Write-Host ""

    $confirmacao = Read-Host "Para confirmar a execucao real, digite EXECUTAR"
    if ($confirmacao -cne "EXECUTAR") {
        throw "Operacao cancelada. Confirmacao final de execucao real nao foi informada."
    }

    Write-Log "Execucao real confirmada pelo usuario."
}

function Resolver-GrupoPorTexto {
    param([string]$Texto)

    if (-not (Get-Command Get-PWGroups -ErrorAction SilentlyContinue)) {
        if ([string]::IsNullOrWhiteSpace($Texto)) {
            throw "NomeGrupo ou TextoGrupo nao informado."
        }
        Write-Warn "Cmdlet Get-PWGroups nao encontrado; usando texto informado sem validar/buscar."
        return $Texto.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        $Texto = (Read-Host "Texto para buscar o Group").Trim()
    }

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        throw "Texto de busca do Group nao informado."
    }

    $todosGrupos = Obter-TodosGrupos
    $exatos = @(
        $todosGrupos | Where-Object {
            (Obter-NomeGrupo -Grupo $_) -ieq $Texto
        }
    )

    if ($exatos.Count -eq 1) {
        $nomeExato = Obter-NomeGrupo -Grupo $exatos[0]
        Write-Log "Group encontrado por nome exato: $nomeExato"
        return Confirmar-GrupoSelecionado -NomeGrupoResolvido $nomeExato -TextoBusca $Texto
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
        throw "Nenhum Group encontrado contendo o texto '$Texto'."
    }

    if ($encontrados.Count -eq 1) {
        $nomeEncontrado = Obter-NomeGrupo -Grupo $encontrados[0]
        Write-Log "Group encontrado por busca textual '$Texto': $nomeEncontrado"
        return Confirmar-GrupoSelecionado -NomeGrupoResolvido $nomeEncontrado -TextoBusca $Texto
    }

    Write-Host ""
    Write-Host "Groups encontrados para '$Texto':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $encontrados.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Obter-NomeGrupo -Grupo $encontrados[$i]) | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero do Group" -Minimo 1 -Maximo $encontrados.Count
    $nomeSelecionado = Obter-NomeGrupo -Grupo $encontrados[$indice - 1]
    Write-Log "Group selecionado por busca textual '$Texto': $nomeSelecionado"
    return Confirmar-GrupoSelecionado -NomeGrupoResolvido $nomeSelecionado -TextoBusca $Texto
}

function Aplicar-PermissaoGrupo {
    param(
        [object]$Pasta,
        [string]$Grupo,
        [string]$ProjetoNome
    )

    $pastaRotulo = Obter-RotuloPasta -Pasta $Pasta
    $pastaCaminho = Obter-CaminhoPasta -Pasta $Pasta
    if ([string]::IsNullOrWhiteSpace($pastaCaminho)) {
        $pastaCaminho = $pastaRotulo
    }

    $resultados = @()

    if ($TipoSeguranca -in @("Folder", "Ambos")) {
        try {
            if ($Executar) {
                Update-PWFolderSecurity -InputFolder $Pasta -FolderSecurity -MemberType Group -MemberName $Grupo -MemberAccess $PermissaoPasta -ErrorAction Stop | Out-Null
                $status = "SUCESSO"
                $detalhe = "FolderSecurity aplicado"
            }
            else {
                $status = "SIMULADO"
                $detalhe = "FolderSecurity seria aplicado"
            }
        }
        catch {
            $status = "ERRO"
            $detalhe = $_.Exception.Message
        }

        $resultados += [PSCustomObject]@{
            Projeto = $ProjetoNome
            Pasta = $pastaCaminho
            Grupo = $Grupo
            TipoSeguranca = "Folder"
            Permissoes = ($PermissaoPasta -join ",")
            Status = $status
            Detalhe = $detalhe
        }
    }

    if ($TipoSeguranca -in @("Document", "Ambos")) {
        try {
            if ($Executar) {
                Update-PWFolderSecurity -InputFolder $Pasta -DocumentSecurity -MemberType Group -MemberName $Grupo -MemberAccess $PermissaoDocumento -ErrorAction Stop | Out-Null
                $status = "SUCESSO"
                $detalhe = "DocumentSecurity aplicado"
            }
            else {
                $status = "SIMULADO"
                $detalhe = "DocumentSecurity seria aplicado"
            }
        }
        catch {
            $status = "ERRO"
            $detalhe = $_.Exception.Message
        }

        $resultados += [PSCustomObject]@{
            Projeto = $ProjetoNome
            Pasta = $pastaCaminho
            Grupo = $Grupo
            TipoSeguranca = "Document"
            Permissoes = ($PermissaoDocumento -join ",")
            Status = $status
            Detalhe = $detalhe
        }
    }

    return $resultados
}

try {
    Assert-Mta
    Importar-ModuloProjectWise
    Conectar-ProjectWise

    $textoBuscaGrupo = $NomeGrupo
    if (-not [string]::IsNullOrWhiteSpace($TextoGrupo)) {
        $textoBuscaGrupo = $TextoGrupo
    }
    $NomeGrupo = Resolver-GrupoPorTexto -Texto $textoBuscaGrupo

    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia"
    $concessoes = Ordenar-ItensPorNome -Itens (Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $pastaEngenharia) -Contexto "Engenharia")
    if ($concessoes.Count -eq 0) {
        throw "Nenhuma concessao encontrada."
    }

    $concessao = Obter-Concessao -Concessoes $concessoes
    $nomeConcessao = Obter-RotuloPasta -Pasta $concessao

    Write-Log "Concessao selecionada: $nomeConcessao"
    Write-Log "Carregando projetos da concessao..."
    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $concessao) -Contexto $nomeConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"
    $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -ProfundidadeMaxima $ProfundidadeProjetos
    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto encontrado em '$nomeConcessao'."
    }

    Write-Host ""
    Write-Host "Projetos encontrados em '$nomeConcessao':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $projetos.Count; $i++) {
        "{0:000}. {1}" -f ($i + 1), (Obter-RotuloPasta -Pasta $projetos[$i]) | Write-Host
    }

    $indices = Ler-SelecaoProjetos -Total $projetos.Count
    $projetosSelecionados = @(
        foreach ($indice in $indices) {
            $projetos[$indice - 1]
        }
    )

    $modo = if ($Executar) { "EXECUCAO REAL" } else { "SIMULACAO" }
    Write-Host ""
    Write-Host "Resumo da operacao ($modo)" -ForegroundColor Yellow
    Write-Host "Concessao      : $nomeConcessao"
    Write-Host "Group          : $NomeGrupo"
    Write-Host "Tipo seguranca : $TipoSeguranca"
    Write-Host "Permissao pasta: $($PermissaoPasta -join ',')"
    Write-Host "Permissao doc. : $($PermissaoDocumento -join ',')"
    Write-Host "Projetos       : $($projetosSelecionados.Count)"
    Write-Host "Subpastas      : $([bool]$IncluirSubpastas)"

    if (-not $Executar) {
        Write-Warn "Modo simulacao ativo. Para gravar no ProjectWise, execute novamente com -Executar."
    }
    else {
        Confirmar-ExecucaoReal -Concessao $nomeConcessao -Grupo $NomeGrupo -QuantidadeProjetos $projetosSelecionados.Count
    }

    $resultados = @()
    foreach ($projeto in $projetosSelecionados) {
        $nomeProjeto = Obter-RotuloPasta -Pasta $projeto
        Write-Log "Processando projeto: $nomeProjeto"

        $pastasParaAplicar = @($projeto)
        if ($IncluirSubpastas) {
            $pastasParaAplicar += @(Obter-SubpastasRecursivas -PastaRaiz $projeto)
        }

        foreach ($pasta in $pastasParaAplicar) {
            $resultados += @(Aplicar-PermissaoGrupo -Pasta $pasta -Grupo $NomeGrupo -ProjetoNome $nomeProjeto)
        }
    }

    $resultados | Export-Csv -Path $ArquivoCsv -NoTypeInformation -Encoding UTF8

    $sucesso = @($resultados | Where-Object { $_.Status -eq "SUCESSO" }).Count
    $simulado = @($resultados | Where-Object { $_.Status -eq "SIMULADO" }).Count
    $erros = @($resultados | Where-Object { $_.Status -eq "ERRO" }).Count

    Write-Host ""
    Write-Host "[OK] Rotina concluida." -ForegroundColor Green
    Write-Host "Sucesso : $sucesso"
    Write-Host "Simulado: $simulado"
    Write-Host "Erros   : $erros"
    Write-Host "Log     : $ArquivoLog"
    Write-Host "CSV     : $ArquivoCsv"

    if ($erros -gt 0) {
        Write-Warn "Houve erro(s). Consulte o CSV para ver pasta, tipo de seguranca e detalhe."
    }
}
finally {
    Encerrar-SessaoProjectWise
}
