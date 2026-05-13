<#
.SYNOPSIS
    Teste verboso: lista concessoes, carrega projetos da concessao escolhida e mostra progresso.

.DESCRIPTION
    Versao de diagnostico com mensagens frequentes para evitar a sensacao de travamento.
    Por padrao usa ProfundidadeProjetos 0, ou seja, somente projetos diretamente dentro
    da pasta Projetos.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\teste_fluxo_concessao_depois_projetos_v2_verbose.ps1"
#>

[CmdletBinding()]
param(
    [int]$ProfundidadeProjetos = 0,
    [int]$LimiteProjetos = 0
)

$ErrorActionPreference = "Stop"
$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")
$script:LoginPW = $null

function Write-Step {
    param([string]$Mensagem)
    Write-Host ("{0} | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Mensagem)
    Write-Host ("{0} | AVISO | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Yellow
}

function Assert-Mta {
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartment -ne [System.Threading.ApartmentState]::MTA) {
        throw "Execute usando powershell.exe -MTA. Estado atual: $apartment"
    }
}

function Importar-ModuloProjectWise {
    foreach ($modulo in @("pwps_dab", "PWPS_DAB", "Bentley.PowerShell.ProjectWise", "ProjectWisePowerShell")) {
        if (Get-Module -Name $modulo) { return }
        if (Get-Module -ListAvailable -Name $modulo) {
            Import-Module $modulo -ErrorAction Stop
            return
        }
    }
    throw "Modulo ProjectWise/PWPS_DAB nao encontrado."
}

function Conectar-ProjectWise {
    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }

        $params = @($cmd.Parameters.Keys)
        try {
            Write-Step "Tentando conexao via $nomeCmdlet..."
            if ($params -contains "BentleyIMS") {
                $script:LoginPW = & $nomeCmdlet -BentleyIMS -ErrorAction Stop
                return
            }
            if ($params -contains "UseGui") {
                $script:LoginPW = & $nomeCmdlet -UseGui -ErrorAction Stop
                return
            }
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
    if ($null -eq $script:LoginPW) { return }
    foreach ($nomeCmdlet in @("Undo-PWLogin", "Close-PWConnection", "Remove-PWLogin")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
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
            return
        }
        catch {
        }
    }
}

function Obter-ValorSeguroPropriedade {
    param([object]$Objeto, [string[]]$PossiveisNomes)
    if ($null -eq $Objeto) { return "" }
    foreach ($nome in $PossiveisNomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = $propriedade.Value.ToString().Trim()
            if ($valor -ne "") { return $valor }
        }
    }
    return ""
}

function Obter-IdPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID")
}

function Obter-GuidPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectGUID", "ProjectGuid", "FolderGUID", "FolderGuid", "GUID", "Guid")
}

function Obter-NomePasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Name", "FolderName", "ProjectName", "ObjectName")
}

function Obter-RotuloPasta {
    param([object]$Pasta)
    $descricao = Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Description", "Descricao", "ProjectDescription", "FolderDescription")
    if (-not [string]::IsNullOrWhiteSpace($descricao)) { return $descricao }
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
        if (-not [string]::IsNullOrWhiteSpace($Contexto)) {
            Write-Step "Listando filhos: $Contexto"
        }
        $inicio = Get-Date
        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $_ -ne $null })
        $segundos = [Math]::Round(((Get-Date) - $inicio).TotalSeconds, 1)
        if (-not [string]::IsNullOrWhiteSpace($Contexto)) {
            Write-Step "Encontrado(s) $($pastas.Count) filho(s) em $segundos s: $Contexto"
        }
        return $pastas
    }
    catch {
        Write-Warn "Falha ao listar filhos de $Contexto ($FolderId): $($_.Exception.Message)"
        return @()
    }
}

function Localizar-PastaPorPossiveisNomes {
    param([array]$Pastas, [string[]]$NomesPossiveis, [string]$Descricao)
    foreach ($nomeEsperado in $NomesPossiveis) {
        $encontrada = @(
            $Pastas | Where-Object {
                (Obter-NomePasta -Pasta $_) -ieq $nomeEsperado -or
                (Obter-RotuloPasta -Pasta $_) -ieq $nomeEsperado
            }
        )
        if ($encontrada.Count -gt 0) { return $encontrada[0] }
    }
    $nomes = @($Pastas | ForEach-Object { Obter-RotuloPasta -Pasta $_ }) -join ", "
    throw "Nao encontrei $Descricao. Pastas disponiveis: $nomes"
}

function Ler-Numero {
    param([string]$Mensagem, [int]$Minimo, [int]$Maximo)
    while ($true) {
        $entrada = Read-Host $Mensagem
        if ($entrada -match '^\d+$') {
            $numero = [int]$entrada
            if ($numero -ge $Minimo -and $numero -le $Maximo) { return $numero }
        }
        Write-Warn "Informe um numero entre $Minimo e $Maximo."
    }
}

function Mostrar-Concessoes {
    param([array]$Concessoes)
    Write-Host ""
    Write-Host "Concessoes disponiveis:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Concessoes.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Obter-RotuloPasta -Pasta $Concessoes[$i]) | Write-Host
    }
}

function Mostrar-Projetos {
    param([array]$Projetos)
    Write-Host ""
    Write-Host "Projetos da concessao escolhida:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Projetos.Count; $i++) {
        $projeto = $Projetos[$i]
        "{0:000}. {1} | ID: {2} | GUID: {3}" -f ($i + 1), (Obter-RotuloPasta -Pasta $projeto), (Obter-IdPasta -Pasta $projeto), (Obter-GuidPasta -Pasta $projeto) | Write-Host
    }
}

function Obter-ProjetosDaPastaProjetos {
    param([object]$PastaProjetos, [int]$ProfundidadeMaxima, [int]$Limite)
    $idPastaProjetos = Obter-IdPasta -Pasta $PastaProjetos
    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idPastaProjetos; Nivel = 0; Nome = "Projetos" })
    $pastasProcessadas = 0

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $pastasProcessadas++
        Write-Step "Processando nivel $($atual.Nivel): $($atual.Nome) | fila=$($fila.Count) | projetos=$($resultado.Count)"
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId -Contexto $atual.Nome

        foreach ($filho in $filhos) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) { continue }

            $resultado += $filho
            if ($resultado.Count % 25 -eq 0) {
                Write-Step "$($resultado.Count) projeto(s)/pasta(s) coletado(s) ate agora..."
            }

            if ($Limite -gt 0 -and $resultado.Count -ge $Limite) {
                Write-Warn "Limite de $Limite projeto(s) atingido; parando coleta para teste."
                return @(Ordenar-ItensPorNome -Itens $resultado)
            }

            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $nomeFilho = Obter-RotuloPasta -Pasta $filho
                $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nivel = ($atual.Nivel + 1); Nome = $nomeFilho })
            }
        }
    }

    Write-Step "Coleta concluida. Pastas processadas=$pastasProcessadas | itens coletados=$($resultado.Count)"
    return @(Ordenar-ItensPorNome -Itens $resultado)
}

try {
    Assert-Mta
    Write-Step "Importando modulo ProjectWise..."
    Importar-ModuloProjectWise
    Conectar-ProjectWise
    Write-Step "Conexao ProjectWise OK."

    Write-Step "Listando pastas raiz..."
    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia"
    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenharia

    $concessoes = Ordenar-ItensPorNome -Itens (Get-PWChildFoldersOrEmpty -FolderId $idEngenharia -Contexto "Engenharia")
    if ($concessoes.Count -eq 0) { throw "Nenhuma concessao encontrada." }

    Mostrar-Concessoes -Concessoes $concessoes
    $indiceConcessao = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $concessoes.Count
    $concessao = $concessoes[$indiceConcessao - 1]
    $nomeConcessao = Obter-RotuloPasta -Pasta $concessao
    $idConcessao = Obter-IdPasta -Pasta $concessao

    Write-Step "Listando pastas internas da concessao '$nomeConcessao'..."
    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId $idConcessao -Contexto $nomeConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"
    Write-Step "Pasta de projetos localizada: $(Obter-RotuloPasta -Pasta $pastaProjetos) | ID=$(Obter-IdPasta -Pasta $pastaProjetos)"

    Write-Step "Carregando projetos com profundidade $ProfundidadeProjetos..."
    $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -ProfundidadeMaxima $ProfundidadeProjetos -Limite $LimiteProjetos
    if ($projetos.Count -eq 0) { throw "Nenhum projeto encontrado em '$nomeConcessao'." }

    Mostrar-Projetos -Projetos $projetos
    $indiceProjeto = Ler-Numero -Mensagem "Numero do projeto" -Minimo 1 -Maximo $projetos.Count
    $projeto = $projetos[$indiceProjeto - 1]
    $idProjeto = Obter-IdPasta -Pasta $projeto

    Write-Step "Projeto selecionado: $(Obter-RotuloPasta -Pasta $projeto) | ID=$idProjeto | GUID=$(Obter-GuidPasta -Pasta $projeto)"
    Write-Step "Rodando diagnostico no projeto selecionado..."
    & (Join-Path $PSScriptRoot "diagnostico_pw_projectwise_project_v2.ps1") -FolderId $idProjeto -NaoDesconectar
}
finally {
    Encerrar-SessaoProjectWise
}
