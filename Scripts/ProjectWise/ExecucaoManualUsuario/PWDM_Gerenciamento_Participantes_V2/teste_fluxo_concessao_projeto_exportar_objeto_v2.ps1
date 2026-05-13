<#
.SYNOPSIS
    Lista concessoes/projetos e exporta o objeto do projeto selecionado sem rebuscar por ID.

.DESCRIPTION
    Este teste evita o problema observado no Get-PWFolders -FolderID. Ele usa o objeto
    retornado por Get-PWFoldersImmediateChildren durante a listagem e exporta suas
    propriedades imediatamente.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\teste_fluxo_concessao_projeto_exportar_objeto_v2.ps1"
#>

[CmdletBinding()]
param(
    [int]$ProfundidadeProjetos = 0,
    [int]$LimiteProjetos = 0
)

$ErrorActionPreference = "Stop"
$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")
$script:LoginPW = $null

function Write-Step {
    param([string]$Mensagem)
    Write-Host ("{0} | {1}" -f (Get-Date -Format "HH:mm:ss"), $Mensagem) -ForegroundColor Cyan
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
            Write-Host "Falha em $nomeCmdlet`: $($_.Exception.Message)" -ForegroundColor Yellow
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
    param([string]$FolderId, [string]$Contexto = "")
    try {
        if ($Contexto) { Write-Step "Listando filhos: $Contexto" }
        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $_ -ne $null })
        if ($Contexto) { Write-Step "Encontrado(s) $($pastas.Count) filho(s): $Contexto" }
        return $pastas
    }
    catch {
        Write-Host "Falha ao listar filhos de $Contexto ($FolderId): $($_.Exception.Message)" -ForegroundColor Yellow
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
        Write-Host "Informe um numero entre $Minimo e $Maximo." -ForegroundColor Yellow
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

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId -Contexto $atual.Nome
        foreach ($filho in $filhos) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) { continue }
            $resultado += $filho

            if ($Limite -gt 0 -and $resultado.Count -ge $Limite) {
                return @(Ordenar-ItensPorNome -Itens $resultado)
            }

            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nivel = ($atual.Nivel + 1); Nome = (Obter-RotuloPasta -Pasta $filho) })
            }
        }
    }

    return @(Ordenar-ItensPorNome -Itens $resultado)
}

function Converter-ValorSeguro {
    param([object]$Valor, [int]$Profundidade = 0)
    if ($null -eq $Valor) { return $null }
    if ($Profundidade -ge 4) { return $Valor.ToString() }
    if ($Valor -is [string] -or $Valor.GetType().IsPrimitive -or $Valor -is [decimal]) { return $Valor }
    if ($Valor -is [System.Collections.IDictionary]) {
        $resultado = [ordered]@{}
        foreach ($chave in $Valor.Keys) {
            $resultado[$chave.ToString()] = Converter-ValorSeguro -Valor $Valor[$chave] -Profundidade ($Profundidade + 1)
        }
        return $resultado
    }
    if ($Valor -is [System.Collections.IEnumerable] -and -not ($Valor -is [string])) {
        $itens = @()
        $contador = 0
        foreach ($item in $Valor) {
            if ($contador -ge 50) {
                $itens += "[lista truncada]"
                break
            }
            $itens += Converter-ValorSeguro -Valor $item -Profundidade ($Profundidade + 1)
            $contador++
        }
        return $itens
    }
    $objeto = [ordered]@{}
    foreach ($prop in $Valor.PSObject.Properties) {
        try {
            $objeto[$prop.Name] = Converter-ValorSeguro -Valor $prop.Value -Profundidade ($Profundidade + 1)
        }
        catch {
            $objeto[$prop.Name] = "[erro ao ler: $($_.Exception.Message)]"
        }
    }
    if ($objeto.Count -eq 0) { return $Valor.ToString() }
    return $objeto
}

function Encontrar-Guids {
    param([object]$Valor)
    $texto = ($Valor | ConvertTo-Json -Depth 30 -Compress)
    $regex = [regex]"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
    return @($regex.Matches($texto) | ForEach-Object { $_.Value.ToLowerInvariant() } | Select-Object -Unique)
}

try {
    Assert-Mta
    Write-Step "Importando modulo ProjectWise..."
    Importar-ModuloProjectWise
    Conectar-ProjectWise
    Write-Step "Conexao ProjectWise OK."

    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis @("Engenharia", "Engineering") -Descricao "pasta de engenharia"
    $concessoes = Ordenar-ItensPorNome -Itens (Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $pastaEngenharia) -Contexto "Engenharia")
    Mostrar-Concessoes -Concessoes $concessoes
    $indiceConcessao = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $concessoes.Count
    $concessao = $concessoes[$indiceConcessao - 1]

    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $concessao) -Contexto (Obter-RotuloPasta -Pasta $concessao)
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos"
    $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -ProfundidadeMaxima $ProfundidadeProjetos -Limite $LimiteProjetos
    Mostrar-Projetos -Projetos $projetos
    $indiceProjeto = Ler-Numero -Mensagem "Numero do projeto" -Minimo 1 -Maximo $projetos.Count
    $projeto = $projetos[$indiceProjeto - 1]

    $propriedades = Converter-ValorSeguro -Valor $projeto
    $guids = Encontrar-Guids -Valor $propriedades
    $guidPrincipal = Obter-GuidPasta -Pasta $projeto
    $urlsPossiveis = @()
    if (-not [string]::IsNullOrWhiteSpace($guidPrincipal)) {
        $urlsPossiveis += "https://pwdm.bentley.com/$guidPrincipal/ProjectSettings/$guidPrincipal/View#PARTICIPANTS"
    }
    foreach ($guid in $guids) {
        $url = "https://pwdm.bentley.com/$guid/ProjectSettings/$guid/View#PARTICIPANTS"
        if ($urlsPossiveis -notcontains $url) { $urlsPossiveis += $url }
    }

    $saida = Join-Path $PastaLogs ("pw_projectwise_project_objeto_selecionado_{0}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $dados = [ordered]@{
        timestamp = (Get-Date).ToString("s")
        concessao = [ordered]@{
            nome = Obter-RotuloPasta -Pasta $concessao
            id = Obter-IdPasta -Pasta $concessao
            guid = Obter-GuidPasta -Pasta $concessao
        }
        projeto = [ordered]@{
            nome = Obter-RotuloPasta -Pasta $projeto
            id = Obter-IdPasta -Pasta $projeto
            guid = $guidPrincipal
        }
        guidsEncontrados = $guids
        urlsPwdmPossiveis = $urlsPossiveis
        propriedadesProjeto = $propriedades
    }
    [System.IO.File]::WriteAllText($saida, ($dados | ConvertTo-Json -Depth 30), [System.Text.Encoding]::UTF8)

    Write-Host ""
    Write-Host "[OK] Objeto do projeto exportado." -ForegroundColor Green
    Write-Host "JSON: $saida"
    if ($urlsPossiveis.Count -gt 0) {
        Write-Host ""
        Write-Host "URLs PWDM possiveis:" -ForegroundColor Cyan
        $urlsPossiveis | ForEach-Object { Write-Host "- $_" }
    }
}
finally {
    Encerrar-SessaoProjectWise
}
