<#
.SYNOPSIS
    Seleciona projetos no ProjectWise para gerenciamento de acessos no ProjectWise Web.

.DESCRIPTION
    Script somente leitura. Ele conecta no ProjectWise, lista concessoes, carrega
    somente os projetos da concessao escolhida e exporta um JSON com os IDs do
    ProjectWise Project/PW Web usados pelo gerenciamento de membros RBAC.

    O campo principal e ConnectedProjectId.Guid. Quando ele existe, o script monta:
    - connectSpaceId
    - projectId

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\pw_selecionar_projetos_para_pwdm_v2.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\pw_selecionar_projetos_para_pwdm_v2.ps1" -ProfundidadeProjetos 0
#>

[CmdletBinding()]
param(
    [string]$Saida,
    [int]$ProfundidadeProjetos = 0,
    [switch]$TodosProjetos,
    [string]$PlanilhaProjetos
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

if ([string]::IsNullOrWhiteSpace($Saida)) {
    $Saida = Join-Path $PastaLogs ("pwdm_projetos_selecionados_pw_{0}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}

$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")
$GuidZero = "00000000-0000-0000-0000-000000000000"
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
    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

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
    if ($null -eq $script:LoginPW) {
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

function Obter-GuidPasta {
    param([object]$Pasta)
    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectGUIDString", "ProjectGUID", "ProjectGuid", "FolderGUID", "FolderGuid", "GUID", "Guid")
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
            Write-Step "Listando filhos: $Contexto"
        }
        $inicio = Get-Date
        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop | Where-Object { $_ -ne $null })
        if ($Contexto) {
            $segundos = [Math]::Round(((Get-Date) - $inicio).TotalSeconds, 1)
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

function Remover-AcentosTexto {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return ""
    }

    try {
        $normalizado = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
        $builder = New-Object System.Text.StringBuilder
        foreach ($ch in $normalizado.ToCharArray()) {
            $categoria = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
            if ($categoria -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$builder.Append($ch)
            }
        }
        return $builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)
    }
    catch {
        return $Texto
    }
}

function Normalizar-TextoComparacaoProjeto {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return ""
    }

    $valor = Remover-AcentosTexto $Texto
    $valor = $valor.ToLowerInvariant()
    $valor = $valor -replace '[^\p{L}\p{Nd}]+', ' '
    $valor = $valor -replace '\s+', ' '
    return $valor.Trim()
}

function Obter-ValorCampoLinha {
    param(
        [object]$Linha,
        [string[]]$Nomes
    )

    foreach ($nome in $Nomes) {
        $prop = $Linha.PSObject.Properties[$nome]
        if ($prop -and $null -ne $prop.Value) {
            $valor = $prop.Value.ToString().Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
}

function Importar-NomesProjetosPlanilha {
    param([string]$CaminhoArquivo)

    if ([string]::IsNullOrWhiteSpace($CaminhoArquivo)) {
        throw "Nenhum caminho de planilha foi informado."
    }

    $CaminhoArquivo = $CaminhoArquivo.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $CaminhoArquivo)) {
        throw "Arquivo nao encontrado: $CaminhoArquivo"
    }

    $ext = [System.IO.Path]::GetExtension($CaminhoArquivo).ToLowerInvariant()
    if ($ext -in @(".xlsx", ".xlsm")) {
        Import-Module ImportExcel -ErrorAction Stop
        $dados = @(Import-Excel -Path $CaminhoArquivo)
    }
    elseif ($ext -eq ".csv") {
        $primeiraLinha = Get-Content -LiteralPath $CaminhoArquivo -TotalCount 1
        $qtdPontoVirgula = ([regex]::Matches($primeiraLinha, ';')).Count
        $qtdVirgula = ([regex]::Matches($primeiraLinha, ',')).Count
        $delimitador = if ($qtdPontoVirgula -gt $qtdVirgula) { ';' } else { ',' }
        $dados = @(Import-Csv -LiteralPath $CaminhoArquivo -Delimiter $delimitador)
    }
    else {
        throw "Extensao nao suportada: $ext. Use .xlsx, .xlsm ou .csv."
    }

    if (-not $dados -or $dados.Count -eq 0) {
        throw "A planilha selecionada nao possui projetos."
    }

    $projetos = New-Object System.Collections.Generic.List[string]
    $unicos = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $duplicados = New-Object System.Collections.Generic.List[string]
    $linhasSemProjeto = New-Object System.Collections.Generic.List[int]
    $numeroLinha = 2

    foreach ($linha in $dados) {
        $valor = Obter-ValorCampoLinha -Linha $linha -Nomes @("Projeto", "Projetos", "Nome Projeto", "Nome do Projeto", "Project", "ProjectName")
        if ([string]::IsNullOrWhiteSpace($valor)) {
            $linhasSemProjeto.Add($numeroLinha)
            $numeroLinha++
            continue
        }

        $nomeProjeto = $valor.Trim()
        if ($unicos.Add($nomeProjeto)) {
            $projetos.Add($nomeProjeto)
        }
        else {
            $duplicados.Add($nomeProjeto)
        }

        $numeroLinha++
    }

    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto foi encontrado. Use uma coluna chamada Projeto."
    }

    return [PSCustomObject]@{
        Projetos             = @($projetos.ToArray())
        LinhasPlanilha       = $dados.Count
        ProjetosUnicos       = $projetos.Count
        DuplicadosIgnorados  = $duplicados.Count
        ProjetosDuplicados   = @($duplicados | Select-Object -Unique)
        LinhasSemProjeto     = @($linhasSemProjeto.ToArray())
    }
}

function Obter-ChavesComparacaoProjeto {
    param([object]$Projeto)

    $valores = @(
        (Obter-RotuloPasta -Pasta $Projeto),
        (Obter-NomePasta -Pasta $Projeto),
        (Obter-DescricaoPasta -Pasta $Projeto),
        (Obter-IdPasta -Pasta $Projeto),
        (Obter-GuidPasta -Pasta $Projeto),
        (Obter-ConnectedProjectId -Projeto $Projeto),
        (Obter-ProjectWiseWebName -Projeto $Projeto),
        (Obter-ProjectWiseWebNumber -Projeto $Projeto)
    )

    return @(
        $valores |
            ForEach-Object { Normalizar-TextoComparacaoProjeto $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
}

function Localizar-ProjetoPorNomePlanilha {
    param(
        [string]$NomeProjetoPlanilha,
        [array]$Projetos
    )

    $alvo = Normalizar-TextoComparacaoProjeto $NomeProjetoPlanilha
    if ([string]::IsNullOrWhiteSpace($alvo)) {
        return [PSCustomObject]@{ Encontrado = $false; Indice = $null; Motivo = "Nome vazio"; Sugestoes = @() }
    }

    for ($i = 0; $i -lt $Projetos.Count; $i++) {
        if (@(Obter-ChavesComparacaoProjeto -Projeto $Projetos[$i]) -contains $alvo) {
            return [PSCustomObject]@{ Encontrado = $true; Indice = $i; Motivo = "Exato"; Sugestoes = @() }
        }
    }

    for ($i = 0; $i -lt $Projetos.Count; $i++) {
        foreach ($chave in @(Obter-ChavesComparacaoProjeto -Projeto $Projetos[$i])) {
            if ($chave.Contains($alvo) -or $alvo.Contains($chave)) {
                return [PSCustomObject]@{ Encontrado = $true; Indice = $i; Motivo = "Parcial"; Sugestoes = @() }
            }
        }
    }

    $termosAlvo = @($alvo -split '\s+' | Where-Object { $_.Length -ge 3 })
    $sugestoesCalculadas = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Projetos.Count; $i++) {
        $melhorPontuacao = 0
        foreach ($chave in @(Obter-ChavesComparacaoProjeto -Projeto $Projetos[$i])) {
            $pontuacao = 0
            foreach ($termo in $termosAlvo) {
                if ($chave.Contains($termo)) {
                    $pontuacao++
                }
            }
            if ($pontuacao -gt $melhorPontuacao) {
                $melhorPontuacao = $pontuacao
            }
        }
        if ($melhorPontuacao -gt 0) {
            $sugestoesCalculadas.Add([PSCustomObject]@{
                Nome      = Obter-RotuloPasta -Pasta $Projetos[$i]
                Pontuacao = $melhorPontuacao
            })
        }
    }
    $sugestoes = @($sugestoesCalculadas.ToArray() | Sort-Object Pontuacao -Descending | Select-Object -First 3)

    return [PSCustomObject]@{
        Encontrado = $false
        Indice     = $null
        Motivo     = "Nao encontrado"
        Sugestoes  = @($sugestoes | ForEach-Object { $_.Nome })
    }
}

function Ler-CaminhoPlanilhaProjetos {
    if (-not [string]::IsNullOrWhiteSpace($PlanilhaProjetos)) {
        return $PlanilhaProjetos
    }

    $scriptDialog = @'
Add-Type -AssemblyName System.Windows.Forms
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
$dialog.Title = "Selecione a planilha com a lista de projetos"
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.FileName
}
'@

    try {
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptDialog)
        $encoded = [Convert]::ToBase64String($bytes)
        $saida = @(powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded)
        $caminhoSelecionado = @($saida | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1)
        if ($caminhoSelecionado.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($caminhoSelecionado[0])) {
            return [string]$caminhoSelecionado[0]
        }
    }
    catch {
        Write-Warn "Nao foi possivel abrir a janela de selecao: $($_.Exception.Message)"
    }

    Write-Host ""
    Write-Host "Nenhuma planilha foi selecionada pela janela." -ForegroundColor Yellow
    Write-Host "Informe o caminho da planilha com a coluna Projeto." -ForegroundColor Cyan
    Write-Host "Dica: voce pode arrastar o arquivo para esta janela e pressionar ENTER." -ForegroundColor Cyan
    return Read-Host "Caminho da planilha"
}

function Ler-SelecaoProjetosPorPlanilha {
    param([array]$Projetos)

    $caminho = Ler-CaminhoPlanilhaProjetos
    $importacao = Importar-NomesProjetosPlanilha -CaminhoArquivo $caminho

    Write-Step "Linhas na planilha de projetos: $($importacao.LinhasPlanilha)"
    Write-Step "Projetos unicos importados: $($importacao.ProjetosUnicos) | Duplicados ignorados: $($importacao.DuplicadosIgnorados)"
    if (@($importacao.LinhasSemProjeto).Count -gt 0) {
        Write-Warn "Linhas sem projeto ignoradas: $(@($importacao.LinhasSemProjeto) -join ', ')"
    }

    $indices = New-Object System.Collections.Generic.List[int]
    $naoEncontrados = New-Object System.Collections.Generic.List[object]
    $vistos = New-Object 'System.Collections.Generic.HashSet[int]'

    foreach ($nomeProjeto in @($importacao.Projetos)) {
        $resultado = Localizar-ProjetoPorNomePlanilha -NomeProjetoPlanilha $nomeProjeto -Projetos $Projetos
        if (-not $resultado.Encontrado) {
            $naoEncontrados.Add([PSCustomObject]@{
                ProjetoPlanilha = $nomeProjeto
                Sugestoes       = @($resultado.Sugestoes)
            })
            continue
        }

        $indiceUmBase = [int]$resultado.Indice + 1
        if ($vistos.Add($indiceUmBase)) {
            $indices.Add($indiceUmBase)
            Write-Step "Projeto localizado [$($resultado.Motivo)]: $nomeProjeto -> $(Obter-RotuloPasta -Pasta $Projetos[$resultado.Indice])"
        }
    }

    if ($naoEncontrados.Count -gt 0) {
        Write-Warn "Projetos nao encontrados na concessao carregada: $($naoEncontrados.Count)"
        foreach ($item in @($naoEncontrados.ToArray() | Select-Object -First 20)) {
            $msg = "Nao encontrado: $($item.ProjetoPlanilha)"
            if (@($item.Sugestoes).Count -gt 0) {
                $msg += " | sugestao: $(@($item.Sugestoes)[0])"
            }
            Write-Warn $msg
        }
    }

    if ($indices.Count -eq 0) {
        throw "Nenhum projeto da planilha foi localizado na concessao selecionada."
    }

    Write-Step "Projetos selecionados pela planilha: $($indices.Count)"
    return @($indices.ToArray() | Sort-Object)
}

function Ler-SelecaoProjetosInterativa {
    param([array]$Projetos)

    if ($TodosProjetos) {
        return @(1..$Projetos.Count)
    }

    if (-not [string]::IsNullOrWhiteSpace($PlanilhaProjetos)) {
        return @(Ler-SelecaoProjetosPorPlanilha -Projetos $Projetos)
    }

    Write-Host ""
    Write-Host "Forma de selecao dos projetos:" -ForegroundColor Cyan
    Write-Host "01. Selecionar manualmente pelos numeros"
    Write-Host "02. Selecionar por planilha com coluna Projeto"
    while ($true) {
        $modo = (Read-Host "Escolha [1/2]").Trim()
        if ($modo -in @("1", "01")) {
            return @(Ler-SelecaoProjetos -Total $Projetos.Count)
        }
        if ($modo -in @("2", "02")) {
            return @(Ler-SelecaoProjetosPorPlanilha -Projetos $Projetos)
        }
        Write-Warn "Escolha 1 ou 2."
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
        # ConnectedProjectId pode exigir uma consulta remota. Ele sera lido
        # somente para os projetos escolhidos, depois da selecao.
        "{0:000}. {1} | ID PW: {2} | GUID PW: {3}" -f (
            $i + 1),
            (Obter-RotuloPasta -Pasta $projeto),
            (Obter-IdPasta -Pasta $projeto),
            (Obter-GuidPasta -Pasta $projeto) | Write-Host
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

function Obter-ConnectedProjectId {
    param([object]$Projeto)

    $prop = $Projeto.PSObject.Properties["ConnectedProjectId"]
    if ($prop -and $null -ne $prop.Value) {
        $guidProp = $prop.Value.PSObject.Properties["Guid"]
        if ($guidProp -and $guidProp.Value) {
            return $guidProp.Value.ToString()
        }
        $texto = $prop.Value.ToString()
        if ($texto -match "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}") {
            return $Matches[0]
        }
    }

    $link = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @("ProjectWiseWebLink")
    if ($link -match "[?&]projectId=([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
        return $Matches[1]
    }

    return ""
}

function Testar-ConnectedProjectIdValido {
    param([string]$ConnectedProjectId)

    if ([string]::IsNullOrWhiteSpace($ConnectedProjectId)) {
        return $false
    }

    if ($ConnectedProjectId -eq $GuidZero) {
        return $false
    }

    return ($ConnectedProjectId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

function Obter-ProjectWiseWebName {
    param([object]$Projeto)

    return Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @(
        "ConnectedProjectName", "ProjectWiseProjectName", "PWProjectName"
    )
}

function Obter-ProjectWiseWebNumber {
    param([object]$Projeto)

    return Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @(
        "ConnectedProjectNumber", "ProjectWiseProjectNumber", "PWProjectNumber"
    )
}

function Converter-ProjetoParaProjectWiseWeb {
    param(
        [object]$Projeto,
        [object]$Concessao
    )

    $connectedProjectId = Obter-ConnectedProjectId -Projeto $Projeto
    $aptoProjectWiseWeb = Testar-ConnectedProjectIdValido -ConnectedProjectId $connectedProjectId

    return [ordered]@{
        nome = Obter-RotuloPasta -Pasta $Projeto
        nomeTecnico = Obter-NomePasta -Pasta $Projeto
        descricao = Obter-DescricaoPasta -Pasta $Projeto
        concessaoProjectWise = Obter-RotuloPasta -Pasta $Concessao
        projectWiseId = Obter-IdPasta -Pasta $Projeto
        projectWiseGuid = Obter-GuidPasta -Pasta $Projeto
        connectedProjectId = $connectedProjectId
        connectSpaceId = $connectedProjectId
        projectId = $connectedProjectId
        projectWiseWebName = Obter-ProjectWiseWebName -Projeto $Projeto
        projectWiseWebNumber = Obter-ProjectWiseWebNumber -Projeto $Projeto
        projectWiseWebLink = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @("ProjectWiseWebLink")
        projectWiseWebViewLink = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @("ProjectWiseWebViewLink")
        origemSelecao = "ProjectWise ConnectedProjectId"
        aptoProjectWiseWeb = $aptoProjectWiseWeb
    }
}

try {
    Assert-Mta
    Write-Step "Importando modulo ProjectWise..."
    Importar-ModuloProjectWise
    Conectar-ProjectWise
    Write-Step "Conexao ProjectWise OK."

    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia"
    $concessoes = Ordenar-ItensPorNome -Itens (Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $pastaEngenharia) -Contexto "Engenharia")
    if ($concessoes.Count -eq 0) {
        throw "Nenhuma concessao encontrada."
    }

    Mostrar-Concessoes -Concessoes $concessoes
    $indiceConcessao = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $concessoes.Count
    $concessao = $concessoes[$indiceConcessao - 1]
    $nomeConcessao = Obter-RotuloPasta -Pasta $concessao

    Write-Step "Carregando projetos somente da concessao '$nomeConcessao'..."
    $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId (Obter-IdPasta -Pasta $concessao) -Contexto $nomeConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"
    $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -ProfundidadeMaxima $ProfundidadeProjetos
    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto encontrado em '$nomeConcessao'."
    }

    Mostrar-Projetos -Projetos $projetos
    $indices = Ler-SelecaoProjetosInterativa -Projetos $projetos
    $selecionados = @(
        foreach ($indice in $indices) {
            Converter-ProjetoParaProjectWiseWeb -Projeto $projetos[$indice - 1] -Concessao $concessao
        }
    )

    $semConnectedProject = @($selecionados | Where-Object { -not $_.aptoProjectWiseWeb })
    if ($semConnectedProject.Count -gt 0) {
        Write-Warn "$($semConnectedProject.Count) projeto(s) selecionado(s) nao possuem ConnectedProjectId e nao podem ser gerenciados no ProjectWise Web."
        foreach ($item in $semConnectedProject) {
            Write-Warn "Sem ConnectedProjectId valido: $($item.nome) | PW ID: $($item.projectWiseId) | ConnectedProjectId: $($item.connectedProjectId)"
        }
    }

    $dados = [ordered]@{
        timestamp = (Get-Date).ToString("s")
        origem = "ProjectWise"
        modoSelecao = "concessao_projetos_projectwise_web"
        profundidadeProjetos = $ProfundidadeProjetos
        concessao = [ordered]@{
            nome = $nomeConcessao
            id = Obter-IdPasta -Pasta $concessao
            guid = Obter-GuidPasta -Pasta $concessao
        }
        totalProjetosListados = $projetos.Count
        totalProjetosSelecionados = $selecionados.Count
        totalProjetosAptosProjectWiseWeb = @($selecionados | Where-Object { $_.aptoProjectWiseWeb }).Count
        projetos = @($selecionados)
    }

    [System.IO.File]::WriteAllText($Saida, ($dados | ConvertTo-Json -Depth 20), [System.Text.Encoding]::UTF8)

    Write-Host ""
    Write-Host "[OK] Projetos selecionados exportados." -ForegroundColor Green
    Write-Host "JSON: $Saida"
    Write-Host "Projetos selecionados: $($selecionados.Count)"
    Write-Host "Projetos aptos PW Web: $(@($selecionados | Where-Object { $_.aptoProjectWiseWeb }).Count)"
}
finally {
    Encerrar-SessaoProjectWise
}
