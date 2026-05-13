<#
.SYNOPSIS
    Lista concessoes e projetos diretamente do ProjectWise e exporta JSON para a V2 do PWDM.

.DESCRIPTION
    Este script e somente leitura. Ele conecta no ProjectWise usando os cmdlets disponiveis
    no ambiente, localiza a raiz de engenharia, lista as concessoes e procura a pasta de
    projetos dentro de cada concessao.

    A saida JSON foi pensada para ser consumida pelo Python da rotina PWDM V2.

.EXAMPLE
    .\pw_listar_concessoes_projetos_v2.ps1

.EXAMPLE
    .\pw_listar_concessoes_projetos_v2.ps1 -ConcessaoNome "Ecovias" -ProfundidadeProjetos 1
#>

[CmdletBinding()]
param(
    [string]$Saida,
    [string]$ConcessaoNome = "",
    [int]$ProfundidadeProjetos = 1,
    [switch]$NaoDesconectar
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")

if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if ([string]::IsNullOrWhiteSpace($Saida)) {
    $Saida = Join-Path $PastaLogs "pw_concessoes_projetos_v2_$TimeStamp.json"
}

$ArquivoLog = Join-Path $PastaLogs "pw_concessoes_projetos_v2_$TimeStamp.log"
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

function Importar-ModuloProjectWise {
    $modulosPossiveis = @(
        "pwps_dab",
        "PWPS_DAB",
        "Bentley.PowerShell.ProjectWise",
        "ProjectWisePowerShell"
    )

    foreach ($modulo in $modulosPossiveis) {
        try {
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
        catch {
            Write-Log "Falha ao importar modulo '$modulo': $($_.Exception.Message)" "WARN"
        }
    }

    throw "Nao foi possivel localizar/importar o modulo do ProjectWise. Verifique se o PWPS_DAB esta instalado."
}

function Conectar-ProjectWise {
    $cmdletsConexao = @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")

    foreach ($nomeCmdlet in $cmdletsConexao) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        try {
            $params = @($cmd.Parameters.Keys)
            Write-Log "Tentando conexao com '$nomeCmdlet'."

            if ($params -contains "BentleyIMS") {
                $script:LoginPW = & $nomeCmdlet -BentleyIMS -ErrorAction Stop
                Write-Log "Conexao realizada via '$nomeCmdlet -BentleyIMS'."
                return
            }

            if ($params -contains "UseGui") {
                $script:LoginPW = & $nomeCmdlet -UseGui -ErrorAction Stop
                Write-Log "Conexao realizada via '$nomeCmdlet -UseGui'."
                return
            }

            $script:LoginPW = & $nomeCmdlet -ErrorAction Stop
            Write-Log "Conexao realizada via '$nomeCmdlet'."
            return
        }
        catch {
            Write-Log "Falha ao conectar via '$nomeCmdlet': $($_.Exception.Message)" "WARN"
        }
    }

    throw "Nao foi possivel estabelecer conexao com o ProjectWise."
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

            Write-Log "Sessao ProjectWise encerrada via '$nomeCmdlet'."
            return
        }
        catch {
            Write-Log "Falha ao encerrar sessao via '$nomeCmdlet': $($_.Exception.Message)" "WARN"
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

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        "ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID"
    )
}

function Obter-GuidPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        "ProjectGUID", "ProjectGuid", "FolderGUID", "FolderGuid", "GUID", "Guid"
    )
}

function Obter-NomePasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        "Name", "FolderName", "ProjectName", "ObjectName"
    )
}

function Obter-DescricaoPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        "Description", "Descricao", "ProjectDescription", "FolderDescription"
    )
}

function Obter-CaminhoPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        "FullPath", "FolderPath", "Path", "ProjectPath"
    )
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

    return @(
        $Itens | Sort-Object -Property @{
            Expression = {
                (Obter-RotuloPasta -Pasta $_).ToLowerInvariant()
            }
        }
    )
}

function Get-PWRootFoldersSafe {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    $params = @($cmd.Parameters.Keys)

    if ($params -contains "Root") {
        $pastas = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    }
    else {
        $pastas = @(Get-PWFoldersImmediateChildren -ErrorAction Stop)
    }

    $pastas = @($pastas | Where-Object { $_ -ne $null })
    if ($pastas.Count -eq 0) {
        throw "Nenhuma pasta foi retornada na raiz do ProjectWise."
    }

    return $pastas
}

function Get-PWChildFoldersOrEmpty {
    param([string]$FolderId)

    if ([string]::IsNullOrWhiteSpace($FolderId)) {
        return @()
    }

    try {
        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
        return @($pastas | Where-Object { $_ -ne $null })
    }
    catch {
        Write-Log "Nao foi possivel listar filhos do FolderID $FolderId`: $($_.Exception.Message)" "WARN"
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

function Converter-PastaParaResumo {
    param(
        [object]$Pasta,
        [string]$CaminhoManual = ""
    )

    $nome = Obter-NomePasta -Pasta $Pasta
    $rotulo = Obter-RotuloPasta -Pasta $Pasta
    $caminho = Obter-CaminhoPasta -Pasta $Pasta
    if ([string]::IsNullOrWhiteSpace($caminho)) {
        $caminho = $CaminhoManual
    }

    return [ordered]@{
        nome      = $rotulo
        nomeTecnico = $nome
        descricao = Obter-DescricaoPasta -Pasta $Pasta
        id        = Obter-IdPasta -Pasta $Pasta
        guid      = Obter-GuidPasta -Pasta $Pasta
        caminho   = $caminho
    }
}

function Obter-ProjetosDaPastaProjetos {
    param(
        [object]$PastaProjetos,
        [string]$CaminhoBase,
        [int]$ProfundidadeMaxima
    )

    $idPastaProjetos = Obter-IdPasta -Pasta $PastaProjetos
    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{
        FolderId = $idPastaProjetos
        Nivel = 0
        Caminho = $CaminhoBase
    })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId

        foreach ($filho in $filhos) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) {
                continue
            }

            $nomeFilho = Obter-RotuloPasta -Pasta $filho
            $caminhoFilho = "$($atual.Caminho)\$nomeFilho"
            $resumo = Converter-PastaParaResumo -Pasta $filho -CaminhoManual $caminhoFilho
            $resumo["nivelRelativo"] = $atual.Nivel
            $resultado += [PSCustomObject]$resumo

            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $netos = Get-PWChildFoldersOrEmpty -FolderId $idFilho
                if ($netos.Count -gt 0) {
                    $fila.Enqueue([PSCustomObject]@{
                        FolderId = $idFilho
                        Nivel = ($atual.Nivel + 1)
                        Caminho = $caminhoFilho
                    })
                }
            }
        }
    }

    $mapa = [ordered]@{}
    foreach ($projeto in $resultado) {
        $chave = if ($projeto.id) { $projeto.id } elseif ($projeto.guid) { $projeto.guid } else { $projeto.caminho }
        if (-not $mapa.Contains($chave)) {
            $mapa[$chave] = $projeto
        }
    }

    return @(Ordenar-ItensPorNome -Itens @($mapa.Values))
}

function Obter-ConcessoesEProjetos {
    Write-Log "Listando pastas raiz do ProjectWise."
    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "pasta de engenharia na raiz"
    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenharia
    $nomeEngenharia = Obter-RotuloPasta -Pasta $pastaEngenharia

    if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
        throw "Nao foi possivel obter o ID da pasta de engenharia."
    }

    Write-Log "Pasta de engenharia encontrada: $nomeEngenharia ($idEngenharia)."
    $concessoesRaw = Get-PWChildFoldersOrEmpty -FolderId $idEngenharia
    $concessoesRaw = Ordenar-ItensPorNome -Itens $concessoesRaw

    if (-not [string]::IsNullOrWhiteSpace($ConcessaoNome)) {
        $concessoesRaw = @(
            $concessoesRaw | Where-Object {
                (Obter-RotuloPasta -Pasta $_) -like "*$ConcessaoNome*" -or
                (Obter-NomePasta -Pasta $_) -like "*$ConcessaoNome*"
            }
        )
    }

    $concessoes = @()
    foreach ($concessao in $concessoesRaw) {
        $nomeConcessao = Obter-RotuloPasta -Pasta $concessao
        $idConcessao = Obter-IdPasta -Pasta $concessao

        if ([string]::IsNullOrWhiteSpace($idConcessao)) {
            Write-Log "Concessao '$nomeConcessao' ignorada: sem ID." "WARN"
            continue
        }

        Write-Log "Lendo concessao '$nomeConcessao'."
        $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId $idConcessao
        $pastaProjetos = $null
        $projetos = @()
        $erroProjetos = ""

        try {
            $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "pasta de projetos da concessao '$nomeConcessao'"
            $caminhoProjetos = "$nomeEngenharia\$nomeConcessao\$(Obter-RotuloPasta -Pasta $pastaProjetos)"
            $projetos = Obter-ProjetosDaPastaProjetos -PastaProjetos $pastaProjetos -CaminhoBase $caminhoProjetos -ProfundidadeMaxima $ProfundidadeProjetos
        }
        catch {
            $erroProjetos = $_.Exception.Message
            Write-Log "Nao foi possivel obter projetos da concessao '$nomeConcessao': $erroProjetos" "WARN"
        }

        $resumoConcessao = Converter-PastaParaResumo -Pasta $concessao -CaminhoManual "$nomeEngenharia\$nomeConcessao"
        $resumoConcessao["pastaProjetosEncontrada"] = ($null -ne $pastaProjetos)
        $resumoConcessao["pastaProjetos"] = if ($pastaProjetos) { Converter-PastaParaResumo -Pasta $pastaProjetos -CaminhoManual "$nomeEngenharia\$nomeConcessao\$(Obter-RotuloPasta -Pasta $pastaProjetos)" } else { $null }
        $resumoConcessao["erroProjetos"] = $erroProjetos
        $resumoConcessao["quantidadeProjetos"] = @($projetos).Count
        $resumoConcessao["projetos"] = @($projetos)

        $concessoes += [PSCustomObject]$resumoConcessao
        Write-Log "Concessao '$nomeConcessao': $(@($projetos).Count) projeto(s)."
    }

    return [ordered]@{
        timestamp = (Get-Date).ToString("s")
        origem = "ProjectWise"
        raizEngenharia = Converter-PastaParaResumo -Pasta $pastaEngenharia -CaminhoManual $nomeEngenharia
        filtroConcessao = $ConcessaoNome
        profundidadeProjetos = $ProfundidadeProjetos
        quantidadeConcessoes = @($concessoes).Count
        concessoes = @($concessoes)
    }
}

try {
    Write-Log "Iniciando exportacao de concessoes/projetos ProjectWise."
    Importar-ModuloProjectWise
    Conectar-ProjectWise

    $dados = Obter-ConcessoesEProjetos
    $json = $dados | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($Saida, $json, [System.Text.Encoding]::UTF8)

    Write-Log "JSON salvo em: $Saida"
    Write-Host ""
    Write-Host "[OK] Exportacao concluida."
    Write-Host "JSON: $Saida"
    Write-Host "Log : $ArquivoLog"
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    Encerrar-SessaoProjectWise
}
