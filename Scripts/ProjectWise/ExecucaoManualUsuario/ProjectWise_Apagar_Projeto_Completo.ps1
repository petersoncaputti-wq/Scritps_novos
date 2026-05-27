<#
.SYNOPSIS
    Remove um projeto do ProjectWise Explorer e limpa acessos do PW Admin.

.DESCRIPTION
    Permite selecionar a concessao e os projetos dentro dela. Para cada projeto,
    remove as permissoes dos grupos/user lists do projeto, apaga a pasta do
    projeto com documentos/subpastas e remove os Groups e User Lists criados
    para o projeto.

    Por seguranca, o script simula por padrao. Use -Executar e confirme digitando
    APAGAR para gravar as remocoes.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Apagar_Projeto_Completo.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\ProjectWise_Apagar_Projeto_Completo.ps1" -NomeConcessao "Ecovias" -SiglaConcessao "ECO" -Executar
#>

[CmdletBinding()]
param(
    [string]$NomeConcessao = "",

    [string]$SiglaConcessao = "",

    [string]$Projeto = "",

    [string]$CaminhoProjeto = "",

    [int]$ProfundidadeProjetos = 0,

    [switch]$TodosProjetos,

    [switch]$RemoverGrupoConcessaoCompartilhado,

    [switch]$Executar,

    [switch]$NaoDesconectar
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArquivoLog = Join-Path $PastaLogs "PW_ApagarProjetoCompleto_$TimeStamp.log"
$ArquivoCsv = Join-Path $PastaLogs "PW_ApagarProjetoCompleto_$TimeStamp.csv"
$script:LoginPW = $null
$script:Resultados = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    $linha = "{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Nivel.ToUpperInvariant(), $Mensagem
    [System.IO.File]::AppendAllText($ArquivoLog, $linha + [Environment]::NewLine)
    Write-Host $linha
}

function Add-Resultado {
    param(
        [string]$Etapa,
        [string]$Tipo,
        [string]$Nome,
        [string]$Status,
        [string]$Mensagem = ""
    )

    $script:Resultados.Add([PSCustomObject]@{
        DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Etapa    = $Etapa
        Tipo     = $Tipo
        Nome     = $Nome
        Status   = $Status
        Mensagem = $Mensagem
    })
}

function Assert-Mta {
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartment -ne [System.Threading.ApartmentState]::MTA) {
        Write-Log "Recomendado executar com powershell.exe -MTA. Estado atual: $apartment" "WARN"
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
        $tentativas = @()

        if ($params -contains "BentleyIMS") {
            $tentativas += [PSCustomObject]@{ Rotulo = "$nomeCmdlet -BentleyIMS"; Args = @{ BentleyIMS = $true } }
        }
        if ($params -contains "UseGui") {
            $tentativas += [PSCustomObject]@{ Rotulo = "$nomeCmdlet -UseGui"; Args = @{ UseGui = $true } }
        }
        $tentativas += [PSCustomObject]@{ Rotulo = $nomeCmdlet; Args = @{} }

        foreach ($tentativa in $tentativas) {
            try {
                Write-Log "Tentando conexao via $($tentativa.Rotulo)..."
                $argsLogin = $tentativa.Args
                $script:LoginPW = & $nomeCmdlet @argsLogin -ErrorAction Stop
                Write-Log "Conexao realizada via $($tentativa.Rotulo)."
                return
            }
            catch {
                Write-Log "Falha em $($tentativa.Rotulo): $($_.Exception.Message)" "WARN"
            }
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
            Write-Log "Falha ao encerrar sessao via $nomeCmdlet`: $($_.Exception.Message)" "WARN"
        }
    }
}

function Get-ValorSeguroPropriedade {
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

function Get-IdPasta {
    param([object]$Pasta)
    return Get-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID")
}

function Get-CaminhoPasta {
    param([object]$Pasta)
    return Get-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("FullPath", "Fullpath", "FolderPath", "Path", "ProjectPath")
}

function Get-NomePasta {
    param([object]$Pasta)
    return Get-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Name", "FolderName", "ProjectName", "ObjectName")
}

function Get-DescricaoPasta {
    param([object]$Pasta)
    return Get-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("Description", "Descricao", "ProjectDescription", "FolderDescription")
}

function Get-RotuloPasta {
    param([object]$Pasta)

    $descricao = Get-DescricaoPasta -Pasta $Pasta
    if (-not [string]::IsNullOrWhiteSpace($descricao)) {
        return $descricao
    }

    return Get-NomePasta -Pasta $Pasta
}

function Get-NomeObjeto {
    param([object]$Objeto)
    return Get-ValorSeguroPropriedade -Objeto $Objeto -PossiveisNomes @("Name", "GroupName", "UserListName", "ObjectName", "MemberName")
}

function Ordenar-PastasPorNome {
    param([array]$Pastas)

    return @(
        $Pastas | Sort-Object -Property @{
            Expression = {
                (Get-RotuloPasta -Pasta $_).ToLowerInvariant()
            }
        }
    )
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
        Write-Log "Falha ao listar filhos de '$Contexto' ($FolderId): $($_.Exception.Message)" "WARN"
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
                (Get-NomePasta -Pasta $_) -ieq $nomeEsperado -or
                (Get-RotuloPasta -Pasta $_) -ieq $nomeEsperado
            }
        )

        if ($encontrada.Count -gt 0) {
            return $encontrada[0]
        }
    }

    $nomes = @($Pastas | ForEach-Object { Get-RotuloPasta -Pasta $_ }) -join ", "
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

        Write-Log "Informe um numero entre $Minimo e $Maximo." "WARN"
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

        Write-Log "Selecao invalida." "WARN"
    }
}

function Obter-ProjetosDaPastaProjetos {
    param(
        [object]$PastaProjetos,
        [int]$ProfundidadeMaxima
    )

    $idPastaProjetos = Get-IdPasta -Pasta $PastaProjetos
    if ([string]::IsNullOrWhiteSpace($idPastaProjetos)) {
        throw "Nao foi possivel obter o ID da pasta Projetos."
    }

    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idPastaProjetos; Nivel = 0; Nome = "Projetos" })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId -Contexto $atual.Nome

        foreach ($filho in $filhos) {
            $idFilho = Get-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) {
                continue
            }

            $resultado += $filho
            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $fila.Enqueue([PSCustomObject]@{
                    FolderId = $idFilho
                    Nivel    = ($atual.Nivel + 1)
                    Nome     = (Get-RotuloPasta -Pasta $filho)
                })
            }
        }
    }

    $mapa = @{}
    foreach ($item in $resultado) {
        $id = Get-IdPasta -Pasta $item
        if (-not [string]::IsNullOrWhiteSpace($id) -and -not $mapa.ContainsKey($id)) {
            $mapa[$id] = $item
        }
    }

    return @(Ordenar-PastasPorNome -Pastas @($mapa.Values))
}

function Obter-Concessao {
    param([array]$Concessoes)

    if (-not [string]::IsNullOrWhiteSpace($NomeConcessao)) {
        $matches = @(
            $Concessoes | Where-Object {
                (Get-RotuloPasta -Pasta $_) -like "*$NomeConcessao*" -or
                (Get-NomePasta -Pasta $_) -like "*$NomeConcessao*"
            }
        )

        if ($matches.Count -eq 1) {
            return $matches[0]
        }
        if ($matches.Count -gt 1) {
            Write-Log "Mais de uma concessao encontrada para '$NomeConcessao'. Selecione pelo numero." "WARN"
            $Concessoes = $matches
        }
        else {
            throw "Nenhuma concessao encontrada contendo '$NomeConcessao'."
        }
    }

    Write-Host ""
    Write-Host "Concessoes disponiveis:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Concessoes.Count; $i++) {
        "{0:00}. {1}" -f ($i + 1), (Get-RotuloPasta -Pasta $Concessoes[$i]) | Write-Host
    }

    $indice = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $Concessoes.Count
    return $Concessoes[$indice - 1]
}

function Obter-ProjetosSelecionados {
    param([object]$PastaProjetos)

    $projetos = @(Obter-ProjetosDaPastaProjetos -PastaProjetos $PastaProjetos -ProfundidadeMaxima $ProfundidadeProjetos)
    if ($projetos.Count -eq 0) {
        throw "Nenhum projeto encontrado na pasta Projetos."
    }

    if (-not [string]::IsNullOrWhiteSpace($Projeto)) {
        $matches = @(
            $projetos | Where-Object {
                (Get-NomePasta -Pasta $_) -like "*$Projeto*" -or
                (Get-RotuloPasta -Pasta $_) -like "*$Projeto*"
            }
        )

        if ($matches.Count -eq 1) {
            return @($matches[0])
        }
        if ($matches.Count -gt 1) {
            Write-Log "Mais de um projeto encontrado para '$Projeto'. Selecione pelo numero." "WARN"
            $projetos = $matches
        }
        else {
            throw "Nenhum projeto encontrado contendo '$Projeto'."
        }
    }

    Write-Host ""
    Write-Host "Projetos encontrados:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $projetos.Count; $i++) {
        $nome = Get-NomePasta -Pasta $projetos[$i]
        $rotulo = Get-RotuloPasta -Pasta $projetos[$i]
        if ($rotulo -and $rotulo -ne $nome) {
            "{0:00}. {1} - {2}" -f ($i + 1), $nome, $rotulo | Write-Host
        }
        else {
            "{0:00}. {1}" -f ($i + 1), $nome | Write-Host
        }
    }

    $indices = Ler-SelecaoProjetos -Total $projetos.Count
    return @($indices | ForEach-Object { $projetos[$_ - 1] })
}

function Confirmar-Execucao {
    param(
        [object[]]$ProjetosSelecionados,
        [string]$Concessao
    )

    if (-not $Executar) {
        Write-Log "Modo simulacao ativo. Nenhuma remocao sera executada." "WARN"
        return $false
    }

    Write-Host ""
    Write-Host "ATENCAO: esta acao vai apagar dados no ProjectWise." -ForegroundColor Red
    Write-Host "Concessao: $Concessao" -ForegroundColor Yellow
    Write-Host "Sigla    : $SiglaConcessao" -ForegroundColor Yellow
    Write-Host "Projetos : $($ProjetosSelecionados.Count)" -ForegroundColor Yellow
    foreach ($item in $ProjetosSelecionados) {
        Write-Host " - $(Get-NomePasta -Pasta $item)"
    }
    Write-Host ""

    $confirmacao = Read-Host "Digite APAGAR para confirmar"
    if ($confirmacao -cne "APAGAR") {
        Write-Log "Confirmacao invalida. Execucao cancelada." "WARN"
        return $false
    }

    return $true
}

function Get-ObjetosProjetoAdmin {
    param([string]$NomeProjeto)

    $prefixo = "$SiglaConcessao-$NomeProjeto"

    $grupos = @(Get-PWGroups -ErrorAction Stop | Where-Object {
        $nome = Get-NomeObjeto $_
        $nome -ieq $prefixo -or $nome -like "$prefixo-*"
    })

    if ($RemoverGrupoConcessaoCompartilhado) {
        $grupoCompartilhado = "$SiglaConcessao-ENGENHARIA UNIDADE"
        $grupo = @(Get-PWGroups -GroupName $grupoCompartilhado -ErrorAction SilentlyContinue | Where-Object {
            (Get-NomeObjeto $_) -ieq $grupoCompartilhado
        })
        $grupos += $grupo
    }

    $userLists = @(Get-PWUserLists -ErrorAction Stop | Where-Object {
        $nome = Get-NomeObjeto $_
        $nome -ieq $prefixo -or $nome -like "$prefixo-*"
    })

    return [PSCustomObject]@{
        Groups    = @($grupos | Sort-Object { Get-NomeObjeto $_ } -Unique)
        UserLists = @($userLists | Sort-Object { (Get-NomeObjeto $_).Length } -Descending -Unique)
    }
}

function Get-ArvorePastasProjeto {
    param([object]$PastaRaiz)

    $resultado = New-Object System.Collections.Generic.List[object]
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue($PastaRaiz)

    while ($fila.Count -gt 0) {
        $pasta = $fila.Dequeue()
        $resultado.Add($pasta)

        $id = Get-IdPasta -Pasta $pasta
        if ([string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        try {
            $filhos = @(Get-PWFoldersImmediateChildren -FolderID ([int]$id) -ErrorAction Stop | Where-Object { $_ -ne $null })
            foreach ($filho in $filhos) {
                $fila.Enqueue($filho)
            }
        }
        catch {
            Write-Log "Falha ao listar subpastas de '$((Get-CaminhoPasta $pasta))': $($_.Exception.Message)" "WARN"
        }
    }

    return @($resultado)
}

function Remover-PermissoesDosMembros {
    param(
        [object[]]$Pastas,
        [object[]]$Groups,
        [object[]]$UserLists,
        [bool]$PodeExecutar
    )

    $membros = @()
    $membros += @($Groups | ForEach-Object {
        [PSCustomObject]@{ Tipo = "Group"; Nome = Get-NomeObjeto $_ }
    })
    $membros += @($UserLists | ForEach-Object {
        [PSCustomObject]@{ Tipo = "UserList"; Nome = Get-NomeObjeto $_ }
    })

    foreach ($membro in @($membros | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Nome) } | Sort-Object Tipo, Nome -Unique)) {
        foreach ($pasta in $Pastas) {
            $caminho = Get-CaminhoPasta -Pasta $pasta

            foreach ($tipoSeguranca in @("FolderSecurity", "DocumentSecurity")) {
                $acao = "Remover permissao $tipoSeguranca de $($membro.Tipo) '$($membro.Nome)' em '$caminho'"
                if (-not $PodeExecutar) {
                    Write-Log "[SIMULACAO] $acao"
                    Add-Resultado -Etapa "Permissao" -Tipo $tipoSeguranca -Nome $membro.Nome -Status "Simulado" -Mensagem $caminho
                    continue
                }

                try {
                    if ($tipoSeguranca -eq "FolderSecurity") {
                        Remove-PWFolderSecurity -InputFolder $pasta -MemberType $membro.Tipo -MemberName $membro.Nome -FolderSecurity -ErrorAction Stop
                    }
                    else {
                        Remove-PWFolderSecurity -InputFolder $pasta -MemberType $membro.Tipo -MemberName $membro.Nome -DocumentSecurity -ErrorAction Stop
                    }

                    Write-Log $acao
                    Add-Resultado -Etapa "Permissao" -Tipo $tipoSeguranca -Nome $membro.Nome -Status "Removido" -Mensagem $caminho
                }
                catch {
                    Write-Log "Falha ao remover permissao de '$($membro.Nome)' em '$caminho': $($_.Exception.Message)" "WARN"
                    Add-Resultado -Etapa "Permissao" -Tipo $tipoSeguranca -Nome $membro.Nome -Status "Erro" -Mensagem $_.Exception.Message
                }
            }
        }
    }
}

function Remover-ProjetoExplorer {
    param(
        [object]$PastaRaiz,
        [string]$CaminhoProjetoAtual,
        [bool]$PodeExecutar
    )

    $caminho = Get-CaminhoPasta -Pasta $PastaRaiz
    if ([string]::IsNullOrWhiteSpace($caminho)) {
        $caminho = $CaminhoProjetoAtual
    }

    if (-not $PodeExecutar) {
        Write-Log "[SIMULACAO] Remover projeto do Explorer: $caminho"
        Add-Resultado -Etapa "Projeto" -Tipo "Folder" -Nome $caminho -Status "Simulado"
        return
    }

    Remove-PWFolder -InputFolder $PastaRaiz -RemoveDocuments -RemoveFolders -ProceedWithDelete -ErrorAction Stop
    Write-Log "Projeto removido do Explorer: $caminho"
    Add-Resultado -Etapa "Projeto" -Tipo "Folder" -Nome $caminho -Status "Removido"
}

function Remover-GruposAdmin {
    param(
        [object[]]$Groups,
        [bool]$PodeExecutar
    )

    foreach ($grupo in $Groups) {
        $nome = Get-NomeObjeto $grupo
        if ([string]::IsNullOrWhiteSpace($nome)) {
            continue
        }

        if (-not $PodeExecutar) {
            Write-Log "[SIMULACAO] Remover Group: $nome"
            Add-Resultado -Etapa "PW Admin" -Tipo "Group" -Nome $nome -Status "Simulado"
            continue
        }

        try {
            Remove-PWGroupByMatch -InputGroups $grupo -ErrorAction Stop
            Write-Log "Group removido: $nome"
            Add-Resultado -Etapa "PW Admin" -Tipo "Group" -Nome $nome -Status "Removido"
        }
        catch {
            Write-Log "Falha ao remover Group '$nome': $($_.Exception.Message)" "WARN"
            Add-Resultado -Etapa "PW Admin" -Tipo "Group" -Nome $nome -Status "Erro" -Mensagem $_.Exception.Message
        }
    }
}

function Remover-UserListsAdmin {
    param(
        [object[]]$UserLists,
        [bool]$PodeExecutar
    )

    foreach ($userList in $UserLists) {
        $nome = Get-NomeObjeto $userList
        if ([string]::IsNullOrWhiteSpace($nome)) {
            continue
        }

        if (-not $PodeExecutar) {
            Write-Log "[SIMULACAO] Remover User List: $nome"
            Add-Resultado -Etapa "PW Admin" -Tipo "UserList" -Nome $nome -Status "Simulado"
            continue
        }

        try {
            Remove-PWUserListByMatch -InputUserLists $userList -ErrorAction Stop
            Write-Log "User List removida: $nome"
            Add-Resultado -Etapa "PW Admin" -Tipo "UserList" -Nome $nome -Status "Removido"
        }
        catch {
            Write-Log "Falha ao remover User List '$nome': $($_.Exception.Message)" "WARN"
            Add-Resultado -Etapa "PW Admin" -Tipo "UserList" -Nome $nome -Status "Erro" -Mensagem $_.Exception.Message
        }
    }
}

try {
    Assert-Mta
    Importar-ModuloProjectWise

    $NomeConcessao = $NomeConcessao.Trim()
    $SiglaConcessao = $SiglaConcessao.Trim()
    $Projeto = $Projeto.Trim()

    Write-Log "Inicio da rotina de exclusao completa de projeto."
    Write-Log "Executar: $Executar"

    Conectar-ProjectWise

    if (-not [string]::IsNullOrWhiteSpace($CaminhoProjeto)) {
        $pastaRaiz = Get-PWFolders -FolderPath $CaminhoProjeto -PopulatePaths -JustOne -ErrorAction Stop
        if (-not $pastaRaiz) {
            throw "Projeto nao encontrado no ProjectWise Explorer: $CaminhoProjeto"
        }

        if ([string]::IsNullOrWhiteSpace($Projeto)) {
            $Projeto = Get-NomePasta -Pasta $pastaRaiz
        }

        if ([string]::IsNullOrWhiteSpace($NomeConcessao)) {
            $partes = $CaminhoProjeto -split '\\'
            if ($partes.Count -ge 2) {
                $NomeConcessao = $partes[1]
            }
        }

        if ([string]::IsNullOrWhiteSpace($SiglaConcessao)) {
            $SiglaConcessao = (Read-Host "Informe a sigla da concessao usada nos Groups/User Lists").Trim()
        }

        $projetosSelecionados = @($pastaRaiz)
        $nomeConcessaoSelecionada = $NomeConcessao
    }
    else {
        $pastasRaiz = Get-PWRootFoldersSafe
        $pastaEngenharia = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis @("Engenharia", "Engineering", "ENGENHARIA") -Descricao "pasta de engenharia"
        $filhosEngenharia = Get-PWChildFoldersOrEmpty -FolderId (Get-IdPasta -Pasta $pastaEngenharia) -Contexto (Get-RotuloPasta -Pasta $pastaEngenharia)
        $concessao = Obter-Concessao -Concessoes (Ordenar-PastasPorNome -Pastas $filhosEngenharia)
        $nomeConcessaoSelecionada = Get-NomePasta -Pasta $concessao
        if ([string]::IsNullOrWhiteSpace($nomeConcessaoSelecionada)) {
            $nomeConcessaoSelecionada = Get-RotuloPasta -Pasta $concessao
        }

        if ([string]::IsNullOrWhiteSpace($SiglaConcessao)) {
            $SiglaConcessao = (Read-Host "Informe a sigla da concessao usada nos Groups/User Lists").Trim()
        }

        if ([string]::IsNullOrWhiteSpace($SiglaConcessao)) {
            throw "A sigla da concessao e obrigatoria para localizar Groups e User Lists no PW Admin."
        }

        $filhosConcessao = Get-PWChildFoldersOrEmpty -FolderId (Get-IdPasta -Pasta $concessao) -Contexto $nomeConcessaoSelecionada
        $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis @("Projetos", "Projeto", "Projects", "Project") -Descricao "pasta de projetos da concessao '$nomeConcessaoSelecionada'"
        $projetosSelecionados = @(Obter-ProjetosSelecionados -PastaProjetos $pastaProjetos)
    }

    if ($projetosSelecionados.Count -eq 0) {
        throw "Nenhum projeto selecionado."
    }

    Write-Host ""
    Write-Host "Resumo da selecao" -ForegroundColor Cyan
    Write-Host "Concessao: $nomeConcessaoSelecionada"
    Write-Host "Sigla    : $SiglaConcessao"
    Write-Host "Projetos : $($projetosSelecionados.Count)"
    Write-Host "Log      : $ArquivoLog"
    Write-Host ""

    foreach ($projetoSelecionado in $projetosSelecionados) {
        Write-Host " - $(Get-NomePasta -Pasta $projetoSelecionado)"
    }

    $podeExecutar = Confirmar-Execucao -ProjetosSelecionados $projetosSelecionados -Concessao $nomeConcessaoSelecionada

    foreach ($projetoSelecionado in $projetosSelecionados) {
        $nomeProjetoAtual = Get-NomePasta -Pasta $projetoSelecionado
        if ([string]::IsNullOrWhiteSpace($nomeProjetoAtual)) {
            throw "Nao foi possivel obter o nome tecnico de um projeto selecionado."
        }

        $caminhoProjetoAtual = Get-CaminhoPasta -Pasta $projetoSelecionado
        if ([string]::IsNullOrWhiteSpace($caminhoProjetoAtual)) {
            $caminhoProjetoAtual = "ENGENHARIA\$nomeConcessaoSelecionada\Projetos\$nomeProjetoAtual"
        }

        Write-Log "Processando projeto: $caminhoProjetoAtual"
        Write-Log "Prefixo Admin: $SiglaConcessao-$nomeProjetoAtual"

        $objetosAdmin = Get-ObjetosProjetoAdmin -NomeProjeto $nomeProjetoAtual
        $pastasProjeto = @(Get-ArvorePastasProjeto -PastaRaiz $projetoSelecionado)

        Write-Host ""
        Write-Host "Resumo encontrado para $nomeProjetoAtual" -ForegroundColor Cyan
        Write-Host "Pastas no projeto : $($pastasProjeto.Count)"
        Write-Host "Groups do projeto : $($objetosAdmin.Groups.Count)"
        Write-Host "User Lists projeto: $($objetosAdmin.UserLists.Count)"

        Add-Resultado -Etapa "Resumo" -Tipo "Folder" -Nome $caminhoProjetoAtual -Status "Encontrado" -Mensagem "$($pastasProjeto.Count) pasta(s)"
        Add-Resultado -Etapa "Resumo" -Tipo "Group" -Nome "$SiglaConcessao-$nomeProjetoAtual" -Status "Encontrado" -Mensagem "$($objetosAdmin.Groups.Count) grupo(s)"
        Add-Resultado -Etapa "Resumo" -Tipo "UserList" -Nome "$SiglaConcessao-$nomeProjetoAtual" -Status "Encontrado" -Mensagem "$($objetosAdmin.UserLists.Count) user list(s)"

        Remover-PermissoesDosMembros -Pastas $pastasProjeto -Groups $objetosAdmin.Groups -UserLists $objetosAdmin.UserLists -PodeExecutar $podeExecutar
        Remover-ProjetoExplorer -PastaRaiz $projetoSelecionado -CaminhoProjetoAtual $caminhoProjetoAtual -PodeExecutar $podeExecutar
        Remover-GruposAdmin -Groups $objetosAdmin.Groups -PodeExecutar $podeExecutar
        Remover-UserListsAdmin -UserLists $objetosAdmin.UserLists -PodeExecutar $podeExecutar
    }

    Write-Log "Rotina finalizada."
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" "ERROR"
    Add-Resultado -Etapa "Fatal" -Tipo "Erro" -Nome "Script" -Status "Erro" -Mensagem $_.Exception.Message
    throw
}
finally {
    if ($script:Resultados.Count -gt 0) {
        $script:Resultados | Export-Csv -Path $ArquivoCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log "CSV de resultado gerado: $ArquivoCsv"
    }

    Encerrar-SessaoProjectWise
}
