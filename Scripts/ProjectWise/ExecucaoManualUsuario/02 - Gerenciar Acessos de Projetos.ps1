# =========================================================
# PROJECTWISE - GESTAO DE ACESSOS COM INTERFACE VISUAL
# VERSAO REVISADA - LOGIN NA ABERTURA + ARVORE DE PROJETOS
# =========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------
# CONFIGURACOES
# ---------------------------------------------------------
$PastaLog = Join-Path $PSScriptRoot "Logs"
$NomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$NomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")

if (-not (Test-Path $PastaLog)) {
    New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArquivoLog = Join-Path $PastaLog "PW_UI_ExecucaoPermissoes_$TimeStamp.log"

# ---------------------------------------------------------
# ESTADO GLOBAL
# ---------------------------------------------------------
$script:LoginPW = $null
$script:UsuarioEncontrado = $null
$script:UsuariosSelecionados = @()
$script:ModoLoteUsuarios = $false
$script:ConcessaoSelecionada = $null
$script:SelecoesProjetos = @()
$script:ResultadosExecucao = @()
$script:AtualizandoCheckTree = $false
$script:InicializacaoConcluida = $false
$script:ProjetosCarregados = @()
$script:CacheGruposPW = $null
$script:IndiceGruposPW = $null
$script:CacheParametrosCmdlet = @{}
$script:VersaoCacheAcessos = 1
$script:DiasValidadeCacheAcessos = 7
$script:ResumoImportacaoUsuarios = $null
$script:ResumoImportacaoProjetos = $null
$script:FiltroExclusaoAvancado = ""
$script:PerfisPadraoProjeto = @(
    [PSCustomObject]@{ Rotulo = "Assistente Eng"; Sufixo = "ASSISTENTE ENG" },
    [PSCustomObject]@{ Rotulo = "Consulta Todas Rev"; Sufixo = "CONSULTA TODAS REV" },
    [PSCustomObject]@{ Rotulo = "Consulta Ultima Rev"; Sufixo = "CONSULTA ULTIMA REV" },
    [PSCustomObject]@{ Rotulo = "Eng Consulta Todas Rev"; Sufixo = "ENG CONSULTA TODAS REV" },
    [PSCustomObject]@{ Rotulo = "Eng Consulta Ultima Rev"; Sufixo = "ENG CONSULTA ULTIMA REV" },
    [PSCustomObject]@{ Rotulo = "Especialista"; Sufixo = "ESPECIALISTA" },
    [PSCustomObject]@{ Rotulo = "Gestor Eng"; Sufixo = "GESTOR ENG" },
    [PSCustomObject]@{ Rotulo = "Gestor Unidade"; Sufixo = "GESTOR UNIDADE" },
    [PSCustomObject]@{ Rotulo = "Poder Concedente"; Sufixo = "PODER CONCEDENTE" },
    [PSCustomObject]@{ Rotulo = "Projetista"; Sufixo = "PROJETISTA" }
)

# ---------------------------------------------------------
# FUNCOES DE LOG
# ---------------------------------------------------------
function Write-Log {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    $linha = "{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Nivel.ToUpper(), $Mensagem
    try {
        [System.IO.File]::AppendAllText($ArquivoLog, $linha + [Environment]::NewLine)
    }
    catch {
    }
}

function Write-UiLog {
    param([string]$Mensagem)

    if ($script:LogBox) {
        $script:LogBox.AppendText($Mensagem + [Environment]::NewLine)
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
        "Informação",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Ask-UiYesNo {
    param(
        [string]$Mensagem,
        [string]$Titulo = "Confirmação"
    )

    $resposta = [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        $Titulo,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($resposta -eq [System.Windows.Forms.DialogResult]::Yes)
}

# ---------------------------------------------------------
# FUNCOES UTILITARIAS
# ---------------------------------------------------------
function Obter-ParametrosCmdlet {
    param([string]$NomeCmdlet)

    if (-not $script:CacheParametrosCmdlet) { $script:CacheParametrosCmdlet = @{} }
    if (-not $script:CacheParametrosCmdlet.ContainsKey($NomeCmdlet)) {
        $cmd = Get-Command $NomeCmdlet -ErrorAction Stop
        $script:CacheParametrosCmdlet[$NomeCmdlet] = @($cmd.Parameters.Keys)
    }
    return @($script:CacheParametrosCmdlet[$NomeCmdlet])
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
        "ProjectID","FolderID","Id","ID"
    )
}

function Obter-NomeAmigavelItem {
    param([object]$Item)

    return Obter-ValorSeguroPropriedade -Objeto $Item -PossiveisNomes @(
        "Name","FolderName","ProjectName","ObjectName","MemberName","UserName","FolderPath","FullPath","Path"
    )
}

function Obter-DescricaoItem {
    param([object]$Item)

    return Obter-ValorSeguroPropriedade -Objeto $Item -PossiveisNomes @(
        "Description","Descricao","ProjectDescription","FolderDescription"
    )
}

function Obter-RotuloProjeto {
    param([object]$Projeto)

    $descricao = Obter-DescricaoItem -Item $Projeto
    $nome = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @("Name","FolderName","ProjectName")

    if (-not [string]::IsNullOrWhiteSpace($descricao)) {
        return $descricao
    }

    return $nome
}

function Obter-RotuloConcessao {
    param([object]$Concessao)

    return Obter-ValorSeguroPropriedade -Objeto $Concessao -PossiveisNomes @("Name","FolderName","ProjectName")
}

function Normalizar-Texto {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return ""
    }

    $normalizado = $Texto.Trim().ToLowerInvariant()
    $normalizado = $normalizado -replace '\s+', ' '
    return $normalizado
}

function Ordenar-ItensPorNome {
    param([array]$Itens)

    return @(
        $Itens | Sort-Object -Property @{
            Expression = {
                (Obter-NomeAmigavelItem -Item $_).ToLower()
            }
        }
    )
}

function Validar-Email {
    param([string]$Email)

    if ([string]::IsNullOrWhiteSpace($Email)) {
        throw "O e-mail informado esta vazio."
    }

    $Email = $Email.Trim()
    $padrao = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    if ($Email -notmatch $padrao) {
        throw "O e-mail informado nao possui um formato valido."
    }

    return $Email
}

function Validar-IdentificadoresUsuario {
    param([object]$Usuario)

    $email = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email","EMail","Mail")
    $userId = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("UserID","UserId","ID","Id")
    $userName = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("UserName","Name","LoginName")

    if ([string]::IsNullOrWhiteSpace($email) -and
        [string]::IsNullOrWhiteSpace($userId) -and
        [string]::IsNullOrWhiteSpace($userName)) {
        throw "O usuario localizado nao possui identificadores validos para inclusao em grupos ou user lists."
    }

    return [PSCustomObject]@{
        Email    = $email
        UserId   = $userId
        UserName = $userName
    }
}

function Descobrir-TipoAcesso {
    param([object]$Acesso)

    $tipoBruto = Obter-ValorSeguroPropriedade -Objeto $Acesso -PossiveisNomes @(
        "Type","ObjectType","MemberType","PrincipalType","AccessType"
    )

    $nomeBruto = Obter-ValorSeguroPropriedade -Objeto $Acesso -PossiveisNomes @(
        "Name","ObjectName","MemberName","UserName"
    )

    $tipoNormalizado = Normalizar-Texto $tipoBruto
    $nomeNormalizado = Normalizar-Texto $nomeBruto

    if ($tipoNormalizado -match 'group') {
        return "Group"
    }

    if ($tipoNormalizado -match 'user\s*list' -or $tipoNormalizado -match 'userlist') {
        return "UserList"
    }

    if ($nomeNormalizado -match 'user\s*list' -or $nomeNormalizado -match 'userlist') {
        return "UserList"
    }

    return $tipoBruto
}

function Localizar-PastaPorPossiveisNomes {
    param(
        [array]$Pastas,
        [string[]]$NomesPossiveis,
        [string]$Descricao = "Pasta"
    )

    if (-not $Pastas -or $Pastas.Count -eq 0) {
        throw "$Descricao nao encontrada. Nenhuma pasta foi retornada para pesquisa."
    }

    $nomesNormalizados = @($NomesPossiveis | ForEach-Object { Normalizar-Texto $_ })

    $encontrada = $Pastas | Where-Object {
        $nome = Normalizar-Texto (Obter-NomeAmigavelItem -Item $_)
        $nomesNormalizados -contains $nome
    } | Select-Object -First 1

    if (-not $encontrada) {
        $encontrada = $Pastas | Where-Object {
            $nome = Normalizar-Texto (Obter-NomeAmigavelItem -Item $_)
            foreach ($possivel in $nomesNormalizados) {
                if ($nome -like "*$possivel*") {
                    return $true
                }
            }
            return $false
        } | Select-Object -First 1
    }

    if (-not $encontrada) {
        $opcoes = @($Pastas | ForEach-Object { Obter-NomeAmigavelItem -Item $_ }) -join ", "
        throw "$Descricao nao encontrada. Nomes esperados: $($NomesPossiveis -join ' / '). Encontradas: $opcoes"
    }

    return $encontrada
}

function Unificar-AcessosSeguranca {
    param(
        [array]$FolderSecurity,
        [array]$TreeAccessObjects,
        [array]$ProjectSecurity
    )

    $todos = @()

    if ($FolderSecurity) {
        $todos += @($FolderSecurity | Where-Object { $_ -ne $null })
    }

    if ($TreeAccessObjects) {
        $todos += @($TreeAccessObjects | Where-Object { $_ -ne $null })
    }

    # NOVO: inclui também a segurança do projeto.
    # Alguns projetos possuem os Groups cadastrados no ProjectWise Administrator
    # em Project Security, e não em Folder Security.
    if ($ProjectSecurity) {
        $todos += @($ProjectSecurity | Where-Object { $_ -ne $null })
    }

    $mapa = @{}

    foreach ($item in $todos) {
        $tipo = Descobrir-TipoAcesso -Acesso $item
        $nome = Obter-ValorSeguroPropriedade -Objeto $item -PossiveisNomes @(
            "Name","ObjectName","MemberName","UserName"
        )

        $tipoNorm = Normalizar-Texto $tipo
        $nomeNorm = Normalizar-Texto $nome

        if ([string]::IsNullOrWhiteSpace($nomeNorm)) {
            continue
        }

        $chave = "$tipoNorm|$nomeNorm"
        if (-not $mapa.ContainsKey($chave)) {
            $mapa[$chave] = $item
        }
    }

    return @($mapa.Values)
}

function Obter-NomeAcesso {
    param([object]$Acesso)

    return Obter-ValorSeguroPropriedade -Objeto $Acesso -PossiveisNomes @(
        "Name","ObjectName","MemberName","UserName"
    )
}


function Remover-AcentosTexto {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return ""
    }

    try {
        # \p{Mn} = NonSpacingMark. Regex nativo e muito mais rapido que iterar
        # caractere a caractere em PowerShell.
        $normalizado = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
        return ([regex]::Replace($normalizado, '\p{Mn}', '')).Normalize([System.Text.NormalizationForm]::FormC)
    }
    catch {
        return $Texto
    }
}

function Normalizar-TextoBuscaGrupo {
    param([string]$Texto)

    $textoNorm = Normalizar-Texto (Remover-AcentosTexto $Texto)
    $textoNorm = $textoNorm -replace '[–—]', '-'
    $textoNorm = $textoNorm -replace '\s+', ' '
    return $textoNorm.Trim()
}

function Normalizar-TextoCompactoBuscaGrupo {
    param([string]$Texto)

    $textoNorm = Normalizar-TextoBuscaGrupo $Texto
    $textoNorm = $textoNorm -replace '[^a-z0-9]+', ''
    return $textoNorm
}

function Obter-NomeGrupoPW {
    param([object]$Grupo)

    if ($null -eq $Grupo) {
        return ""
    }

    if ($Grupo -is [string]) {
        return $Grupo.Trim()
    }

    return Obter-ValorSeguroPropriedade -Objeto $Grupo -PossiveisNomes @(
        "Name", "GroupName", "ObjectName", "MemberName", "UserName"
    )
}

function Obter-TodosGruposPWCache {
    if ($null -ne $script:CacheGruposPW) {
        return @($script:CacheGruposPW)
    }

    $grupos = @()

    try {
        if (Get-Command Get-PWGroups -ErrorAction SilentlyContinue) {
            $grupos = @(Get-PWGroups -ErrorAction Stop | Where-Object { $_ -ne $null })
            Write-Log "Cache de grupos carregado via Get-PWGroups: $($grupos.Count) item(ns)."
        }
        elseif (Get-Command Get-PWGroupNames -ErrorAction SilentlyContinue) {
            $grupos = @(Get-PWGroupNames -ErrorAction Stop | Where-Object { $_ -ne $null })
            Write-Log "Cache de grupos carregado via Get-PWGroupNames: $($grupos.Count) item(ns)."
        }
        else {
            Write-Log "Nenhum cmdlet para listar grupos foi encontrado. Fallback por nome nao sera usado." "WARN"
        }
    }
    catch {
        Write-Log "Falha ao carregar cache de grupos do PW. Detalhe: $($_.Exception.Message)" "WARN"
        $grupos = @()
    }

    $script:CacheGruposPW = @($grupos)
    return @($script:CacheGruposPW)
}

# Indice com os nomes de grupos ja normalizados. A normalizacao (remocao de
# acentos caractere a caractere) e cara; feita uma unica vez por sessao em vez
# de repetida para cada grupo x cada projeto que cai no fallback.
function Obter-IndiceGruposPW {
    if ($null -ne $script:IndiceGruposPW) {
        return $script:IndiceGruposPW
    }

    $inicio = Get-Date
    $indice = New-Object System.Collections.Generic.List[object]
    $rxAcento = New-Object System.Text.RegularExpressions.Regex('\p{Mn}', 'Compiled')
    $rxTraco = New-Object System.Text.RegularExpressions.Regex('[–—]', 'Compiled')
    $rxEspaco = New-Object System.Text.RegularExpressions.Regex('\s+', 'Compiled')
    foreach ($grupo in @(Obter-TodosGruposPWCache)) {
        if ($grupo -is [string]) { $nomeGrupo = $grupo.Trim() }
        elseif ($grupo.PSObject.Properties['Name'] -and $grupo.Name) { $nomeGrupo = ([string]$grupo.Name).Trim() }
        else { $nomeGrupo = Obter-NomeGrupoPW -Grupo $grupo }
        if ([string]::IsNullOrWhiteSpace($nomeGrupo)) { continue }

        # Mesmo resultado de Normalizar-TextoBuscaGrupo, sem o custo de 3 chamadas de funcao por grupo.
        $norm = $rxAcento.Replace($nomeGrupo.Normalize([System.Text.NormalizationForm]::FormD), '').Normalize([System.Text.NormalizationForm]::FormC)
        $norm = $rxEspaco.Replace($rxTraco.Replace($norm.Trim().ToLowerInvariant(), '-'), ' ').Trim()
        $indice.Add([PSCustomObject]@{
            Nome     = $nomeGrupo
            Norm     = $norm
            Compacto = ($norm -replace '[^a-z0-9]+', '')
            Objeto   = $grupo
        })
    }

    $script:IndiceGruposPW = $indice.ToArray()
    Write-Log ("Indice de grupos montado: {0} grupo(s) em {1:N1}s." -f $script:IndiceGruposPW.Count, ((Get-Date) - $inicio).TotalSeconds)
    return $script:IndiceGruposPW
}

function Obter-ApelidosProjetoParaBuscaGrupo {
    param(
        [string]$ProjetoNome,
        [string]$ProjetoNomeTecnico
    )

    $candidatos = New-Object System.Collections.Generic.List[string]

    foreach ($valor in @($ProjetoNome, $ProjetoNomeTecnico)) {
        if ([string]::IsNullOrWhiteSpace($valor)) {
            continue
        }

        $base = Normalizar-TextoBuscaGrupo $valor
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            $candidatos.Add($base)
        }

        # Remove trechos de localização comuns que não aparecem no nome dos grupos.
        $semKm = $base -replace '\s+-\s+km\b.*$', ''
        if ($semKm -ne $base -and -not [string]::IsNullOrWhiteSpace($semKm)) {
            $candidatos.Add($semKm.Trim())
        }

        $semRodovia = $semKm -replace '\s+-\s+(sp|br)[-/ ]?\d+.*$', ''
        if ($semRodovia -ne $semKm -and -not [string]::IsNullOrWhiteSpace($semRodovia)) {
            $candidatos.Add($semRodovia.Trim())
        }

        $semParenteses = $base -replace '\s+-\s+\(.*$', ''
        if ($semParenteses -ne $base -and -not [string]::IsNullOrWhiteSpace($semParenteses)) {
            $candidatos.Add($semParenteses.Trim())
        }

        # Exemplo: "03.11 - HS-WIM 1 - km 20 - SP-280" -> "03.11 - HS-WIM 1".
        $partes = @($base -split '\s+-\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($partes.Count -ge 2) {
            $candidatos.Add(("$($partes[0]) - $($partes[1])").Trim())
            $candidatos.Add(("$($partes[0]) $($partes[1])").Trim())
        }
    }

    $mapa = @{}
    foreach ($c in $candidatos) {
        $norm = Normalizar-TextoBuscaGrupo $c
        $compacto = Normalizar-TextoCompactoBuscaGrupo $norm

        # Evita apelidos genéricos demais.
        if ([string]::IsNullOrWhiteSpace($norm) -or $compacto.Length -lt 6) {
            continue
        }

        if (-not $mapa.ContainsKey($norm)) {
            $mapa[$norm] = $true
        }
    }

    return @($mapa.Keys | Sort-Object { $_.Length } -Descending)
}

function Obter-GruposInferidosPorProjeto {
    param(
        [string]$ProjetoNome,
        [string]$ProjetoNomeTecnico,
        [array]$GroupsAtuais
    )

    $indiceGrupos = @(Obter-IndiceGruposPW)
    if ($indiceGrupos.Count -eq 0) {
        return @()
    }

    $apelidos = Obter-ApelidosProjetoParaBuscaGrupo -ProjetoNome $ProjetoNome -ProjetoNomeTecnico $ProjetoNomeTecnico
    if (-not $apelidos -or $apelidos.Count -eq 0) {
        Write-Log "Fallback de grupos ignorado para '$ProjetoNome': nenhum apelido seguro gerado." "WARN"
        return @()
    }

    $nomesExistentes = @{}
    foreach ($g in @($GroupsAtuais)) {
        if ($null -eq $g) { continue }
        $nomeAtual = ""
        if ($g.PSObject.Properties['Nome']) { $nomeAtual = $g.Nome }
        else { $nomeAtual = Obter-NomeGrupoPW -Grupo $g }
        $nomeNorm = Normalizar-TextoBuscaGrupo $nomeAtual
        if (-not [string]::IsNullOrWhiteSpace($nomeNorm)) { $nomesExistentes[$nomeNorm] = $true }
    }

    # Normaliza os apelidos uma unica vez (antes era feito dentro do laco de grupos).
    $apelidosPreparados = @(
        foreach ($apelido in $apelidos) {
            $apelidoNorm = Normalizar-TextoBuscaGrupo $apelido
            $apelidoCompacto = $apelidoNorm -replace '[^a-z0-9]+', ''
            if ($apelidoCompacto.Length -ge 6) {
                [PSCustomObject]@{ Norm = $apelidoNorm; Compacto = $apelidoCompacto }
            }
        }
    )

    $ordinal = [System.StringComparison]::Ordinal
    $resultado = New-Object System.Collections.Generic.List[object]
    foreach ($g in $indiceGrupos) {
        if ($nomesExistentes.ContainsKey($g.Norm)) { continue }

        $match = $false
        foreach ($a in $apelidosPreparados) {
            if ($g.Norm.IndexOf($a.Norm, $ordinal) -ge 0 -or $g.Compacto.IndexOf($a.Compacto, $ordinal) -ge 0) {
                $match = $true
                break
            }
        }

        if ($match) {
            $resultado.Add([PSCustomObject]@{
                Tipo     = "Group"
                Nome     = $g.Nome
                ObjetoPw = $g.Objeto
                Origem   = "InferidoPorNome"
            })
            $nomesExistentes[$g.Norm] = $true
        }
    }

    $resultado = @($resultado | Sort-Object Nome)

    if ($resultado.Count -gt 0) {
        Write-Log "Fallback por nome encontrou $($resultado.Count) group(s) para projeto '$ProjetoNome'. Apelidos usados: $($apelidos -join ' | ')" "WARN"
        foreach ($g in $resultado) {
            Write-Log "    [GROUP INFERIDO] $($g.Nome)" "WARN"
        }
    }
    else {
        Write-Log "Fallback por nome nao encontrou groups para projeto '$ProjetoNome'. Apelidos usados: $($apelidos -join ' | ')" "WARN"
    }

    return @($resultado)
}

# ---------------------------------------------------------
# FUNCOES DE CONEXAO / MODULO
# ---------------------------------------------------------
function Importar-ModuloProjectWise {
    Write-Log "Tentando importar modulo do ProjectWise."

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
                Write-Log "Modulo '$modulo' importado com sucesso."
                return
            }
        }
        catch {
            Write-Log "Falha ao importar modulo '$modulo'. Detalhe: $($_.Exception.Message)" "WARN"
        }
    }

    throw "Nao foi possivel localizar/importar o modulo do ProjectWise. Verifique se o PWPS_DAB esta instalado."
}

function Conectar-ProjectWise {
    param([ref]$LoginPW)

    Write-Log "Tentando conectar ao datasource do ProjectWise."

    $cmdletsConexao = @(
        "New-PWLogin",
        "Get-PWLogin",
        "Open-PWConnection"
    )

    foreach ($nomeCmdlet in $cmdletsConexao) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        try {
            $params = @($cmd.Parameters.Keys)
            Write-Log "Tentando conexao via '$nomeCmdlet'. Parametros: $($params -join ', ')"

            if ($params -contains "BentleyIMS") {
                $sessao = & $nomeCmdlet -BentleyIMS -ErrorAction Stop
                $LoginPW.Value = $sessao
                Write-Log "Conexao realizada via '$nomeCmdlet -BentleyIMS'."
                return
            }

            if ($params -contains "UseGui") {
                $sessao = & $nomeCmdlet -UseGui -ErrorAction Stop
                $LoginPW.Value = $sessao
                Write-Log "Conexao realizada via '$nomeCmdlet -UseGui'."
                return
            }

            if ($params.Count -eq 0) {
                $sessao = & $nomeCmdlet -ErrorAction Stop
                $LoginPW.Value = $sessao
                Write-Log "Conexao realizada via '$nomeCmdlet' sem parametros."
                return
            }

            $sessao = & $nomeCmdlet -ErrorAction Stop
            $LoginPW.Value = $sessao
            Write-Log "Conexao realizada via '$nomeCmdlet'."
            return
        }
        catch {
            Write-Log "Falha ao conectar via '$nomeCmdlet'. Detalhe: $($_.Exception.Message)" "WARN"
        }
    }

    throw "Nao foi possivel estabelecer conexao com o ProjectWise automaticamente."
}

function Encerrar-SessaoProjectWise {
    param($LoginPW)

    if ($null -eq $LoginPW) {
        return
    }

    $cmdletsDesconexao = @(
        "Undo-PWLogin",
        "Close-PWConnection",
        "Remove-PWLogin"
    )

    foreach ($nomeCmdlet in $cmdletsDesconexao) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        try {
            $params = @($cmd.Parameters.Keys)

            if ($params -contains "Login") {
                & $nomeCmdlet -Login $LoginPW -ErrorAction Stop
                Write-Log "Sessao encerrada com '$nomeCmdlet -Login'."
                return
            }

            if ($params -contains "InputObject") {
                & $nomeCmdlet -InputObject $LoginPW -ErrorAction Stop
                Write-Log "Sessao encerrada com '$nomeCmdlet -InputObject'."
                return
            }

            & $nomeCmdlet -ErrorAction Stop
            Write-Log "Sessao encerrada com '$nomeCmdlet'."
            return
        }
        catch {
            Write-Log "Falha ao encerrar sessao via '$nomeCmdlet'. Detalhe: $($_.Exception.Message)" "WARN"
        }
    }
}

function Inicializar-Aplicacao {
    Write-Log "Aplicacao iniciada."
    Write-UiLog "Iniciando aplicacao..."
    Write-UiLog "Carregando modulo do ProjectWise..."
    Importar-ModuloProjectWise

    Write-UiLog "Conectando ao ProjectWise..."
    Conectar-ProjectWise -LoginPW ([ref]$script:LoginPW)

    if (-not $script:LoginPW) {
        throw "Nao foi possivel concluir o login no ProjectWise."
    }

    Write-UiLog "Conexao com ProjectWise realizada com sucesso."
    Write-UiLog "Log: $ArquivoLog"
    Write-Log "Conexao com ProjectWise realizada com sucesso."
    $script:InicializacaoConcluida = $true
}

# ---------------------------------------------------------
# DIALOGOS VISUAIS
# ---------------------------------------------------------
function Selecionar-ItemDialog {
    param(
        [array]$Itens,
        [string]$Titulo,
        [string]$Descricao = "",
        [scriptblock]$FormatarItem = $null
    )

    if (-not $Itens -or $Itens.Count -eq 0) {
        throw "Nenhum item disponivel para selecao."
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Titulo
    $form.Size = New-Object System.Drawing.Size(520, 520)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Descricao
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(15,15)
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(15,45)
    $listBox.Size = New-Object System.Drawing.Size(470,380)
    $listBox.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($listBox)

    foreach ($item in $Itens) {
        if ($FormatarItem) {
            $nome = [string](& $FormatarItem $item)
        }
        else {
            $nome = Obter-NomeAmigavelItem -Item $item
        }
        $null = $listBox.Items.Add([PSCustomObject]@{
            Texto = $nome
            Valor = $item
        })
    }

    $listBox.DisplayMember = "Texto"

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(310,440)
    $btnOk.Size = New-Object System.Drawing.Size(80,30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(405,440)
    $btnCancelar.Size = New-Object System.Drawing.Size(80,30)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancelar)

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancelar

    $resultado = $form.ShowDialog()

    if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Selecao cancelada pelo operador."
    }

    if (-not $listBox.SelectedItem) {
        throw "Nenhum item foi selecionado."
    }

    return $listBox.SelectedItem.Valor
}

function Mostrar-ResumoConsolidadoDialog {
    param(
        [object]$Usuario,
        [array]$Usuarios,
        [array]$SelecoesProjetos
    )

    if (-not $SelecoesProjetos -or $SelecoesProjetos.Count -eq 0) {
        return
    }

    $texto = ""
    $usuariosResumo = @($Usuarios | Where-Object { $_ -ne $null })
    if ($usuariosResumo.Count -eq 0 -and $Usuario) {
        $usuariosResumo = @($Usuario)
    }

    if ($usuariosResumo.Count -eq 1) {
        $nomeUsuario  = Obter-ValorSeguroPropriedade -Objeto $usuariosResumo[0] -PossiveisNomes @("Name","UserName","FullName")
        $emailUsuario = Obter-ValorSeguroPropriedade -Objeto $usuariosResumo[0] -PossiveisNomes @("Email","EMail","Mail")
        $texto += "Usuario: $nomeUsuario`r`n"
        $texto += "Email  : $emailUsuario`r`n"
    }
    else {
        $texto += "Usuarios selecionados: $($usuariosResumo.Count)`r`n"
        foreach ($usuarioResumo in $usuariosResumo) {
            $nomeUsuario  = Obter-ValorSeguroPropriedade -Objeto $usuarioResumo -PossiveisNomes @("Name","UserName","FullName")
            $emailUsuario = Obter-ValorSeguroPropriedade -Objeto $usuarioResumo -PossiveisNomes @("Email","EMail","Mail")
            $texto += "- $nomeUsuario | $emailUsuario`r`n"
        }
    }

    $texto += "----------------------------------------`r`n"

    $bloco = 1
    foreach ($selecao in $SelecoesProjetos) {
        $texto += "[${bloco}] Concessao: $($selecao.ConcessaoNome)`r`n"
        $texto += "    Projeto : $($selecao.ProjetoNome)`r`n"
        $texto += "    Acessos :`r`n"

        foreach ($acesso in $selecao.AcessosSelecionados) {
            $texto += "      - $($acesso.Tipo) | $($acesso.Nome)`r`n"
        }

        $texto += "`r`n"
        $bloco++
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Resumo das seleções"
    $form.Size = New-Object System.Drawing.Size(720, 520)
    $form.StartPosition = "CenterParent"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Font = New-Object System.Drawing.Font("Consolas",10)
    $txt.Location = New-Object System.Drawing.Point(15,15)
    $txt.Size = New-Object System.Drawing.Size(675,420)
    $txt.Text = $texto
    $form.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Fechar"
    $btnOk.Location = New-Object System.Drawing.Point(610,445)
    $btnOk.Size = New-Object System.Drawing.Size(80,30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $form.AcceptButton = $btnOk
    [void]$form.ShowDialog()
}

# ---------------------------------------------------------
# EXECUCAO - COMANDOS PW
# ---------------------------------------------------------
function Invocar-BuscaUsuariosPW {
    param(
        [string]$Valor,
        [string]$Parametro
    )

    if ($Parametro) {
        $argumentos = @{ $Parametro = $Valor; ErrorAction = 'SilentlyContinue' }
        return @(Get-PWUsersByMatch @argumentos | Where-Object { $_ -ne $null })
    }

    return @(Get-PWUsersByMatch $Valor -ErrorAction SilentlyContinue | Where-Object { $_ -ne $null })
}

# Resolve um e-mail sem lancar excecao. Retorna um objeto com Status:
#   Encontrado  -> um unico usuario com e-mail exato
#   Parcial     -> um unico usuario retornado, mas o e-mail dele e diferente do informado
#   Ambiguo     -> mais de um usuario possivel (Candidatos preenchido)
#   NaoEncontrado / Erro
function Resolver-UsuarioPW {
    param([string]$Email)

    $resultado = [PSCustomObject]@{
        Email      = $Email
        Status     = ""
        Usuario    = $null
        Candidatos = @()
        Detalhe    = ""
    }

    try {
        Write-Log "Buscando usuario pelo email: $Email"

        if (-not $script:ParametrosGetPWUsersByMatch) {
            $script:ParametrosGetPWUsersByMatch = @(Obter-ParametrosCmdlet -NomeCmdlet "Get-PWUsersByMatch")
            Write-Log ("Parametros encontrados no cmdlet: " + ($script:ParametrosGetPWUsersByMatch -join ", "))
        }
        $parametros = $script:ParametrosGetPWUsersByMatch

        $parametroPrincipal = @("Email", "Identity", "UserName", "UserId") | Where-Object { $parametros -contains $_ } | Select-Object -First 1
        $usuarios = @(Invocar-BuscaUsuariosPW -Valor $Email -Parametro $parametroPrincipal)

        # Alguns usuarios tem o e-mail como login, mas o campo Email vazio no PW.
        if ($usuarios.Count -eq 0 -and $parametroPrincipal -ne "UserName" -and $parametros -contains "UserName") {
            $usuarios = @(Invocar-BuscaUsuariosPW -Valor $Email -Parametro "UserName")
            if ($usuarios.Count -gt 0) {
                Write-Log "Usuario '$Email' localizado pelo login (UserName) em vez do campo Email."
            }
        }

        if ($usuarios.Count -eq 0) {
            $resultado.Status = "NaoEncontrado"
            $resultado.Detalhe = "Nenhum usuario no ProjectWise com este e-mail. Verifique a grafia ou crie o usuario (rotina 01)."
            Write-Log "Usuario NAO encontrado: $Email" "WARN"
            return $resultado
        }

        $comEmailExato = @(
            $usuarios | Where-Object {
                $emailItem = Obter-ValorSeguroPropriedade -Objeto $_ -PossiveisNomes @("Email","EMail","Mail")
                $loginItem = Obter-ValorSeguroPropriedade -Objeto $_ -PossiveisNomes @("UserName","LoginName","Name")
                ($emailItem -and $emailItem -ieq $Email) -or ($loginItem -and $loginItem -ieq $Email)
            }
        )

        if ($comEmailExato.Count -eq 1) {
            Write-Log "Usuario localizado com correspondencia exata de email."
            $resultado.Status = "Encontrado"
            $resultado.Usuario = $comEmailExato[0]
            return $resultado
        }

        if ($comEmailExato.Count -gt 1) {
            $resultado.Status = "Ambiguo"
            $resultado.Candidatos = $comEmailExato
            $resultado.Detalhe = "$($comEmailExato.Count) usuarios com este mesmo e-mail no PW. Escolha o correto."
            Write-Log "Usuario AMBIGUO: $Email ($($comEmailExato.Count) contas com o mesmo e-mail): $((@($comEmailExato | ForEach-Object { Obter-RotuloUsuario -Usuario $_ })) -join ' || ')" "WARN"
            return $resultado
        }

        if ($usuarios.Count -eq 1) {
            $emailRetornado = Obter-ValorSeguroPropriedade -Objeto $usuarios[0] -PossiveisNomes @("Email","EMail","Mail")
            $resultado.Status = "Parcial"
            $resultado.Usuario = $usuarios[0]
            $resultado.Detalhe = "A busca retornou um usuario com e-mail diferente ('$emailRetornado'). Confira antes de incluir."
            Write-Log "Usuario com correspondencia PARCIAL: informado '$Email', retornado '$emailRetornado'." "WARN"
            return $resultado
        }

        $resultado.Status = "Ambiguo"
        $resultado.Candidatos = $usuarios
        $resultado.Detalhe = "$($usuarios.Count) usuarios parecidos, nenhum com e-mail exato. Escolha o correto."
        Write-Log "Usuario AMBIGUO: $Email ($($usuarios.Count) resultados, nenhum exato)." "WARN"
        return $resultado
    }
    catch {
        $resultado.Status = "Erro"
        $resultado.Detalhe = "Falha na consulta ao PW: $($_.Exception.Message)"
        Write-Log "Erro ao buscar usuario '$Email'. Detalhe: $($_.Exception.Message)" "ERROR"
        return $resultado
    }
}

function Buscar-UsuarioPW {
    param([string]$Email)

    $r = Resolver-UsuarioPW -Email $Email
    switch ($r.Status) {
        "Encontrado" { return $r.Usuario }
        "Parcial"    { return $r.Usuario }
        default      { throw "Usuario '$Email': $($r.Detalhe)" }
    }
}

function Get-PWRootFoldersSafe {
    Write-Log "Listando pastas da raiz por Root."
    $pastas = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    $pastas = @($pastas | Where-Object { $_ -ne $null })

    if ($pastas.Count -eq 0) {
        throw "Nenhuma pasta retornada na raiz."
    }

    return $pastas
}

function Get-PWChildFoldersSafe {
    param([string]$FolderId)

    Write-Log "Listando filhos da pasta ID: $FolderId"
    $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
    $pastas = @($pastas | Where-Object { $_ -ne $null })

    if ($pastas.Count -eq 0) {
        throw "Nenhuma subpasta retornada para o FolderID $FolderId."
    }

    return $pastas
}

function Get-PWChildFoldersOrEmpty {
    param([string]$FolderId)

    try {
        $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
        return @($pastas | Where-Object { $_ -ne $null })
    }
    catch {
        Write-Log "Nao foi possivel listar filhos do FolderID $FolderId. Detalhe: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Obter-ProjetosDaPastaProjetos {
    param(
        [string]$FolderIdProjetos,
        [int]$ProfundidadeMaxima = 2
    )

    $resultado = @()
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $FolderIdProjetos; Nivel = 0 })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = Get-PWChildFoldersOrEmpty -FolderId $atual.FolderId

        foreach ($filho in $filhos) {
            $idFilho = Obter-IdPasta -Pasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho)) { continue }
            $resultado += $filho
            if ($atual.Nivel -lt $ProfundidadeMaxima) {
                $netos = Get-PWChildFoldersOrEmpty -FolderId $idFilho
                if ($netos.Count -gt 0) {
                    $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nivel = ($atual.Nivel + 1) })
                }
            }
        }
    }

    $mapa = @{}
    foreach ($item in $resultado) {
        $id = Obter-IdPasta -Pasta $item
        if (-not [string]::IsNullOrWhiteSpace($id) -and -not $mapa.ContainsKey($id)) { $mapa[$id] = $item }
    }
    return @(Ordenar-ItensPorNome -Itens @($mapa.Values))
}

function Ler-SegurancaProjeto {
    param([object]$PastaProjeto)

    $resultado = [ordered]@{
        FolderSecurity    = @()
        TreeAccessObjects = @()
        ProjectSecurity   = @()
        TodosAcessos      = @()
    }

    $nomeProjeto = Obter-ValorSeguroPropriedade -Objeto $PastaProjeto -PossiveisNomes @("Name","FolderName","ProjectName")
    $idProjeto = Obter-IdPasta -Pasta $PastaProjeto

    Write-Log "Iniciando leitura da seguranca do projeto '$nomeProjeto' (ID: $idProjeto)"

    try {
        $parametros = Obter-ParametrosCmdlet -NomeCmdlet "Get-PWFolderSecurity"

        if ($parametros -contains "InputFolder") {
            $objetos = @(Get-PWFolderSecurity -InputFolder $PastaProjeto -ErrorAction Stop)
        }
        else {
            $objetos = @(Get-PWFolderSecurity $PastaProjeto -ErrorAction Stop)
        }

        $resultado["FolderSecurity"] = @($objetos | Where-Object { $_ -ne $null })
        Write-Log "Get-PWFolderSecurity retornou $($resultado['FolderSecurity'].Count) item(ns)."
    }
    catch {
        Write-Log "Falha em Get-PWFolderSecurity. Detalhe: $($_.Exception.Message)" "WARN"
    }

    try {
        $parametros = Obter-ParametrosCmdlet -NomeCmdlet "Get-PWFolderTreeAccessObjects"

        if ($parametros -contains "InputFolders") {
            $objetos = @(Get-PWFolderTreeAccessObjects -InputFolders @($PastaProjeto) -NoSubfolders -ErrorAction Stop)
        }
        else {
            $objetos = @(Get-PWFolderTreeAccessObjects @($PastaProjeto) -NoSubfolders -ErrorAction Stop)
        }

        $resultado["TreeAccessObjects"] = @($objetos | Where-Object { $_ -ne $null })
        Write-Log "Get-PWFolderTreeAccessObjects retornou $($resultado['TreeAccessObjects'].Count) item(ns)."
    }
    catch {
        Write-Log "Falha em Get-PWFolderTreeAccessObjects. Detalhe: $($_.Exception.Message)" "WARN"
    }

    # NOVO: leitura da segurança de projeto.
    # Importante para projetos em que os grupos existem no ProjectWise Administrator,
    # mas não aparecem na segurança da pasta raiz do projeto.
    try {
        $cmdProjectSecurity = Get-Command Get-PWProjectSecurity -ErrorAction SilentlyContinue

        if ($cmdProjectSecurity) {
            $parametros = @($cmdProjectSecurity.Parameters.Keys)
            $objetos = @()

            if ($parametros -contains "InputProject") {
                $objetos = @(Get-PWProjectSecurity -InputProject $PastaProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "InputFolder") {
                $objetos = @(Get-PWProjectSecurity -InputFolder $PastaProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "InputObject") {
                $objetos = @(Get-PWProjectSecurity -InputObject $PastaProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "ProjectID") {
                $objetos = @(Get-PWProjectSecurity -ProjectID $idProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "ProjectId") {
                $objetos = @(Get-PWProjectSecurity -ProjectId $idProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "FolderID") {
                $objetos = @(Get-PWProjectSecurity -FolderID $idProjeto -ErrorAction Stop)
            }
            elseif ($parametros -contains "FolderId") {
                $objetos = @(Get-PWProjectSecurity -FolderId $idProjeto -ErrorAction Stop)
            }
            else {
                $objetos = @(Get-PWProjectSecurity $PastaProjeto -ErrorAction Stop)
            }

            $resultado["ProjectSecurity"] = @($objetos | Where-Object { $_ -ne $null })
            Write-Log "Get-PWProjectSecurity retornou $($resultado['ProjectSecurity'].Count) item(ns)."
        }
        else {
            Write-Log "Cmdlet Get-PWProjectSecurity nao encontrado no ambiente. A leitura seguira apenas com FolderSecurity/TreeAccessObjects." "WARN"
        }
    }
    catch {
        Write-Log "Falha em Get-PWProjectSecurity. Detalhe: $($_.Exception.Message)" "WARN"
    }

    $resultado["TodosAcessos"] = Unificar-AcessosSeguranca `
        -FolderSecurity $resultado["FolderSecurity"] `
        -TreeAccessObjects $resultado["TreeAccessObjects"] `
        -ProjectSecurity $resultado["ProjectSecurity"]

    Write-Log "Total consolidado de acessos encontrados: $($resultado['TodosAcessos'].Count)"
    return $resultado
}

function Filtrar-AcessosProjeto {
    param([array]$ObjetosSeguranca)

    $filtrados = @(
        $ObjetosSeguranca | Where-Object {
            $_ -ne $null
        }
    )

    $mapa = @{}

    foreach ($item in $filtrados) {
        $tipo = Descobrir-TipoAcesso -Acesso $item
        $nome = Obter-NomeAcesso -Acesso $item

        if ([string]::IsNullOrWhiteSpace($nome)) {
            continue
        }

        $chave = "$(Normalizar-Texto $tipo)|$(Normalizar-Texto $nome)"

        if (-not $mapa.ContainsKey($chave)) {
            $mapa[$chave] = $item
        }
    }

    return @(
        $mapa.Values |
        Sort-Object `
            @{ Expression = { (Descobrir-TipoAcesso -Acesso $_).ToLower() } },
            @{ Expression = { (Obter-NomeAcesso -Acesso $_).ToLower() } }
    )
}

function Adicionar-Usuario-Em-Grupo {
    param(
        [object]$Usuario,
        [string]$NomeGrupo
    )

    $ids = Validar-IdentificadoresUsuario -Usuario $Usuario

    $email = $ids.Email
    $userId = $ids.UserId
    $userName = $ids.UserName

    $cmd = Get-Command Add-PWUserToGroup -ErrorAction SilentlyContinue
    if ($cmd) {
        $params = @($cmd.Parameters.Keys)

        if (($params -contains "GroupName") -and ($params -contains "UserName") -and -not [string]::IsNullOrWhiteSpace($userName)) {
            Add-PWUserToGroup -GroupName $NomeGrupo -UserName $userName -ErrorAction Stop
            return "Add-PWUserToGroup(GroupName/UserName)"
        }
        elseif (($params -contains "GroupName") -and ($params -contains "UserID") -and -not [string]::IsNullOrWhiteSpace($userId)) {
            Add-PWUserToGroup -GroupName $NomeGrupo -UserID $userId -ErrorAction Stop
            return "Add-PWUserToGroup(GroupName/UserID)"
        }
        elseif (($params -contains "GroupName") -and ($params -contains "Email") -and -not [string]::IsNullOrWhiteSpace($email)) {
            Add-PWUserToGroup -GroupName $NomeGrupo -Email $email -ErrorAction Stop
            return "Add-PWUserToGroup(GroupName/Email)"
        }
    }

    $cmd = Get-Command Add-PWGroupMember -ErrorAction SilentlyContinue
    if ($cmd) {
        $params = @($cmd.Parameters.Keys)

        if (($params -contains "GroupName") -and ($params -contains "UserName") -and -not [string]::IsNullOrWhiteSpace($userName)) {
            Add-PWGroupMember -GroupName $NomeGrupo -UserName $userName -ErrorAction Stop
            return "Add-PWGroupMember(GroupName/UserName)"
        }
        elseif (($params -contains "GroupName") -and ($params -contains "UserID") -and -not [string]::IsNullOrWhiteSpace($userId)) {
            Add-PWGroupMember -GroupName $NomeGrupo -UserID $userId -ErrorAction Stop
            return "Add-PWGroupMember(GroupName/UserID)"
        }
        elseif (($params -contains "GroupName") -and ($params -contains "Email") -and -not [string]::IsNullOrWhiteSpace($email)) {
            Add-PWGroupMember -GroupName $NomeGrupo -Email $email -ErrorAction Stop
            return "Add-PWGroupMember(GroupName/Email)"
        }
    }

    throw "Nenhum cmdlet compativel para inclusao em grupo foi encontrado ou os parametros nao bateram."
}

function Adicionar-Usuario-Em-UserList {
    param(
        [object]$Usuario,
        [string]$NomeUserList
    )

    $ids = Validar-IdentificadoresUsuario -Usuario $Usuario

    $email = $ids.Email
    $userId = $ids.UserId
    $userName = $ids.UserName

    $cmdletsPossiveis = @(
        "Add-PWUserListMember",
        "Add-PWMemberToUserList",
        "Add-PWUserToUserList"
    )

    foreach ($nomeCmdlet in $cmdletsPossiveis) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue

        if (-not $cmd) {
            Write-Log "Cmdlet '$nomeCmdlet' nao encontrado no ambiente." "WARN"
            continue
        }

        $params = @($cmd.Parameters.Keys)

        try {
            if (($params -contains "UserListName") -and ($params -contains "UserName") -and -not [string]::IsNullOrWhiteSpace($userName)) {
                & $nomeCmdlet -UserListName $NomeUserList -UserName $userName -ErrorAction Stop
                return "${nomeCmdlet}(UserListName/UserName)"
            }

            if (($params -contains "UserListName") -and ($params -contains "UserID") -and -not [string]::IsNullOrWhiteSpace($userId)) {
                & $nomeCmdlet -UserListName $NomeUserList -UserID $userId -ErrorAction Stop
                return "${nomeCmdlet}(UserListName/UserID)"
            }

            if (($params -contains "UserListName") -and ($params -contains "Email") -and -not [string]::IsNullOrWhiteSpace($email)) {
                & $nomeCmdlet -UserListName $NomeUserList -Email $email -ErrorAction Stop
                return "${nomeCmdlet}(UserListName/Email)"
            }

            if (($params -contains "Name") -and ($params -contains "UserName") -and -not [string]::IsNullOrWhiteSpace($userName)) {
                & $nomeCmdlet -Name $NomeUserList -UserName $userName -ErrorAction Stop
                return "${nomeCmdlet}(Name/UserName)"
            }

            if (($params -contains "Name") -and ($params -contains "UserID") -and -not [string]::IsNullOrWhiteSpace($userId)) {
                & $nomeCmdlet -Name $NomeUserList -UserID $userId -ErrorAction Stop
                return "${nomeCmdlet}(Name/UserID)"
            }

            if (($params -contains "Name") -and ($params -contains "Email") -and -not [string]::IsNullOrWhiteSpace($email)) {
                & $nomeCmdlet -Name $NomeUserList -Email $email -ErrorAction Stop
                return "${nomeCmdlet}(Name/Email)"
            }

            if (($params -contains "UserList") -and ($params -contains "UserName") -and -not [string]::IsNullOrWhiteSpace($userName)) {
                & $nomeCmdlet -UserList $NomeUserList -UserName $userName -ErrorAction Stop
                return "${nomeCmdlet}(UserList/UserName)"
            }

            if (($params -contains "UserList") -and ($params -contains "UserID") -and -not [string]::IsNullOrWhiteSpace($userId)) {
                & $nomeCmdlet -UserList $NomeUserList -UserID $userId -ErrorAction Stop
                return "${nomeCmdlet}(UserList/UserID)"
            }

            if (($params -contains "UserList") -and ($params -contains "Email") -and -not [string]::IsNullOrWhiteSpace($email)) {
                & $nomeCmdlet -UserList $NomeUserList -Email $email -ErrorAction Stop
                return "${nomeCmdlet}(UserList/Email)"
            }
        }
        catch {
            Write-Log "Tentativa com '$nomeCmdlet' falhou. Detalhe: $($_.Exception.Message)" "WARN"
        }
    }

    throw "Nenhum cmdlet compativel para inclusao em user list foi encontrado ou os parametros nao bateram."
}

function Classificar-ResultadoExecucao {
    param([string]$MensagemErro)

    $mensagem = ($MensagemErro | Out-String).Trim()

    if ($mensagem -match "Unique constraint violated" -or
        $mensagem -match "already exists" -or
        $mensagem -match "already a member" -or
        $mensagem -match "duplicate") {

        return [PSCustomObject]@{
            Status  = "JA_EXISTE"
            Detalhe = "Usuario ja pertence a este acesso."
        }
    }

    return [PSCustomObject]@{
        Status  = "ERRO"
        Detalhe = $mensagem
    }
}

function Executar-InclusoesSelecionadas {
    param(
        [object]$Usuario,
        [array]$SelecoesProjetos
    )

    $resultados = @()
    $usuarioNome = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Name","UserName","FullName")
    $usuarioEmail = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email","EMail","Mail")

    foreach ($selecao in $SelecoesProjetos) {
        foreach ($acesso in $selecao.AcessosSelecionados) {
            $tipo = $acesso.Tipo
            $nome = $acesso.Nome

            try {
                $metodo = ""

                if ($tipo -eq "Group") {
                    $metodo = Adicionar-Usuario-Em-Grupo -Usuario $Usuario -NomeGrupo $nome
                }
                elseif ($tipo -eq "UserList") {
                    $metodo = Adicionar-Usuario-Em-UserList -Usuario $Usuario -NomeUserList $nome
                }
                else {
                    throw "Tipo de acesso nao suportado para execucao: $tipo"
                }

                $resultados += [PSCustomObject]@{
                    Usuario   = $usuarioNome
                    Email     = $usuarioEmail
                    Concessao = $selecao.ConcessaoNome
                    Projeto   = $selecao.ProjetoNome
                    Tipo      = $tipo
                    Nome      = $nome
                    Status    = "SUCESSO"
                    Detalhe   = $metodo
                }
            }
            catch {
                $classificacao = Classificar-ResultadoExecucao -MensagemErro $_.Exception.Message

                $resultados += [PSCustomObject]@{
                    Usuario   = $usuarioNome
                    Email     = $usuarioEmail
                    Concessao = $selecao.ConcessaoNome
                    Projeto   = $selecao.ProjetoNome
                    Tipo      = $tipo
                    Nome      = $nome
                    Status    = $classificacao.Status
                    Detalhe   = $classificacao.Detalhe
                }
            }
        }
    }

    return $resultados
}

function Mostrar-ResultadoExecucaoDialog {
    param([array]$Resultados)

    $texto = ""

    if (-not $Resultados -or $Resultados.Count -eq 0) {
        $texto = "Nenhuma execucao foi realizada."
    }
    else {
        $i = 1
        foreach ($resultado in $Resultados) {
            $usuarioPrefixo = ""
            if ($resultado.PSObject.Properties["Usuario"]) {
                $usuarioPrefixo = "$($resultado.Usuario) | "
            }

            $texto += "[${i}] $($resultado.Status) | ${usuarioPrefixo}$($resultado.Concessao) | $($resultado.Projeto) | $($resultado.Tipo) | $($resultado.Nome)`r`n"
            $texto += "    Detalhe: $($resultado.Detalhe)`r`n`r`n"
            $i++
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Resultado final da execução"
    $form.Size = New-Object System.Drawing.Size(860, 560)
    $form.StartPosition = "CenterParent"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Font = New-Object System.Drawing.Font("Consolas",10)
    $txt.Location = New-Object System.Drawing.Point(15,15)
    $txt.Size = New-Object System.Drawing.Size(810,450)
    $txt.Text = $texto
    $form.Controls.Add($txt)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Fechar"
    $btn.Location = New-Object System.Drawing.Point(745,475)
    $btn.Size = New-Object System.Drawing.Size(80,30)
    $btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btn)

    $form.AcceptButton = $btn
    [void]$form.ShowDialog()
}

function Mostrar-ResumoFinalExecucaoDialog {
    param([array]$Resultados)

    if (-not $Resultados -or $Resultados.Count -eq 0) {
        return
    }

    $total = $Resultados.Count
    $sucesso = @($Resultados | Where-Object { $_.Status -eq "SUCESSO" }).Count
    $jaExiste = @($Resultados | Where-Object { $_.Status -eq "JA_EXISTE" }).Count
    $erro = @($Resultados | Where-Object { $_.Status -eq "ERRO" }).Count

    $texto = ""
    $texto += "Total processado : $total`r`n"
    $texto += "Sucessos         : $sucesso`r`n"
    $texto += "Ja existentes    : $jaExiste`r`n"
    $texto += "Erros            : $erro`r`n"
    $texto += "----------------------------------------`r`n`r`n"

    $temUsuarioResultado = @($Resultados | Where-Object { $_.PSObject.Properties["Usuario"] }).Count -gt 0
    if ($temUsuarioResultado) {
        $gruposUsuario = $Resultados | Group-Object Usuario
        foreach ($grupoUsuario in $gruposUsuario) {
            $texto += "Usuario: $($grupoUsuario.Name)`r`n"
            $gruposProjeto = $grupoUsuario.Group | Group-Object Projeto
            foreach ($grupoProjeto in $gruposProjeto) {
                $texto += "  Projeto: $($grupoProjeto.Name)`r`n"
                foreach ($item in $grupoProjeto.Group) {
                    $texto += "  - $($item.Status) | $($item.Tipo) | $($item.Nome)`r`n"
                }
            }
            $texto += "`r`n"
        }
    }
    else {
        $gruposProjeto = $Resultados | Group-Object Projeto
        foreach ($grupoProjeto in $gruposProjeto) {
            $texto += "Projeto: $($grupoProjeto.Name)`r`n"
            foreach ($item in $grupoProjeto.Group) {
                $texto += "- $($item.Status) | $($item.Tipo) | $($item.Nome)`r`n"
            }
            $texto += "`r`n"
        }
    }

    Write-Log "Resumo final da execução | Total=$total | Sucesso=$sucesso | JaExiste=$jaExiste | Erro=$erro"

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Resumo final da execução"
    $form.Size = New-Object System.Drawing.Size(760, 520)
    $form.StartPosition = "CenterParent"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Font = New-Object System.Drawing.Font("Consolas",10)
    $txt.Location = New-Object System.Drawing.Point(15,15)
    $txt.Size = New-Object System.Drawing.Size(710,420)
    $txt.Text = $texto
    $form.Controls.Add($txt)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Fechar"
    $btn.Location = New-Object System.Drawing.Point(645,445)
    $btn.Size = New-Object System.Drawing.Size(80,30)
    $btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btn)

    $form.AcceptButton = $btn
    [void]$form.ShowDialog()
}

# ---------------------------------------------------------
# FLUXO VISUAL / ARVORE
# ---------------------------------------------------------
function Limpar-ArvoreProjetos {
    $script:TreeProjetos.Nodes.Clear()
}

function Obter-ChaveAcessoProjeto {
    param(
        [string]$ProjetoNome,
        [string]$Tipo,
        [string]$Nome
    )

    return "{0}|{1}|{2}" -f (Normalizar-Texto $ProjetoNome), (Normalizar-Texto $Tipo), (Normalizar-Texto $Nome)
}

function Obter-ChaveProjetoSelecao {
    param(
        [string]$ConcessaoNome,
        [string]$ProjetoNome
    )

    return "{0}|{1}" -f (Normalizar-Texto $ConcessaoNome), (Normalizar-Texto $ProjetoNome)
}

function Obter-MapaAcessosSelecoesConfirmadas {
    $mapa = @{}

    foreach ($selecao in @($script:SelecoesProjetos)) {
        if ($null -eq $selecao) {
            continue
        }

        foreach ($acesso in @($selecao.AcessosSelecionados)) {
            if ($null -eq $acesso) {
                continue
            }

            $chave = Obter-ChaveAcessoProjeto -ProjetoNome $selecao.ProjetoNome -Tipo $acesso.Tipo -Nome $acesso.Nome
            $mapa[$chave] = $true
        }
    }

    return $mapa
}

function Obter-MapaAcessosMarcados {
    $mapa = Obter-MapaAcessosSelecoesConfirmadas

    if (-not $script:TreeProjetos) {
        return $mapa
    }

    foreach ($projetoNode in $script:TreeProjetos.Nodes) {
        $tagProjeto = $projetoNode.Tag
        if ($null -eq $tagProjeto -or $tagProjeto.NodeType -ne 'Project') {
            continue
        }

        foreach ($categoriaNode in $projetoNode.Nodes) {
            foreach ($acessoNode in $categoriaNode.Nodes) {
                if (-not $acessoNode.Checked) {
                    continue
                }

                $tagAcesso = $acessoNode.Tag
                if ($null -eq $tagAcesso -or $tagAcesso.NodeType -ne 'Access') {
                    continue
                }

                $chave = Obter-ChaveAcessoProjeto -ProjetoNome $tagProjeto.ProjetoNome -Tipo $tagAcesso.Tipo -Nome $tagAcesso.Nome
                $mapa[$chave] = $true
            }
        }
    }

    return $mapa
}

function Atualizar-CheckVisualProjeto {
    param([System.Windows.Forms.TreeNode]$ProjetoNode)

    if ($null -eq $ProjetoNode) {
        return
    }

    $projetoMarcado = $false

    foreach ($categoriaNode in $ProjetoNode.Nodes) {
        $categoriaMarcada = $false
        $tagCategoria = $categoriaNode.Tag

        foreach ($acessoNode in $categoriaNode.Nodes) {
            $tagAcesso = $acessoNode.Tag

            if ($null -eq $tagAcesso -or $tagAcesso.NodeType -ne 'Access') {
                $acessoNode.Checked = $false
                continue
            }

            if ($acessoNode.Checked) {
                $categoriaMarcada = $true
                $projetoMarcado = $true
                break
            }
        }

        if ($tagCategoria -and ($tagCategoria.NodeType -eq 'CategoryInfo' -or ($tagCategoria.PSObject.Properties['SomenteVisual'] -and $tagCategoria.SomenteVisual))) {
            $categoriaNode.Checked = $false
        }
        else {
            $categoriaNode.Checked = $categoriaMarcada
        }
    }

    $ProjetoNode.Checked = $projetoMarcado
}

function Renderizar-ArvoreProjetos {
    param(
        [array]$ProjetosData,
        [string]$Filtro = "",
        [string]$ExcluirFiltro = "",
        [bool]$TermoExato = $false,
        [hashtable]$MapaMarcados = $null
    )

    if ($null -eq $MapaMarcados) {
        $MapaMarcados = Obter-MapaAcessosMarcados
    }

    $filtroNormalizado = Normalizar-Texto $Filtro
    $excluirNormalizado = Normalizar-Texto $ExcluirFiltro

    $script:AtualizandoCheckTree = $true
    try {
        Limpar-ArvoreProjetos

        foreach ($projetoData in $ProjetosData) {
            $nomeProjeto = $projetoData.ProjetoNome
            $projetoCombina = $false
            $nomeProjetoNormalizado = Normalizar-Texto $nomeProjeto

            if (-not [string]::IsNullOrWhiteSpace($excluirNormalizado) -and $nomeProjetoNormalizado.Contains($excluirNormalizado)) {
                continue
            }

            if ([string]::IsNullOrWhiteSpace($filtroNormalizado)) {
                $projetoCombina = $true
            }
            else {
                if ($TermoExato) {
                    $projetoCombina = $nomeProjetoNormalizado.Contains($filtroNormalizado)
                }
                else {
                    $projetoCombina = $nomeProjetoNormalizado -like "*$filtroNormalizado*"
                }
            }

            if ($projetoCombina) {
                $gruposFiltrados = @($projetoData.Groups)
                $listasFiltradas = @($projetoData.UserLists)
                $outrosFiltrados = @($projetoData.OutrosAcessos)
            }
            else {
                if ($TermoExato) {
                    $gruposFiltrados = @($projetoData.Groups | Where-Object {
                        (Normalizar-Texto $_.Nome).Contains($filtroNormalizado)
                    })
                    $listasFiltradas = @($projetoData.UserLists | Where-Object {
                        (Normalizar-Texto $_.Nome).Contains($filtroNormalizado)
                    })
                    $outrosFiltrados = @($projetoData.OutrosAcessos | Where-Object {
                        (Normalizar-Texto $_.Nome).Contains($filtroNormalizado) -or
                        (Normalizar-Texto $_.Tipo).Contains($filtroNormalizado)
                    })
                }
                else {
                    $gruposFiltrados = @($projetoData.Groups | Where-Object {
                        (Normalizar-Texto $_.Nome) -like "*$filtroNormalizado*"
                    })
                    $listasFiltradas = @($projetoData.UserLists | Where-Object {
                        (Normalizar-Texto $_.Nome) -like "*$filtroNormalizado*"
                    })
                    $outrosFiltrados = @($projetoData.OutrosAcessos | Where-Object {
                        (Normalizar-Texto $_.Nome) -like "*$filtroNormalizado*" -or
                        (Normalizar-Texto $_.Tipo) -like "*$filtroNormalizado*"
                    })
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($excluirNormalizado)) {
                $gruposFiltrados = @($gruposFiltrados | Where-Object {
                    -not (Normalizar-Texto $_.Nome).Contains($excluirNormalizado)
                })
                $listasFiltradas = @($listasFiltradas | Where-Object {
                    -not (Normalizar-Texto $_.Nome).Contains($excluirNormalizado)
                })
                $outrosFiltrados = @($outrosFiltrados | Where-Object {
                    -not (Normalizar-Texto $_.Nome).Contains($excluirNormalizado) -and
                    -not (Normalizar-Texto $_.Tipo).Contains($excluirNormalizado)
                })
            }

            if (($gruposFiltrados.Count + $listasFiltradas.Count + $outrosFiltrados.Count) -eq 0) {
                continue
            }

            $projectNode = New-Object System.Windows.Forms.TreeNode($nomeProjeto)
            $projectNode.Tag = [PSCustomObject]@{
                NodeType        = 'Project'
                ConcessaoNome   = $projetoData.ConcessaoNome
                ConcessaoObjeto  = $projetoData.ConcessaoObjeto
                ProjetoNome     = $projetoData.ProjetoNome
                ProjetoObjeto   = $projetoData.ProjetoObjeto
                ProjetoId       = $projetoData.ProjetoId
            }

            if ($gruposFiltrados.Count -gt 0) {
                $qtdInferidosGrupo = @($gruposFiltrados | Where-Object { $_.PSObject.Properties['Origem'] -and $_.Origem -eq 'InferidoPorNome' }).Count
                if ($qtdInferidosGrupo -gt 0) {
                    $groupNode = New-Object System.Windows.Forms.TreeNode(("Groups ({0}) - inclui {1} inferido(s)" -f $gruposFiltrados.Count, $qtdInferidosGrupo))
                    $groupNode.ForeColor = [System.Drawing.Color]::DarkOrange
                }
                else {
                    $groupNode = New-Object System.Windows.Forms.TreeNode(("Groups ({0})" -f $gruposFiltrados.Count))
                }
                $groupNode.Tag = [PSCustomObject]@{
                    NodeType    = 'Category'
                    Category    = 'Group'
                    ProjetoNome = $projetoData.ProjetoNome
                }

                foreach ($acesso in $gruposFiltrados) {
                    $origemAcesso = "PW"
                    if ($acesso.PSObject.Properties['Origem'] -and -not [string]::IsNullOrWhiteSpace($acesso.Origem)) {
                        $origemAcesso = $acesso.Origem
                    }

                    $textoAcesso = $acesso.Nome
                    if ($origemAcesso -eq "InferidoPorNome") {
                        $textoAcesso = $acesso.Nome
                    }

                    $accessNode = New-Object System.Windows.Forms.TreeNode($textoAcesso)
                    if ($origemAcesso -eq "InferidoPorNome") {
                        $accessNode.ForeColor = [System.Drawing.Color]::DarkOrange
                    }
                    $accessNode.Tag = [PSCustomObject]@{
                        NodeType    = 'Access'
                        Tipo        = $acesso.Tipo
                        Nome        = $acesso.Nome
                        ObjetoPw    = $acesso.ObjetoPw
                        ProjetoNome = $projetoData.ProjetoNome
                        Origem      = $origemAcesso
                    }

                    $chaveAcesso = Obter-ChaveAcessoProjeto -ProjetoNome $projetoData.ProjetoNome -Tipo $acesso.Tipo -Nome $acesso.Nome
                    if ($MapaMarcados.ContainsKey($chaveAcesso)) {
                        $accessNode.Checked = $true
                    }

                    [void]$groupNode.Nodes.Add($accessNode)
                }

                [void]$projectNode.Nodes.Add($groupNode)
            }

            if ($listasFiltradas.Count -gt 0) {
                $userListNode = New-Object System.Windows.Forms.TreeNode(("User Lists ({0})" -f $listasFiltradas.Count))
                $userListNode.Tag = [PSCustomObject]@{
                    NodeType    = 'Category'
                    Category    = 'UserList'
                    ProjetoNome = $projetoData.ProjetoNome
                }

                foreach ($acesso in $listasFiltradas) {
                    $accessNode = New-Object System.Windows.Forms.TreeNode($acesso.Nome)
                    $accessNode.Tag = [PSCustomObject]@{
                        NodeType    = 'Access'
                        Tipo        = $acesso.Tipo
                        Nome        = $acesso.Nome
                        ObjetoPw    = $acesso.ObjetoPw
                        ProjetoNome = $projetoData.ProjetoNome
                    }

                    $chaveAcesso = Obter-ChaveAcessoProjeto -ProjetoNome $projetoData.ProjetoNome -Tipo $acesso.Tipo -Nome $acesso.Nome
                    if ($MapaMarcados.ContainsKey($chaveAcesso)) {
                        $accessNode.Checked = $true
                    }

                    [void]$userListNode.Nodes.Add($accessNode)
                }

                [void]$projectNode.Nodes.Add($userListNode)
            }

            if ($outrosFiltrados.Count -gt 0) {
                $outrosNode = New-Object System.Windows.Forms.TreeNode(("Outros acessos - somente visual ({0})" -f $outrosFiltrados.Count))
                $outrosNode.ForeColor = [System.Drawing.Color]::Gray
                $outrosNode.Tag = [PSCustomObject]@{
                    NodeType      = 'CategoryInfo'
                    Category      = 'Other'
                    ProjetoNome   = $projetoData.ProjetoNome
                    SomenteVisual = $true
                }

                foreach ($acesso in $outrosFiltrados) {
                    $textoOutro = "{0} | {1}" -f $acesso.Tipo, $acesso.Nome
                    $otherNode = New-Object System.Windows.Forms.TreeNode($textoOutro)
                    $otherNode.ForeColor = [System.Drawing.Color]::Gray
                    $otherNode.Tag = [PSCustomObject]@{
                        NodeType    = 'OtherAccess'
                        Tipo        = $acesso.Tipo
                        Nome        = $acesso.Nome
                        ObjetoPw    = $acesso.ObjetoPw
                        ProjetoNome = $projetoData.ProjetoNome
                        SomenteVisual = $true
                    }

                    [void]$outrosNode.Nodes.Add($otherNode)
                }

                [void]$projectNode.Nodes.Add($outrosNode)
            }

            Atualizar-CheckVisualProjeto -ProjetoNode $projectNode
            [void]$script:TreeProjetos.Nodes.Add($projectNode)
        }
    }
    finally {
        $script:AtualizandoCheckTree = $false
    }

    $script:TreeProjetos.ExpandAll()

    if ($script:LabelResultadoFiltro) {
        $qtdProjetosVisiveis = $script:TreeProjetos.Nodes.Count
        $qtdProjetosOriginais = @($ProjetosData).Count

        if ([string]::IsNullOrWhiteSpace($Filtro) -and [string]::IsNullOrWhiteSpace($ExcluirFiltro)) {
            $script:LabelResultadoFiltro.Text = "Projetos exibidos: $qtdProjetosVisiveis"
        }
        elseif ([string]::IsNullOrWhiteSpace($Filtro)) {
            $script:LabelResultadoFiltro.Text = "Exclusao ativa: $qtdProjetosVisiveis de $qtdProjetosOriginais projeto(s) exibidos"
        }
        elseif ($TermoExato) {
            if ([string]::IsNullOrWhiteSpace($ExcluirFiltro)) {
                $script:LabelResultadoFiltro.Text = "Filtro avancado ativo: $qtdProjetosVisiveis de $qtdProjetosOriginais projeto(s) exibidos | pesquisa: $Filtro"
            }
            else {
                $script:LabelResultadoFiltro.Text = "Filtro avancado ativo: $qtdProjetosVisiveis de $qtdProjetosOriginais projeto(s) exibidos | pesquisa: $Filtro | excluindo: $ExcluirFiltro"
            }
        }
        else {
            $script:LabelResultadoFiltro.Text = "Filtro ativo: $qtdProjetosVisiveis de $qtdProjetosOriginais projeto(s) exibidos"
        }
    }
}

function Aplicar-FiltroProjetos {
    param(
        [switch]$TermoExato,
        [string]$ExcluirFiltro = ""
    )

    $filtro = ""
    if ($script:TextFiltroAcessos) {
        $filtro = $script:TextFiltroAcessos.Text
    }

    Renderizar-ArvoreProjetos -ProjetosData $script:ProjetosCarregados -Filtro $filtro -ExcluirFiltro $ExcluirFiltro -TermoExato:$TermoExato.IsPresent
    Atualizar-EstadoInterface
}

function Limpar-CamposFiltroProjetos {
    if ($script:TextFiltroAcessos) {
        $script:TextFiltroAcessos.Text = ""
    }
    $script:FiltroExclusaoAvancado = ""
}

function Mostrar-ModalFiltroAvancado {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Filtro avancado"
    $dialog.Size = New-Object System.Drawing.Size(520,230)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ShowInTaskbar = $false

    $labelPesquisa = New-Object System.Windows.Forms.Label
    $labelPesquisa.Text = "Termo de pesquisa:"
    $labelPesquisa.AutoSize = $true
    $labelPesquisa.Location = New-Object System.Drawing.Point(18,22)
    $dialog.Controls.Add($labelPesquisa)

    $txtPesquisa = New-Object System.Windows.Forms.TextBox
    $txtPesquisa.Location = New-Object System.Drawing.Point(18,45)
    $txtPesquisa.Size = New-Object System.Drawing.Size(466,25)
    $txtPesquisa.Font = $fontePadrao
    if ($script:TextFiltroAcessos) {
        $txtPesquisa.Text = $script:TextFiltroAcessos.Text
    }
    $dialog.Controls.Add($txtPesquisa)

    $labelExclusao = New-Object System.Windows.Forms.Label
    $labelExclusao.Text = "Termo de exclusao:"
    $labelExclusao.AutoSize = $true
    $labelExclusao.Location = New-Object System.Drawing.Point(18,82)
    $dialog.Controls.Add($labelExclusao)

    $txtExclusao = New-Object System.Windows.Forms.TextBox
    $txtExclusao.Location = New-Object System.Drawing.Point(18,105)
    $txtExclusao.Size = New-Object System.Drawing.Size(466,25)
    $txtExclusao.Font = $fontePadrao
    $txtExclusao.Text = $script:FiltroExclusaoAvancado
    $dialog.Controls.Add($txtExclusao)

    $btnPesquisar = New-Object System.Windows.Forms.Button
    $btnPesquisar.Text = "Pesquisar"
    $btnPesquisar.Location = New-Object System.Drawing.Point(282,148)
    $btnPesquisar.Size = New-Object System.Drawing.Size(95,30)
    $btnPesquisar.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnPesquisar)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(389,148)
    $btnCancelar.Size = New-Object System.Drawing.Size(95,30)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnCancelar)

    $dialog.AcceptButton = $btnPesquisar
    $dialog.CancelButton = $btnCancelar

    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $termoPesquisa = $txtPesquisa.Text
        $termoExclusao = $txtExclusao.Text

        if ($script:TextFiltroAcessos) {
            $script:TextFiltroAcessos.Text = $termoPesquisa
        }
        $script:FiltroExclusaoAvancado = $termoExclusao

        if ([string]::IsNullOrWhiteSpace($termoPesquisa) -and [string]::IsNullOrWhiteSpace($termoExclusao)) {
            Aplicar-FiltroProjetos
        }
        else {
            Aplicar-FiltroProjetos -TermoExato -ExcluirFiltro $termoExclusao
        }
    }
}

function Selecionar-AcessosExibidosNaArvore {
    if (-not $script:TreeProjetos -or $script:TreeProjetos.Nodes.Count -eq 0) {
        return 0
    }

    $totalMarcados = 0
    $script:AtualizandoCheckTree = $true

    try {
        foreach ($projetoNode in $script:TreeProjetos.Nodes) {
            foreach ($categoriaNode in $projetoNode.Nodes) {
                foreach ($acessoNode in $categoriaNode.Nodes) {
                    $tagAcesso = $acessoNode.Tag

                    if ($null -eq $tagAcesso -or $tagAcesso.NodeType -ne 'Access') {
                        $acessoNode.Checked = $false
                        continue
                    }

                    if (-not $acessoNode.Checked) {
                        $acessoNode.Checked = $true
                    }

                    $totalMarcados++
                }
            }

            Atualizar-CheckVisualProjeto -ProjetoNode $projetoNode
        }
    }
    finally {
        $script:AtualizandoCheckTree = $false
    }

    return $totalMarcados
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

function Obter-ChavesComparacaoProjeto {
    param([object]$ProjetoData)

    $valores = @(
        $ProjetoData.ProjetoNome,
        $ProjetoData.ProjetoNomeTecnico
    )

    if ($ProjetoData.ProjetoObjeto) {
        $valores += @(Obter-DescricaoItem -Item $ProjetoData.ProjetoObjeto)
        $valores += @(Obter-NomeAmigavelItem -Item $ProjetoData.ProjetoObjeto)
    }

    return @(
        $valores |
            ForEach-Object { Normalizar-TextoComparacaoProjeto $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
}

function Localizar-ProjetoCarregadoPorNome {
    param([string]$NomeProjetoPlanilha)

    $alvo = Normalizar-TextoComparacaoProjeto $NomeProjetoPlanilha
    if ([string]::IsNullOrWhiteSpace($alvo)) {
        return [PSCustomObject]@{
            Encontrado = $false
            Projeto    = $null
            Motivo     = "Nome vazio"
            Sugestoes  = @()
        }
    }

    foreach ($projetoData in @($script:ProjetosCarregados)) {
        $chaves = @(Obter-ChavesComparacaoProjeto -ProjetoData $projetoData)
        if ($chaves -contains $alvo) {
            return [PSCustomObject]@{
                Encontrado = $true
                Projeto    = $projetoData
                Motivo     = "Exato"
                Sugestoes  = @()
            }
        }
    }

    foreach ($projetoData in @($script:ProjetosCarregados)) {
        $chaves = @(Obter-ChavesComparacaoProjeto -ProjetoData $projetoData)
        foreach ($chave in $chaves) {
            if ($chave.Contains($alvo) -or $alvo.Contains($chave)) {
                return [PSCustomObject]@{
                    Encontrado = $true
                    Projeto    = $projetoData
                    Motivo     = "Parcial"
                    Sugestoes  = @()
                }
            }
        }
    }

    $termosAlvo = @($alvo -split '\s+' | Where-Object { $_.Length -ge 3 })
    $sugestoes = @(
        $script:ProjetosCarregados |
            ForEach-Object {
                $projetoData = $_
                $melhorPontuacao = 0
                foreach ($chave in @(Obter-ChavesComparacaoProjeto -ProjetoData $projetoData)) {
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
                    [PSCustomObject]@{
                        Projeto   = $projetoData
                        Pontuacao = $melhorPontuacao
                    }
                }
            } |
            Sort-Object Pontuacao -Descending |
            Select-Object -First 3
    )

    return [PSCustomObject]@{
        Encontrado = $false
        Projeto    = $null
        Motivo     = "Nao encontrado"
        Sugestoes  = @($sugestoes | ForEach-Object { $_.Projeto.ProjetoNome })
    }
}

function Obter-SufixoPerfilNomeAcesso {
    param([string]$NomeAcesso)

    if ([string]::IsNullOrWhiteSpace($NomeAcesso)) {
        return ""
    }

    $nome = $NomeAcesso.Trim()
    $indice = $nome.LastIndexOf('-')
    if ($indice -ge 0 -and $indice -lt ($nome.Length - 1)) {
        return Normalizar-TextoComparacaoProjeto ($nome.Substring($indice + 1))
    }

    return Normalizar-TextoComparacaoProjeto $nome
}

function Mostrar-ModalSelecaoPerfisProjeto {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Selecionar perfis"
    $dialog.Size = New-Object System.Drawing.Size(430,460)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ShowInTaskbar = $false

    $labelInfo = New-Object System.Windows.Forms.Label
    $labelInfo.Text = "Selecione os perfis que devem ser marcados nos projetos encontrados:"
    $labelInfo.Location = New-Object System.Drawing.Point(15,18)
    $labelInfo.Size = New-Object System.Drawing.Size(380,40)
    $dialog.Controls.Add($labelInfo)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(18,65)
    $checkedList.Size = New-Object System.Drawing.Size(375,285)
    $checkedList.CheckOnClick = $true
    $checkedList.Font = $fontePadrao
    foreach ($perfil in @($script:PerfisPadraoProjeto)) {
        [void]$checkedList.Items.Add($perfil.Rotulo)
    }
    $dialog.Controls.Add($checkedList)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Aplicar"
    $btnOk.Location = New-Object System.Drawing.Point(190,370)
    $btnOk.Size = New-Object System.Drawing.Size(95,30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnOk)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(298,370)
    $btnCancelar.Size = New-Object System.Drawing.Size(95,30)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnCancelar)

    $dialog.AcceptButton = $btnOk
    $dialog.CancelButton = $btnCancelar

    if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $selecionados = @()
    foreach ($item in $checkedList.CheckedItems) {
        $perfil = @($script:PerfisPadraoProjeto | Where-Object { $_.Rotulo -eq [string]$item } | Select-Object -First 1)
        if ($perfil.Count -gt 0) {
            $selecionados += $perfil[0]
        }
    }

    return @($selecionados)
}

function Limpar-MarcacoesArvoreProjetos {
    if (-not $script:TreeProjetos) {
        return
    }

    foreach ($projetoNode in $script:TreeProjetos.Nodes) {
        $projetoNode.Checked = $false
        foreach ($categoriaNode in $projetoNode.Nodes) {
            $categoriaNode.Checked = $false
            foreach ($acessoNode in $categoriaNode.Nodes) {
                $acessoNode.Checked = $false
            }
        }
    }
}

function Aplicar-PerfisEmProjetosImportados {
    param(
        [string[]]$ProjetosPlanilha,
        [array]$Perfis
    )

    if (-not $script:ConcessaoSelecionada -or -not $script:ProjetosCarregados -or $script:ProjetosCarregados.Count -eq 0) {
        throw "Selecione e carregue uma concessao antes de importar a lista de projetos."
    }

    if (-not $ProjetosPlanilha -or $ProjetosPlanilha.Count -eq 0) {
        throw "Nenhum projeto foi informado na planilha."
    }

    if (-not $Perfis -or $Perfis.Count -eq 0) {
        throw "Nenhum perfil foi selecionado."
    }

    Limpar-CamposFiltroProjetos
    Renderizar-ArvoreProjetos -ProjetosData $script:ProjetosCarregados -Filtro ""

    $script:AtualizandoCheckTree = $true
    $encontrados = New-Object System.Collections.Generic.List[object]
    $naoEncontrados = New-Object System.Collections.Generic.List[object]
    $perfisNaoEncontrados = New-Object System.Collections.Generic.List[object]
    $totalAcessosMarcados = 0

    try {
        Limpar-MarcacoesArvoreProjetos

        $mapaProjetosImportados = @{}
        foreach ($nomeProjeto in @($ProjetosPlanilha)) {
            $resultado = Localizar-ProjetoCarregadoPorNome -NomeProjetoPlanilha $nomeProjeto
            if (-not $resultado.Encontrado) {
                $naoEncontrados.Add([PSCustomObject]@{
                    ProjetoPlanilha = $nomeProjeto
                    Sugestoes       = @($resultado.Sugestoes)
                })
                continue
            }

            $chaveProjeto = Obter-ChaveProjetoSelecao -ConcessaoNome $resultado.Projeto.ConcessaoNome -ProjetoNome $resultado.Projeto.ProjetoNome
            if (-not $mapaProjetosImportados.ContainsKey($chaveProjeto)) {
                $mapaProjetosImportados[$chaveProjeto] = $resultado.Projeto
                $encontrados.Add([PSCustomObject]@{
                    ProjetoPlanilha = $nomeProjeto
                    ProjetoNome     = $resultado.Projeto.ProjetoNome
                    Motivo          = $resultado.Motivo
                })
            }
        }

        foreach ($projetoNode in $script:TreeProjetos.Nodes) {
            $tagProjeto = $projetoNode.Tag
            if ($null -eq $tagProjeto -or $tagProjeto.NodeType -ne 'Project') {
                continue
            }

            $chaveProjetoNode = Obter-ChaveProjetoSelecao -ConcessaoNome $tagProjeto.ConcessaoNome -ProjetoNome $tagProjeto.ProjetoNome
            if (-not $mapaProjetosImportados.ContainsKey($chaveProjetoNode)) {
                continue
            }

            foreach ($perfil in @($Perfis)) {
                $sufixoNormalizado = Normalizar-TextoComparacaoProjeto $perfil.Sufixo
                $matchesPerfil = New-Object System.Collections.Generic.List[System.Windows.Forms.TreeNode]

                foreach ($categoriaNode in $projetoNode.Nodes) {
                    foreach ($acessoNode in $categoriaNode.Nodes) {
                        $tagAcesso = $acessoNode.Tag
                        if ($null -eq $tagAcesso -or $tagAcesso.NodeType -ne 'Access') {
                            continue
                        }

                        if ((Normalizar-Texto $tagAcesso.Tipo) -ne "group") {
                            continue
                        }

                        $perfilAcessoNormalizado = Obter-SufixoPerfilNomeAcesso -NomeAcesso $tagAcesso.Nome
                        if ($perfilAcessoNormalizado -eq $sufixoNormalizado) {
                            $matchesPerfil.Add($acessoNode)
                        }
                    }
                }

                if ($matchesPerfil.Count -eq 0) {
                    $perfisNaoEncontrados.Add([PSCustomObject]@{
                        Projeto = $tagProjeto.ProjetoNome
                        Perfil  = $perfil.Rotulo
                    })
                    continue
                }

                foreach ($nodePerfil in $matchesPerfil) {
                    if (-not $nodePerfil.Checked) {
                        $nodePerfil.Checked = $true
                        $totalAcessosMarcados++
                    }
                }
            }

            Atualizar-CheckVisualProjeto -ProjetoNode $projetoNode
        }
    }
    finally {
        $script:AtualizandoCheckTree = $false
    }

    Atualizar-EstadoInterface

    return [PSCustomObject]@{
        ProjetosPlanilha       = @($ProjetosPlanilha).Count
        ProjetosEncontrados    = @($encontrados.ToArray())
        ProjetosNaoEncontrados = @($naoEncontrados.ToArray())
        PerfisSelecionados     = @($Perfis)
        PerfisNaoEncontrados   = @($perfisNaoEncontrados.ToArray())
        AcessosMarcados        = $totalAcessosMarcados
    }
}

function Mostrar-ResumoImportacaoProjetosPerfis {
    param([object]$Resumo)

    $mensagem = "Importacao de projetos concluida.`r`n`r`n"
    $mensagem += "Projetos na planilha: $($Resumo.ProjetosPlanilha)`r`n"
    $mensagem += "Projetos encontrados: $(@($Resumo.ProjetosEncontrados).Count)`r`n"
    $mensagem += "Projetos nao encontrados: $(@($Resumo.ProjetosNaoEncontrados).Count)`r`n"
    $mensagem += "Perfis selecionados: $(@($Resumo.PerfisSelecionados).Count)`r`n"
    $mensagem += "Acessos marcados na arvore: $($Resumo.AcessosMarcados)`r`n"
    $mensagem += "Pendencias de perfil: $(@($Resumo.PerfisNaoEncontrados).Count)`r`n"

    if (@($Resumo.ProjetosNaoEncontrados).Count -gt 0) {
        $mensagem += "`r`nProjetos nao encontrados:"
        foreach ($item in @($Resumo.ProjetosNaoEncontrados | Select-Object -First 8)) {
            $mensagem += "`r`n- $($item.ProjetoPlanilha)"
            if (@($item.Sugestoes).Count -gt 0) {
                $mensagem += " | sugestao: $(@($item.Sugestoes)[0])"
            }
        }
        if (@($Resumo.ProjetosNaoEncontrados).Count -gt 8) {
            $mensagem += "`r`n- ..."
        }
    }

    if (@($Resumo.PerfisNaoEncontrados).Count -gt 0) {
        $mensagem += "`r`n`r`nPerfis nao encontrados:"
        foreach ($item in @($Resumo.PerfisNaoEncontrados | Select-Object -First 8)) {
            $mensagem += "`r`n- $($item.Projeto) | $($item.Perfil)"
        }
        if (@($Resumo.PerfisNaoEncontrados).Count -gt 8) {
            $mensagem += "`r`n- ..."
        }
    }

    Show-UiInfo $mensagem
}

function Obter-ChaveUsuario {
    param([object]$Usuario)

    $email = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email","EMail","Mail")
    if (-not [string]::IsNullOrWhiteSpace($email)) {
        return (Normalizar-Texto $email)
    }

    return Normalizar-Texto (Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("UserID","UserId","ID","Id","Name","UserName","LoginName"))
}

function Obter-RotuloUsuario {
    param([object]$Usuario)

    $nome = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Name","UserName","FullName")
    $email = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email","EMail","Mail")
    $login = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("LoginName","Login","UserID")

    return "$nome | $email | Login: $login"
}

function Atualizar-ListaUsuariosSelecionados {
    if (-not $script:ListaUsuariosSelecionados) {
        return
    }

    $script:ListaUsuariosSelecionados.Items.Clear()
    foreach ($usuario in @($script:UsuariosSelecionados)) {
        [void]$script:ListaUsuariosSelecionados.Items.Add((Obter-RotuloUsuario -Usuario $usuario))
    }

    if ($script:UsuariosSelecionados.Count -eq 0) {
        $labelUsuarioEncontrado.Text = "Usuário: -"
    }
    elseif ($script:ModoLoteUsuarios) {
        $labelUsuarioEncontrado.Text = "Usuários confirmados: $($script:UsuariosSelecionados.Count)"
    }
    else {
        $labelUsuarioEncontrado.Text = "Usuario: $(Obter-RotuloUsuario -Usuario $script:UsuariosSelecionados[0])"
    }
}

function Adicionar-UsuarioSelecionado {
    param(
        [object]$Usuario,
        [switch]$Substituir
    )

    if ($Substituir) {
        $script:UsuariosSelecionados = @()
    }

    $chaveNova = Obter-ChaveUsuario -Usuario $Usuario
    $jaExiste = $false
    foreach ($usuarioExistente in @($script:UsuariosSelecionados)) {
        if ((Obter-ChaveUsuario -Usuario $usuarioExistente) -eq $chaveNova) {
            $jaExiste = $true
            break
        }
    }

    if (-not $jaExiste) {
        $script:UsuariosSelecionados = @($script:UsuariosSelecionados + $Usuario)
    }

    $script:UsuarioEncontrado = @($script:UsuariosSelecionados | Select-Object -First 1)[0]
    Atualizar-ListaUsuariosSelecionados
    return (-not $jaExiste)
}

function Obter-UsuariosParaExecucao {
    if ($script:ModoLoteUsuarios) {
        return @($script:UsuariosSelecionados | Where-Object { $_ -ne $null })
    }

    if ($script:UsuarioEncontrado) {
        return @($script:UsuarioEncontrado)
    }

    return @()
}

function Obter-EmailsInformados {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return @()
    }

    # Remove espacos invisiveis e "mailto:" comuns quando o e-mail vem copiado do Outlook.
    $limpo = $Texto -replace '[ ​﻿]', ' ' -replace '(?i)mailto:', ''

    # Se ha "@", extrai somente os e-mails (aceita "Fulano <fulano@x.com>").
    if ($limpo.Contains('@')) {
        $extraidos = @(
            [regex]::Matches($limpo, '[^\s<>"''()\[\];:,|]+@[^\s<>"''()\[\];:,|]+') |
                ForEach-Object { $_.Value.Trim().TrimEnd('.') } |
                Where-Object { $_ -ne "" } |
                Select-Object -Unique
        )
        if ($extraidos.Count -gt 0) {
            return $extraidos
        }
    }

    # Sem "@": devolve o texto para que seja reportado como e-mail invalido.
    return @(
        $limpo -split '[,;|\r\n]+' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -ne "" } |
            Select-Object -Unique
    )
}

function Obter-ValorCampoLinha {
    param(
        [object]$Linha,
        [string[]]$Nomes
    )

    foreach ($nome in $Nomes) {
        $propriedade = @($Linha.PSObject.Properties | Where-Object { $_.Name -ieq $nome } | Select-Object -First 1)
        if ($propriedade.Count -gt 0 -and $null -ne $propriedade[0].Value) {
            $valor = ([string]$propriedade[0].Value).Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
}

function Selecionar-PlanilhaUsuarios {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
    $dialog.Title = "Selecione a planilha com os usuarios"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    throw "Nenhuma planilha foi selecionada."
}

function Importar-EmailsUsuarios {
    param([string]$CaminhoArquivo)

    if (-not (Test-Path $CaminhoArquivo)) {
        throw "Arquivo nao encontrado: $CaminhoArquivo"
    }

    $nomesColunaEmail = @("E-mail", "Email", "E-mails", "Emails", "Mail", "Usuario", "Usuário", "Usuarios", "Usuários", "User", "Login")

    $ext = [System.IO.Path]::GetExtension($CaminhoArquivo).ToLowerInvariant()
    $planilhas = @()   # lista de @{ Nome; Dados }
    if ($ext -in @(".xlsx", ".xlsm")) {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw "O modulo 'ImportExcel' nao esta instalado neste computador. Instale com: Install-Module ImportExcel -Scope CurrentUser`r`nOu salve a planilha como CSV e importe novamente."
        }
        Import-Module ImportExcel -ErrorAction Stop

        try {
            $nomesAbas = @(Get-ExcelSheetInfo -Path $CaminhoArquivo | ForEach-Object { $_.Name })
        }
        catch {
            throw "Nao foi possivel abrir a planilha. Verifique se ela esta aberta no Excel (feche e tente de novo) ou corrompida.`r`nDetalhe: $($_.Exception.Message)"
        }

        foreach ($aba in $nomesAbas) {
            try {
                $planilhas += [PSCustomObject]@{ Nome = $aba; Dados = @(Import-Excel -Path $CaminhoArquivo -WorksheetName $aba) }
            }
            catch {
                Write-Log "Aba '$aba' ignorada na leitura de usuarios: $($_.Exception.Message)" "WARN"
            }
        }
    }
    elseif ($ext -eq ".csv") {
        $primeiraLinha = Get-Content -Path $CaminhoArquivo -TotalCount 1
        $qtdPontoVirgula = ([regex]::Matches($primeiraLinha, ';')).Count
        $qtdVirgula = ([regex]::Matches($primeiraLinha, ',')).Count
        $delimitador = ','
        if ($qtdPontoVirgula -gt $qtdVirgula) {
            $delimitador = ';'
        }

        $planilhas += [PSCustomObject]@{ Nome = "CSV"; Dados = @(Import-Csv -Path $CaminhoArquivo -Delimiter $delimitador) }
    }
    else {
        throw "Extensao nao suportada: $ext. Use .xlsx, .xlsm ou .csv."
    }

    # Escolhe a primeira aba que tenha coluna de e-mail reconhecida. Se nenhuma
    # tiver, tenta achar uma coluna cujo conteudo seja majoritariamente e-mail.
    $dados = $null
    $colunaEmail = $null
    $abaUsada = $null
    $colunasVistas = New-Object System.Collections.Generic.List[string]

    foreach ($planilha in $planilhas) {
        if ($planilha.Dados.Count -eq 0) { continue }
        $colunas = @($planilha.Dados[0].PSObject.Properties | ForEach-Object { $_.Name })
        $colunasVistas.Add("[$($planilha.Nome)] " + ($colunas -join ", "))

        $coluna = $colunas | Where-Object { $c = $_.Trim(); $nomesColunaEmail | Where-Object { $_ -ieq $c } } | Select-Object -First 1
        if (-not $coluna) {
            foreach ($c in $colunas) {
                $valores = @($planilha.Dados | Select-Object -First 50 | ForEach-Object { [string]$_.$c } | Where-Object { $_ -and $_.Trim() })
                $qtdEmail = @($valores | Where-Object { $_ -match '[^@\s<>]+@[^@\s<>]+\.[^@\s<>]+' }).Count
                if ($valores.Count -gt 0 -and $qtdEmail -ge [Math]::Ceiling($valores.Count / 2)) {
                    $coluna = $c
                    Write-Log "Coluna de e-mail detectada pelo conteudo: '$c' (aba '$($planilha.Nome)')." "WARN"
                    break
                }
            }
        }

        if ($coluna) {
            $dados = $planilha.Dados
            $colunaEmail = $coluna
            $abaUsada = $planilha.Nome
            break
        }
    }

    if (-not $dados) {
        if ($colunasVistas.Count -eq 0) {
            throw "A planilha selecionada esta vazia (nenhuma linha de dados abaixo do cabecalho)."
        }
        throw "Nao encontrei a coluna de e-mail. A 1a linha deve ter um cabecalho como: E-mail, Email, Mail ou Usuario.`r`n`r`nColunas encontradas:`r`n$($colunasVistas -join "`r`n")"
    }

    $emails = New-Object System.Collections.Generic.List[string]
    $emailsUnicos = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $linhaPorEmail = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $duplicados = New-Object System.Collections.Generic.List[string]
    $linhasSemEmail = New-Object System.Collections.Generic.List[int]
    $totalEmailsEncontrados = 0
    $numeroLinha = 2

    foreach ($linha in $dados) {
        $valor = Obter-ValorCampoLinha -Linha $linha -Nomes @($colunaEmail)
        if ([string]::IsNullOrWhiteSpace($valor)) {
            $linhasSemEmail.Add($numeroLinha)
            $numeroLinha++
            continue
        }

        foreach ($email in @(Obter-EmailsInformados -Texto $valor)) {
            $totalEmailsEncontrados++
            if ($emailsUnicos.Add($email)) {
                $emails.Add($email)
                $linhaPorEmail[$email] = $numeroLinha
            }
            else {
                $duplicados.Add("$email (linha $numeroLinha)")
            }
        }

        $numeroLinha++
    }

    if ($emails.Count -eq 0) {
        throw "A coluna '$colunaEmail' (aba '$abaUsada') nao possui nenhum e-mail preenchido."
    }

    Write-Log "Planilha de usuarios lida: aba '$abaUsada', coluna '$colunaEmail', $($dados.Count) linha(s), $($emails.Count) e-mail(s) unico(s)."

    $script:ResumoImportacaoUsuarios = [PSCustomObject]@{
        Aba                     = $abaUsada
        Coluna                  = $colunaEmail
        LinhasPlanilha          = $dados.Count
        TotalEmailsEncontrados  = $totalEmailsEncontrados
        EmailsUnicos            = $emails.Count
        DuplicadosIgnorados     = $duplicados.Count
        EmailsDuplicados        = @($duplicados | Select-Object -Unique)
        LinhasSemEmail          = @($linhasSemEmail.ToArray())
        LinhaPorEmail           = $linhaPorEmail
    }

    return $emails.ToArray()
}

function Selecionar-PlanilhaProjetos {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
    $dialog.Title = "Selecione a planilha com a lista de projetos"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    throw "Nenhuma planilha de projetos foi selecionada."
}

function Importar-NomesProjetos {
    param([string]$CaminhoArquivo)

    if (-not (Test-Path $CaminhoArquivo)) {
        throw "Arquivo nao encontrado: $CaminhoArquivo"
    }

    $ext = [System.IO.Path]::GetExtension($CaminhoArquivo).ToLowerInvariant()
    if ($ext -in @(".xlsx", ".xlsm")) {
        Import-Module ImportExcel -ErrorAction Stop
        $dados = @(Import-Excel -Path $CaminhoArquivo)
    }
    elseif ($ext -eq ".csv") {
        $primeiraLinha = Get-Content -Path $CaminhoArquivo -TotalCount 1
        $qtdPontoVirgula = ([regex]::Matches($primeiraLinha, ';')).Count
        $qtdVirgula = ([regex]::Matches($primeiraLinha, ',')).Count
        $delimitador = ','
        if ($qtdPontoVirgula -gt $qtdVirgula) {
            $delimitador = ';'
        }

        $dados = @(Import-Csv -Path $CaminhoArquivo -Delimiter $delimitador)
    }
    else {
        throw "Extensao nao suportada: $ext. Use .xlsx, .xlsm ou .csv."
    }

    if (-not $dados -or $dados.Count -eq 0) {
        throw "A planilha selecionada nao possui projetos."
    }

    $projetos = New-Object System.Collections.Generic.List[string]
    $projetosUnicos = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
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
        if ($projetosUnicos.Add($nomeProjeto)) {
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

    $script:ResumoImportacaoProjetos = [PSCustomObject]@{
        LinhasPlanilha       = $dados.Count
        ProjetosUnicos       = $projetos.Count
        DuplicadosIgnorados  = $duplicados.Count
        ProjetosDuplicados   = @($duplicados | Select-Object -Unique)
        LinhasSemProjeto     = @($linhasSemProjeto.ToArray())
    }

    return $projetos.ToArray()
}

$script:RotulosStatusUsuario = @{
    Encontrado    = "OK"
    Escolhido     = "OK (escolhido)"
    Parcial       = "Conferir"
    Ambiguo       = "Mais de um"
    NaoEncontrado = "Não encontrado"
    Invalido      = "E-mail inválido"
    Erro          = "Erro na consulta"
}

function Obter-RotuloUsuarioDetalhado {
    param([object]$Usuario)

    $nome = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Name","UserName","FullName")
    $descricao = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Description","FullName")
    $email = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email","EMail","Mail")
    $id = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("UserID","UserId","ID","Id")
    $desativado = Obter-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Disabled","IsDisabled")
    $texto = "$nome | $email | ID: $id"
    if ($descricao -and $descricao -ne $nome) { $texto += " | $descricao" }
    if ($desativado -eq "True") { $texto += " | DESATIVADO" }
    return $texto
}

function Resolver-ListaEmailsUsuarios {
    param(
        [string[]]$Emails,
        $LinhaPorEmail = $null
    )

    $itens = New-Object System.Collections.Generic.List[object]
    $total = @($Emails).Count
    $i = 0
    foreach ($emailInformado in $Emails) {
        $i++
        $linha = ""
        if ($LinhaPorEmail -and $LinhaPorEmail.ContainsKey($emailInformado)) {
            $linha = [string]$LinhaPorEmail[$emailInformado]
        }

        try {
            $email = Validar-Email -Email $emailInformado
        }
        catch {
            $itens.Add([PSCustomObject]@{
                Linha = $linha; Email = $emailInformado; Status = "Invalido"; Usuario = $null; Candidatos = @()
                Detalhe = "Formato de e-mail invalido."
            })
            Write-UiLog "[$i/$total] E-mail invalido$(if ($linha) { " (linha $linha)" }): $emailInformado"
            Write-Log "E-mail invalido$(if ($linha) { " (linha $linha)" }): '$emailInformado'" "WARN"
            continue
        }

        Write-UiLog "[$i/$total] Buscando usuario: $email"
        $r = Resolver-UsuarioPW -Email $email

        if ($r.Usuario) {
            try {
                $null = Validar-IdentificadoresUsuario -Usuario $r.Usuario
            }
            catch {
                $r.Status = "Erro"
                $r.Detalhe = $_.Exception.Message
                $r.Usuario = $null
            }
        }

        $itens.Add([PSCustomObject]@{
            Linha = $linha; Email = $email; Status = $r.Status; Usuario = $r.Usuario; Candidatos = @($r.Candidatos)
            Detalhe = $r.Detalhe
        })

        if ($r.Status -ne "Encontrado") {
            Write-UiLog "    -> $($script:RotulosStatusUsuario[$r.Status]): $($r.Detalhe)"
        }
    }

    return $itens.ToArray()
}

function Salvar-PendenciasUsuarios {
    param(
        [array]$Itens,
        [string]$Caminho
    )

    $pendentes = @($Itens | Where-Object { $_.Status -notin @("Encontrado", "Escolhido") })
    if ($pendentes.Count -eq 0) {
        return $null
    }

    $pendentes | ForEach-Object {
        [PSCustomObject]@{
            Linha    = $_.Linha
            Email    = $_.Email
            Situacao = $script:RotulosStatusUsuario[$_.Status]
            Detalhe  = $_.Detalhe
            Candidatos = (@($_.Candidatos | ForEach-Object { Obter-RotuloUsuarioDetalhado -Usuario $_ }) -join " || ")
        }
    } | Export-Csv -Path $Caminho -NoTypeInformation -Delimiter ';' -Encoding UTF8

    return $Caminho
}

function Atualizar-LinhaRevisaoUsuario {
    param([int]$Indice)

    $item = $script:RevisaoUsuariosItens[$Indice]
    $row = $script:RevisaoUsuariosGrid.Rows[$Indice]
    $u = $item.Usuario

    $row.Cells["Situacao"].Value = $script:RotulosStatusUsuario[$item.Status]
    $row.Cells["Nome"].Value = if ($u) { Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("Name","UserName","FullName") } else { "" }
    $row.Cells["EmailPW"].Value = if ($u) { Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("Email","EMail","Mail") } else { "" }
    $row.Cells["Detalhe"].Value = $item.Detalhe

    $podeIncluir = [bool]$u
    $row.Cells["Incluir"].ReadOnly = -not $podeIncluir
    if (-not $podeIncluir) { $row.Cells["Incluir"].Value = $false }

    $cor = switch ($item.Status) {
        "Encontrado" { [System.Drawing.Color]::White }
        "Escolhido"  { [System.Drawing.Color]::Honeydew }
        "Parcial"    { [System.Drawing.Color]::LightYellow }
        "Ambiguo"    { [System.Drawing.Color]::Moccasin }
        default      { [System.Drawing.Color]::MistyRose }
    }
    $row.DefaultCellStyle.BackColor = $cor
}

function Atualizar-ResumoRevisaoUsuarios {
    $itens = $script:RevisaoUsuariosItens
    $ok = @($itens | Where-Object { $_.Status -in @("Encontrado","Escolhido") }).Count
    $conferir = @($itens | Where-Object { $_.Status -eq "Parcial" }).Count
    $ambiguos = @($itens | Where-Object { $_.Status -eq "Ambiguo" }).Count
    $falhas = @($itens | Where-Object { $_.Status -in @("NaoEncontrado","Invalido","Erro") }).Count

    $marcados = 0
    foreach ($row in $script:RevisaoUsuariosGrid.Rows) {
        if ($row.Cells["Incluir"].Value -eq $true) { $marcados++ }
    }

    $script:RevisaoUsuariosLabel.Text = "Total: $($itens.Count)   |   OK: $ok   |   Conferir: $conferir   |   Mais de um: $ambiguos   |   Não encontrados/inválidos: $falhas`r`n" +
        "Linhas vermelhas não serão incluídas. Linhas laranja: dê duplo clique para escolher o usuário correto. Linhas amarelas: confira e marque se estiver certo."
    $script:RevisaoUsuariosBtnOk.Text = "Incluir marcados ($marcados)"
    $script:RevisaoUsuariosBtnOk.Enabled = $marcados -gt 0
}

function Escolher-CandidatoRevisaoUsuario {
    param([int]$Indice)

    if ($Indice -lt 0) { return }
    $item = $script:RevisaoUsuariosItens[$Indice]
    if (@($item.Candidatos).Count -eq 0) { return }

    try {
        $escolhido = Selecionar-ItemDialog -Itens @($item.Candidatos) -Titulo "Escolher usuário" `
            -Descricao "Mais de um usuário para '$($item.Email)'. Selecione o correto:" `
            -FormatarItem { param($u) Obter-RotuloUsuarioDetalhado -Usuario $u }
    }
    catch {
        return
    }

    $item.Usuario = $escolhido
    $item.Status = "Escolhido"
    $item.Detalhe = "Escolhido manualmente entre $(@($item.Candidatos).Count) candidatos."
    Write-Log "Usuario '$($item.Email)' resolvido manualmente: $(Obter-RotuloUsuarioDetalhado -Usuario $escolhido)"
    Atualizar-LinhaRevisaoUsuario -Indice $Indice
    $script:RevisaoUsuariosGrid.Rows[$Indice].Cells["Incluir"].Value = $true
    Atualizar-ResumoRevisaoUsuarios
}

# Mostra uma unica tela com o resultado de todos os e-mails. Retorna os usuarios
# marcados para inclusao, ou $null se o operador cancelar.
function Mostrar-RevisaoUsuariosDialog {
    param([array]$Itens)

    $script:RevisaoUsuariosItens = @($Itens)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Revisar usuários da planilha"
    $dlg.Size = New-Object System.Drawing.Size(1150, 620)
    $dlg.MinimumSize = New-Object System.Drawing.Size(900, 450)
    $dlg.StartPosition = "CenterParent"
    $dlg.ShowInTaskbar = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Dock = [System.Windows.Forms.DockStyle]::Top
    $label.Height = 48
    $label.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 0)
    $script:RevisaoUsuariosLabel = $label

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $script:RevisaoUsuariosGrid = $grid

    $colIncluir = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colIncluir.Name = "Incluir"; $colIncluir.HeaderText = "Incluir"; $colIncluir.FillWeight = 22
    [void]$grid.Columns.Add($colIncluir)
    foreach ($def in @(
        @("Linha", "Linha", 22), @("Email", "E-mail na planilha", 110), @("Situacao", "Situação", 45),
        @("Nome", "Nome no PW", 70), @("EmailPW", "E-mail no PW", 100), @("Detalhe", "Detalhe", 170)
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $def[0]; $col.HeaderText = $def[1]; $col.FillWeight = $def[2]; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }

    for ($i = 0; $i -lt $script:RevisaoUsuariosItens.Count; $i++) {
        $item = $script:RevisaoUsuariosItens[$i]
        $idx = $grid.Rows.Add()
        $grid.Rows[$idx].Cells["Linha"].Value = $item.Linha
        $grid.Rows[$idx].Cells["Email"].Value = $item.Email
        $grid.Rows[$idx].Cells["Incluir"].Value = ($item.Status -eq "Encontrado")
        Atualizar-LinhaRevisaoUsuario -Indice $idx
    }

    $grid.Add_CurrentCellDirtyStateChanged({
        if ($script:RevisaoUsuariosGrid.IsCurrentCellDirty) {
            $script:RevisaoUsuariosGrid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null
        }
    })
    $grid.Add_CellValueChanged({ Atualizar-ResumoRevisaoUsuarios })
    $grid.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -ge 0) { Escolher-CandidatoRevisaoUsuario -Indice $e.RowIndex }
    })

    $painelBotoes = New-Object System.Windows.Forms.FlowLayoutPanel
    $painelBotoes.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $painelBotoes.Height = 46
    $painelBotoes.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $painelBotoes.Padding = New-Object System.Windows.Forms.Padding(6)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"; $btnCancelar.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Size = New-Object System.Drawing.Size(170, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $script:RevisaoUsuariosBtnOk = $btnOk

    $btnEscolher = New-Object System.Windows.Forms.Button
    $btnEscolher.Text = "Escolher usuário..."; $btnEscolher.Size = New-Object System.Drawing.Size(140, 30)
    $btnEscolher.Add_Click({
        $linhaAtual = $script:RevisaoUsuariosGrid.CurrentRow
        if ($linhaAtual) { Escolher-CandidatoRevisaoUsuario -Indice $linhaAtual.Index }
    })

    $btnCopiar = New-Object System.Windows.Forms.Button
    $btnCopiar.Text = "Copiar pendências"; $btnCopiar.Size = New-Object System.Drawing.Size(140, 30)
    $btnCopiar.Add_Click({
        $linhas = @($script:RevisaoUsuariosItens | Where-Object { $_.Status -notin @("Encontrado","Escolhido") } | ForEach-Object {
            "{0}`t{1}`t{2}`t{3}" -f $_.Linha, $_.Email, $script:RotulosStatusUsuario[$_.Status], $_.Detalhe
        })
        if ($linhas.Count -gt 0) {
            [System.Windows.Forms.Clipboard]::SetText("Linha`tE-mail`tSituacao`tDetalhe`r`n" + ($linhas -join "`r`n"))
            Show-UiInfo "$($linhas.Count) pendência(s) copiada(s). Cole no Excel ou no e-mail."
        }
    })

    $painelBotoes.Controls.AddRange(@($btnCancelar, $btnOk, $btnEscolher, $btnCopiar))

    $dlg.Controls.Add($grid)
    $dlg.Controls.Add($label)
    $dlg.Controls.Add($painelBotoes)
    $dlg.CancelButton = $btnCancelar

    Atualizar-ResumoRevisaoUsuarios

    $resultadoDialogo = $dlg.ShowDialog()

    $selecionados = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $item = $script:RevisaoUsuariosItens[$i]
        if ($grid.Rows[$i].Cells["Incluir"].Value -eq $true -and $item.Usuario) {
            $selecionados.Add($item.Usuario)
        }
    }
    $dlg.Dispose()

    if ($resultadoDialogo -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }
    return ,$selecionados.ToArray()
}

function Confirmar-E-AdicionarUsuariosPorEmail {
    param(
        [string[]]$Emails,
        [switch]$SubstituirQuandoUnico,
        $LinhaPorEmail = $null
    )

    $script:UltimaResolucaoUsuarios = $null
    $itens = @(Resolver-ListaEmailsUsuarios -Emails $Emails -LinhaPorEmail $LinhaPorEmail)

    # Um unico e-mail digitado: mantem a confirmacao simples, mas com mensagem clara.
    if ($itens.Count -eq 1 -and -not $LinhaPorEmail) {
        $item = $itens[0]
        switch ($item.Status) {
            "Ambiguo" {
                try {
                    $item.Usuario = Selecionar-ItemDialog -Itens @($item.Candidatos) -Titulo "Escolher usuário" `
                        -Descricao "Mais de um usuário para '$($item.Email)'. Selecione o correto:" `
                        -FormatarItem { param($u) Obter-RotuloUsuarioDetalhado -Usuario $u }
                }
                catch {
                    Write-UiLog "Selecao cancelada para: $($item.Email)"
                    return 0
                }
            }
            { $_ -in @("NaoEncontrado", "Invalido", "Erro") } {
                throw "Usuário '$($item.Email)': $($script:RotulosStatusUsuario[$item.Status]).`r`n`r`n$($item.Detalhe)"
            }
        }

        $u = $item.Usuario
        $nome = Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("Name","UserName","FullName")
        $emailEncontrado = Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("Email","EMail","Mail")
        $login = Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("LoginName","Login","UserID")
        $aviso = if ($item.Status -eq "Parcial") { "ATENÇÃO: o e-mail encontrado é diferente do informado ($($item.Email)).`r`n`r`n" } else { "" }

        if (-not (Ask-UiYesNo -Mensagem "${aviso}Confirmar usuário encontrado?`r`n`r`nNome : $nome`r`nEmail: $emailEncontrado`r`nLogin: $login" -Titulo "Confirmar usuário")) {
            Write-UiLog "Usuario nao confirmado: $($item.Email)"
            return 0
        }

        if (Adicionar-UsuarioSelecionado -Usuario $u -Substituir:$SubstituirQuandoUnico) {
            Write-UiLog "Usuario confirmado: $nome"
            return 1
        }
        Write-UiLog "Usuario ja estava na lista: $nome"
        return 0
    }

    # Lote: uma tela de revisao para tudo, sem interromper no primeiro problema.
    $selecionados = Mostrar-RevisaoUsuariosDialog -Itens $itens

    $caminhoPendencias = $null
    try {
        $caminhoPendencias = Salvar-PendenciasUsuarios -Itens $itens -Caminho (Join-Path $PastaLog ("UsuariosPendentes_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
        if ($caminhoPendencias) {
            Write-Log "Pendencias de usuarios salvas em: $caminhoPendencias" "WARN"
        }
    }
    catch {
        Write-Log "Nao foi possivel salvar o CSV de pendencias: $($_.Exception.Message)" "WARN"
    }

    $script:UltimaResolucaoUsuarios = [PSCustomObject]@{
        Itens             = $itens
        Cancelado         = ($null -eq $selecionados)
        CaminhoPendencias = $caminhoPendencias
    }

    $pendentes = @($itens | Where-Object { $_.Status -notin @("Encontrado","Escolhido") -and -not ($_.Status -eq "Parcial" -and $selecionados -contains $_.Usuario) })
    if ($pendentes.Count -gt 0) {
        Write-UiLog "Usuarios com pendencia ($($pendentes.Count)):"
        foreach ($p in $pendentes) {
            Write-UiLog (" - {0}{1} | {2} | {3}" -f $(if ($p.Linha) { "linha $($p.Linha): " } else { "" }), $p.Email, $script:RotulosStatusUsuario[$p.Status], $p.Detalhe)
        }
    }

    if ($null -eq $selecionados) {
        Write-UiLog "Importacao de usuarios cancelada pelo operador. Nenhum usuario foi adicionado."
        return 0
    }

    $adicionados = 0
    foreach ($u in $selecionados) {
        $nome = Obter-ValorSeguroPropriedade -Objeto $u -PossiveisNomes @("Name","UserName","FullName")
        if (Adicionar-UsuarioSelecionado -Usuario $u) {
            $adicionados++
            Write-UiLog "Usuario confirmado: $nome"
        }
        else {
            Write-UiLog "Usuario ja estava na lista: $nome"
        }
    }

    return $adicionados
}

function Atualizar-ResumoNaTela {
    $script:ListaResumo.Items.Clear()

    foreach ($selecao in $script:SelecoesProjetos) {
        $itemTexto = "{0} | {1} | {2} acesso(s)" -f $selecao.ConcessaoNome, $selecao.ProjetoNome, $selecao.AcessosSelecionados.Count
        [void]$script:ListaResumo.Items.Add($itemTexto)
    }
}

function Atualizar-EstadoInterface {
    $temLogin = $script:InicializacaoConcluida
    $usuariosExecucao = @(Obter-UsuariosParaExecucao)
    $temUsuario = $usuariosExecucao.Count -gt 0
    $temConcessao = $null -ne $script:ConcessaoSelecionada
    $temArvore = $script:TreeProjetos.Nodes.Count -gt 0
    $temSelecoes = $script:SelecoesProjetos.Count -gt 0

    $btnBuscarUsuario.Enabled = $temLogin
    $textEmail.Enabled = $temLogin
    if ($script:CheckModoLoteUsuarios) { $script:CheckModoLoteUsuarios.Enabled = $temLogin }
    if ($script:BtnImportarUsuarios) { $script:BtnImportarUsuarios.Enabled = $temLogin }
    if ($script:BtnRemoverUsuarioSelecionado) { $script:BtnRemoverUsuarioSelecionado.Enabled = $temLogin -and $script:ModoLoteUsuarios -and $script:ListaUsuariosSelecionados.SelectedIndex -ge 0 }
    if ($script:ListaUsuariosSelecionados) { $script:ListaUsuariosSelecionados.Enabled = $temLogin -and $script:ModoLoteUsuarios }

    $script:BtnLocalizarConcessao.Enabled = $temUsuario
    if ($script:BtnImportarProjetosPerfis) { $script:BtnImportarProjetosPerfis.Enabled = $temConcessao -and $temArvore }
    $script:BtnConfirmarSelecao.Enabled = $temConcessao -and $temArvore
    $script:BtnResumo.Enabled = $temSelecoes
    $script:BtnExecutar.Enabled = $temSelecoes
    $script:BtnLimparSelecoes.Enabled = $temSelecoes
    $script:TreeProjetos.Enabled = $temConcessao
    $script:TextFiltroAcessos.Enabled = $temConcessao
    $script:BtnAplicarFiltro.Enabled = $temConcessao
    if ($script:BtnFiltroAvancado) { $script:BtnFiltroAvancado.Enabled = $temConcessao }
    $script:BtnLimparFiltro.Enabled = $temConcessao
    if ($script:BtnSelecionarExibidos) { $script:BtnSelecionarExibidos.Enabled = $temConcessao -and $temArvore }
}

function Marcar-FilhosNode {
    param(
        [System.Windows.Forms.TreeNode]$Node,
        [bool]$Checked
    )

    foreach ($child in $Node.Nodes) {
        $tag = $child.Tag

        if ($tag -and $tag.PSObject.Properties['SomenteVisual'] -and $tag.SomenteVisual) {
            $child.Checked = $false
            continue
        }

        if ($tag -and ($tag.NodeType -eq 'OtherAccess' -or $tag.NodeType -eq 'CategoryInfo')) {
            $child.Checked = $false
            continue
        }

        $child.Checked = $Checked
        if ($child.Nodes.Count -gt 0) {
            Marcar-FilhosNode -Node $child -Checked $Checked
        }
    }
}

function Atualizar-CheckPais {
    param([System.Windows.Forms.TreeNode]$Node)

    $parent = $Node.Parent
    if ($null -eq $parent) {
        return
    }

    $algumMarcado = $false
    foreach ($child in $parent.Nodes) {
        if ($child.Checked) {
            $algumMarcado = $true
            break
        }
    }

    $parent.Checked = $algumMarcado
    Atualizar-CheckPais -Node $parent
}

function Obter-SelecoesDaArvore {
    $mapaProjetos = @{}

    foreach ($projetoNode in $script:TreeProjetos.Nodes) {
        $tagProjeto = $projetoNode.Tag
        if ($null -eq $tagProjeto -or $tagProjeto.NodeType -ne 'Project') {
            continue
        }

        $acessosSelecionados = @()

        foreach ($categoriaNode in $projetoNode.Nodes) {
            foreach ($acessoNode in $categoriaNode.Nodes) {
                if (-not $acessoNode.Checked) {
                    continue
                }

                $tagAcesso = $acessoNode.Tag
                if ($null -eq $tagAcesso -or $tagAcesso.NodeType -ne 'Access') {
                    continue
                }

                $chaveProjeto = "$($tagProjeto.ConcessaoNome)|$($tagProjeto.ProjetoNome)"
                if (-not $mapaProjetos.ContainsKey($chaveProjeto)) {
                    $mapaProjetos[$chaveProjeto] = [PSCustomObject]@{
                        ConcessaoNome       = $tagProjeto.ConcessaoNome
                        ConcessaoObjeto     = $tagProjeto.ConcessaoObjeto
                        ProjetoNome         = $tagProjeto.ProjetoNome
                        ProjetoObjeto       = $tagProjeto.ProjetoObjeto
                        ProjetoId           = $tagProjeto.ProjetoId
                        AcessosSelecionados = @()
                    }
                }

                $mapaProjetos[$chaveProjeto].AcessosSelecionados += [PSCustomObject]@{
                    Tipo     = $tagAcesso.Tipo
                    Nome     = $tagAcesso.Nome
                    ObjetoPw = $tagAcesso.ObjetoPw
                }
            }
        }
    }

    return @(
        $mapaProjetos.Values | Sort-Object ConcessaoNome, ProjetoNome
    )
}

function Mesclar-SelecoesProjetos {
    param(
        [array]$SelecoesExistentes,
        [array]$NovasSelecoes
    )

    $mapaProjetos = @{}
    $mapaAcessosPorProjeto = @{}

    foreach ($selecao in @($SelecoesExistentes) + @($NovasSelecoes)) {
        if ($null -eq $selecao) {
            continue
        }

        $chaveProjeto = Obter-ChaveProjetoSelecao -ConcessaoNome $selecao.ConcessaoNome -ProjetoNome $selecao.ProjetoNome

        if (-not $mapaProjetos.ContainsKey($chaveProjeto)) {
            $mapaProjetos[$chaveProjeto] = [PSCustomObject]@{
                ConcessaoNome       = $selecao.ConcessaoNome
                ConcessaoObjeto     = $selecao.ConcessaoObjeto
                ProjetoNome         = $selecao.ProjetoNome
                ProjetoObjeto       = $selecao.ProjetoObjeto
                ProjetoId           = $selecao.ProjetoId
                AcessosSelecionados = @()
            }
            $mapaAcessosPorProjeto[$chaveProjeto] = @{}
        }

        foreach ($acesso in @($selecao.AcessosSelecionados)) {
            $chaveAcesso = "{0}|{1}" -f (Normalizar-Texto $acesso.Tipo), (Normalizar-Texto $acesso.Nome)

            if ($mapaAcessosPorProjeto[$chaveProjeto].ContainsKey($chaveAcesso)) {
                continue
            }

            $mapaProjetos[$chaveProjeto].AcessosSelecionados += [PSCustomObject]@{
                Tipo     = $acesso.Tipo
                Nome     = $acesso.Nome
                ObjetoPw = $acesso.ObjetoPw
            }
            $mapaAcessosPorProjeto[$chaveProjeto][$chaveAcesso] = $true
        }
    }

    return @(
        $mapaProjetos.Values | Sort-Object ConcessaoNome, ProjetoNome
    )
}

# ---------------------------------------------------------
# CACHE LOCAL DE ACESSOS POR CONCESSAO
# A lista de projetos e sempre lida do PW; o cache guarda apenas os groups /
# user lists de cada projeto (por ID), que mudam raramente. Projetos novos ou
# com cache vencido sao lidos novamente do PW.
# ---------------------------------------------------------
function Obter-CaminhoCacheConcessao {
    param([string]$IdConcessao)

    $datasource = "pw"
    try {
        if (Get-Command Get-PWCurrentDatasource -ErrorAction SilentlyContinue) {
            $ds = Get-PWCurrentDatasource -ErrorAction Stop
            if ($ds) { $datasource = [string]$ds }
        }
    }
    catch { }
    $datasource = $datasource -replace '[^A-Za-z0-9_.-]+', '_'

    $pasta = Join-Path $env:LOCALAPPDATA "CentralScriptsPW\02-GerenciarAcessos\Cache"
    if (-not (Test-Path -LiteralPath $pasta)) {
        New-Item -ItemType Directory -Path $pasta -Force | Out-Null
    }
    return (Join-Path $pasta ("acessos_{0}_{1}.json" -f $datasource, $IdConcessao))
}

function Ler-CacheConcessao {
    param([string]$IdConcessao)

    $mapa = @{}
    try {
        $caminho = Obter-CaminhoCacheConcessao -IdConcessao $IdConcessao
        if (-not (Test-Path -LiteralPath $caminho)) {
            return $mapa
        }

        $dados = Get-Content -LiteralPath $caminho -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($dados.Versao -ne $script:VersaoCacheAcessos) {
            return $mapa
        }
        foreach ($p in @($dados.Projetos)) {
            if ($p -and $p.ProjetoId) { $mapa[[string]$p.ProjetoId] = $p }
        }
    }
    catch {
        Write-Log "Cache de acessos ignorado (arquivo invalido): $($_.Exception.Message)" "WARN"
        $mapa = @{}
    }
    return $mapa
}

function Salvar-CacheConcessao {
    param(
        [string]$IdConcessao,
        [array]$Entradas
    )

    try {
        $caminho = Obter-CaminhoCacheConcessao -IdConcessao $IdConcessao
        $conteudo = [PSCustomObject]@{
            Versao   = $script:VersaoCacheAcessos
            SalvoEm  = (Get-Date).ToString("o")
            Projetos = @($Entradas)
        }
        $json = ConvertTo-Json -InputObject $conteudo -Depth 6 -Compress
        [System.IO.File]::WriteAllText($caminho, $json, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        Write-Log "Nao foi possivel salvar o cache de acessos: $($_.Exception.Message)" "WARN"
    }
}

function Obter-InfoCacheConcessao {
    param([string]$IdConcessao)

    try {
        $caminho = Obter-CaminhoCacheConcessao -IdConcessao $IdConcessao
        if (Test-Path -LiteralPath $caminho) {
            return Get-Item -LiteralPath $caminho
        }
    }
    catch { }
    return $null
}

function Converter-AcessosParaCache {
    param([array]$Acessos)

    return @($Acessos | Where-Object { $_ } | ForEach-Object {
        [PSCustomObject]@{ Tipo = [string]$_.Tipo; Nome = [string]$_.Nome; Origem = [string]$_.Origem }
    })
}

function Converter-AcessosDoCache {
    param($Acessos)

    # ObjetoPw nao e persistido: a execucao usa apenas Tipo/Nome.
    return @(@($Acessos) | Where-Object { $_ -and $_.Nome } | ForEach-Object {
        [PSCustomObject]@{ Tipo = [string]$_.Tipo; Nome = [string]$_.Nome; ObjetoPw = $null; Origem = [string]$_.Origem }
    })
}

function Ler-AcessosProjetoDoPW {
    param(
        [object]$Projeto,
        [string]$NomeProjeto,
        [string]$NomeProjetoTecnico
    )

    $seguranca = Ler-SegurancaProjeto -PastaProjeto $Projeto
    $acessos = Filtrar-AcessosProjeto -ObjetosSeguranca $seguranca["TodosAcessos"]

    $groups = New-Object System.Collections.Generic.List[object]
    $userLists = New-Object System.Collections.Generic.List[object]
    $outrosAcessos = New-Object System.Collections.Generic.List[object]

    foreach ($acesso in $acessos) {
        $tipo = Descobrir-TipoAcesso -Acesso $acesso
        $nomeAcesso = Obter-NomeAcesso -Acesso $acesso
        if ([string]::IsNullOrWhiteSpace($nomeAcesso)) {
            continue
        }

        $itemAcesso = [PSCustomObject]@{
            Tipo     = $tipo
            Nome     = $nomeAcesso
            ObjetoPw = $acesso
            Origem   = "PW"
        }

        if ($tipo -eq 'Group') { $groups.Add($itemAcesso) }
        elseif ($tipo -eq 'UserList') { $userLists.Add($itemAcesso) }
        else { $outrosAcessos.Add($itemAcesso) }
    }

    # FALLBACK INTELIGENTE:
    # Se a API não retornou groups suficientes na segurança da pasta/projeto,
    # tenta localizar grupos existentes pelo padrão do nome do projeto.
    # Isso cobre cenários em que o ProjectWise Administrator mostra groups no nível de projeto,
    # mas o módulo pwps_dab não expõe um cmdlet de Project Security no ambiente.
    if ($groups.Count -lt 3) {
        Write-Log "Projeto '$NomeProjeto' retornou apenas $($groups.Count) group(s) pela API. Tentando fallback por nome..." "WARN"
        # PowerShell 5.1: passar @(List[object]) direto como argumento gera
        # "Os tipos de argumento nao correspondem". Converte antes.
        [object[]]$groupsAtuais = $groups.ToArray()
        $gruposInferidos = @(Obter-GruposInferidosPorProjeto `
            -ProjetoNome $NomeProjeto `
            -ProjetoNomeTecnico $NomeProjetoTecnico `
            -GroupsAtuais $groupsAtuais)

        foreach ($g in $gruposInferidos) { $groups.Add($g) }
    }

    [object[]]$arrGroups = $groups.ToArray()
    [object[]]$arrUserLists = $userLists.ToArray()
    [object[]]$arrOutros = $outrosAcessos.ToArray()
    return [PSCustomObject]@{
        Groups        = $arrGroups
        UserLists     = $arrUserLists
        OutrosAcessos = $arrOutros
    }
}

function Carregar-ArvoreProjetosDaConcessao {
    param(
        [object]$Concessao,
        [switch]$IgnorarCache
    )

    Limpar-ArvoreProjetos
    $script:ProjetosCarregados = @()
    $inicioCarga = Get-Date

    $nomeConcessao = Obter-RotuloConcessao -Concessao $Concessao
    $idConcessao = Obter-IdPasta -Pasta $Concessao

    if ([string]::IsNullOrWhiteSpace($idConcessao)) {
        throw "Nao foi possivel obter o ID da concessao selecionada."
    }

    Write-UiLog "Carregando projetos da concessao '$nomeConcessao'..."

    $filhosConcessao = Get-PWChildFoldersSafe -FolderId $idConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $filhosConcessao -NomesPossiveis $NomesPossiveisPastaProjetos -Descricao "Pasta de projetos da concessao"
    $idPastaProjetos = Obter-IdPasta -Pasta $pastaProjetos

    if ([string]::IsNullOrWhiteSpace($idPastaProjetos)) {
        throw "Nao foi possivel obter o ID da pasta de projetos da concessao."
    }

    $projetos = Get-PWChildFoldersSafe -FolderId $idPastaProjetos
    $projetos = Ordenar-ItensPorNome -Itens $projetos

    if (-not $projetos -or $projetos.Count -eq 0) {
        throw "Nenhum projeto foi encontrado para a concessao selecionada."
    }

    $cache = @{}
    if (-not $IgnorarCache) {
        $cache = Ler-CacheConcessao -IdConcessao $idConcessao
    }
    $entradasCache = New-Object System.Collections.Generic.List[object]
    $listaCarregados = New-Object System.Collections.Generic.List[object]
    $qtdDoCache = 0
    $qtdDoPW = 0
    $total = @($projetos).Count
    $indice = 0

    foreach ($projeto in $projetos) {
        $indice++
        $nomeProjeto = Obter-RotuloProjeto -Projeto $projeto
        $nomeProjetoTecnico = Obter-ValorSeguroPropriedade -Objeto $projeto -PossiveisNomes @("Name","FolderName","ProjectName")
        $idProjeto = Obter-IdPasta -Pasta $projeto

        $entrada = $null
        if ($idProjeto -and $cache.ContainsKey([string]$idProjeto)) {
            $entrada = $cache[[string]$idProjeto]
            $dataCache = [datetime]::MinValue
            if (-not [datetime]::TryParse([string]$entrada.LidoEm, [ref]$dataCache) -or
                ((Get-Date) - $dataCache).TotalDays -gt $script:DiasValidadeCacheAcessos) {
                $entrada = $null
            }
        }

        if ($entrada) {
            $qtdDoCache++
            $groups = Converter-AcessosDoCache $entrada.Groups
            $userLists = Converter-AcessosDoCache $entrada.UserLists
            $outrosAcessos = Converter-AcessosDoCache $entrada.OutrosAcessos
            $lidoEm = [string]$entrada.LidoEm
        }
        else {
            $qtdDoPW++
            Write-UiLog "Lendo acessos do projeto '$nomeProjeto' ($indice/$total)..."
            $lidos = Ler-AcessosProjetoDoPW -Projeto $projeto -NomeProjeto $nomeProjeto -NomeProjetoTecnico $nomeProjetoTecnico
            $groups = $lidos.Groups
            $userLists = $lidos.UserLists
            $outrosAcessos = $lidos.OutrosAcessos
            $lidoEm = (Get-Date).ToString("o")
        }

        if ($idProjeto) {
            $entradasCache.Add([PSCustomObject]@{
                ProjetoId     = [string]$idProjeto
                ProjetoNome   = $nomeProjeto
                LidoEm        = $lidoEm
                Groups        = @(Converter-AcessosParaCache $groups)
                UserLists     = @(Converter-AcessosParaCache $userLists)
                OutrosAcessos = @(Converter-AcessosParaCache $outrosAcessos)
            })
        }

        if (($groups.Count + $userLists.Count + $outrosAcessos.Count) -eq 0) {
            Write-Log "Projeto '$nomeProjeto' ignorado: nenhum acesso com nome identificado." "WARN"
            continue
        }

        if (($groups.Count + $userLists.Count) -eq 0 -and $outrosAcessos.Count -gt 0) {
            Write-Log "Projeto '$nomeProjeto' possui apenas outros acessos. Sera exibido somente para diagnostico visual." "WARN"
        }

        Write-Log "Projeto '$nomeProjeto' carregado para arvore. Groups=$($groups.Count) | UserLists=$($userLists.Count) | Outros=$($outrosAcessos.Count)"

        $listaCarregados.Add([PSCustomObject]@{
            ConcessaoNome     = $nomeConcessao
            ConcessaoObjeto   = $Concessao
            ProjetoNome       = $nomeProjeto
            ProjetoNomeTecnico = $nomeProjetoTecnico
            ProjetoObjeto     = $projeto
            ProjetoId         = $idProjeto
            Groups            = @($groups | Sort-Object Nome)
            UserLists         = @($userLists | Sort-Object Nome)
            OutrosAcessos     = @($outrosAcessos | Sort-Object Tipo, Nome)
        })
    }

    $script:ProjetosCarregados = @($listaCarregados.ToArray())

    # Projetos removidos do PW somem do cache porque so as entradas atuais sao salvas.
    if ($qtdDoPW -gt 0 -or $cache.Count -ne $entradasCache.Count) {
        Salvar-CacheConcessao -IdConcessao $idConcessao -Entradas $entradasCache.ToArray()
    }

    $script:ConcessaoSelecionada = $Concessao
    $script:LabelConcessaoAtual.Text = "Concessão: $nomeConcessao"
    Limpar-CamposFiltroProjetos
    Renderizar-ArvoreProjetos -ProjetosData $script:ProjetosCarregados -Filtro ""

    $segundos = ((Get-Date) - $inicioCarga).TotalSeconds
    Write-Log ("Concessao '{0}' carregada em {1:N1}s. Projetos: {2} | do cache: {3} | lidos do PW: {4}" -f $nomeConcessao, $segundos, $total, $qtdDoCache, $qtdDoPW)
    $origem = if ($qtdDoCache -gt 0) { "; $qtdDoCache do cache local, $qtdDoPW lidos do PW" } else { "" }
    # Mantem o prefixo "Concessao * carregada com sucesso." usado pela Central para esconder a barra de progresso.
    Write-UiLog ("Concessao '{0}' carregada com sucesso. ({1:N0}s{2})" -f $nomeConcessao, $segundos, $origem)
}

# ---------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Gestão de Acessos ProjectWise"
$form.Size = New-Object System.Drawing.Size(1200, 870)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 810)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true

$fontePadrao = New-Object System.Drawing.Font("Segoe UI", 10)

$labelTitulo = New-Object System.Windows.Forms.Label
$labelTitulo.Text = "Gestão de Acessos ProjectWise"
$labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$labelTitulo.AutoSize = $true
$labelTitulo.Location = New-Object System.Drawing.Point(20,15)
$form.Controls.Add($labelTitulo)

$groupUsuario = New-Object System.Windows.Forms.GroupBox
$groupUsuario.Text = "Etapa 1 - Localizar usuário"
$groupUsuario.Font = $fontePadrao
$groupUsuario.Location = New-Object System.Drawing.Point(20,55)
$groupUsuario.Size = New-Object System.Drawing.Size(1140,155)
$form.Controls.Add($groupUsuario)

$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Text = "E-mail:"
$labelEmail.AutoSize = $true
$labelEmail.Location = New-Object System.Drawing.Point(15,30)
$groupUsuario.Controls.Add($labelEmail)

$textEmail = New-Object System.Windows.Forms.TextBox
$textEmail.Location = New-Object System.Drawing.Point(70,27)
$textEmail.Size = New-Object System.Drawing.Size(920,25)
$textEmail.Font = $fontePadrao
$textEmail.Multiline = $true
$textEmail.ScrollBars = "Vertical"
$groupUsuario.Controls.Add($textEmail)

$btnBuscarUsuario = New-Object System.Windows.Forms.Button
$btnBuscarUsuario.Text = "Buscar usuário"
$btnBuscarUsuario.Location = New-Object System.Drawing.Point(995,24)
$btnBuscarUsuario.Size = New-Object System.Drawing.Size(130,30)
$groupUsuario.Controls.Add($btnBuscarUsuario)

$script:CheckModoLoteUsuarios = New-Object System.Windows.Forms.CheckBox
$script:CheckModoLoteUsuarios.Text = "Modo lote de usuários"
$script:CheckModoLoteUsuarios.AutoSize = $true
$script:CheckModoLoteUsuarios.Location = New-Object System.Drawing.Point(15,62)
$groupUsuario.Controls.Add($script:CheckModoLoteUsuarios)

$labelUsuarioEncontrado = New-Object System.Windows.Forms.Label
$labelUsuarioEncontrado.Text = "Usuário: -"
$labelUsuarioEncontrado.AutoSize = $true
$labelUsuarioEncontrado.Location = New-Object System.Drawing.Point(190,64)
$groupUsuario.Controls.Add($labelUsuarioEncontrado)

$script:ListaUsuariosSelecionados = New-Object System.Windows.Forms.ListBox
$script:ListaUsuariosSelecionados.Location = New-Object System.Drawing.Point(15,88)
$script:ListaUsuariosSelecionados.Size = New-Object System.Drawing.Size(970,52)
$script:ListaUsuariosSelecionados.Font = $fontePadrao
$script:ListaUsuariosSelecionados.Enabled = $false
$groupUsuario.Controls.Add($script:ListaUsuariosSelecionados)

$script:BtnRemoverUsuarioSelecionado = New-Object System.Windows.Forms.Button
$script:BtnRemoverUsuarioSelecionado.Text = "Remover"
$script:BtnRemoverUsuarioSelecionado.Location = New-Object System.Drawing.Point(995,88)
$script:BtnRemoverUsuarioSelecionado.Size = New-Object System.Drawing.Size(130,28)
$script:BtnRemoverUsuarioSelecionado.Enabled = $false
$groupUsuario.Controls.Add($script:BtnRemoverUsuarioSelecionado)

$script:BtnImportarUsuarios = New-Object System.Windows.Forms.Button
$script:BtnImportarUsuarios.Text = "Importar planilha"
$script:BtnImportarUsuarios.Location = New-Object System.Drawing.Point(995,120)
$script:BtnImportarUsuarios.Size = New-Object System.Drawing.Size(130,28)
$script:BtnImportarUsuarios.Enabled = $false
$groupUsuario.Controls.Add($script:BtnImportarUsuarios)

$groupSelecao = New-Object System.Windows.Forms.GroupBox
$groupSelecao.Text = "Etapa 2 - Concessão, projetos, grupos e listas de usuários"
$groupSelecao.Font = $fontePadrao
$groupSelecao.Location = New-Object System.Drawing.Point(20,225)
$groupSelecao.Size = New-Object System.Drawing.Size(1140,330)
$form.Controls.Add($groupSelecao)

$script:BtnLocalizarConcessao = New-Object System.Windows.Forms.Button
$script:BtnLocalizarConcessao.Text = "Selecionar concessão"
$script:BtnLocalizarConcessao.Location = New-Object System.Drawing.Point(15,28)
$script:BtnLocalizarConcessao.Size = New-Object System.Drawing.Size(150,32)
$script:BtnLocalizarConcessao.Enabled = $false
$groupSelecao.Controls.Add($script:BtnLocalizarConcessao)

$script:BtnImportarProjetosPerfis = New-Object System.Windows.Forms.Button
$script:BtnImportarProjetosPerfis.Text = "Importar projetos"
$script:BtnImportarProjetosPerfis.Location = New-Object System.Drawing.Point(970,28)
$script:BtnImportarProjetosPerfis.Size = New-Object System.Drawing.Size(150,32)
$script:BtnImportarProjetosPerfis.Enabled = $false
$groupSelecao.Controls.Add($script:BtnImportarProjetosPerfis)

$script:BtnConfirmarSelecao = New-Object System.Windows.Forms.Button
$script:BtnConfirmarSelecao.Text = "Confirmar seleções"
$script:BtnConfirmarSelecao.Location = New-Object System.Drawing.Point(970,288)
$script:BtnConfirmarSelecao.Size = New-Object System.Drawing.Size(150,32)
$script:BtnConfirmarSelecao.Enabled = $false
$groupSelecao.Controls.Add($script:BtnConfirmarSelecao)

$script:LabelConcessaoAtual = New-Object System.Windows.Forms.Label
$script:LabelConcessaoAtual.Text = "Concessão: -"
$script:LabelConcessaoAtual.AutoSize = $true
$script:LabelConcessaoAtual.Location = New-Object System.Drawing.Point(180,35)
$groupSelecao.Controls.Add($script:LabelConcessaoAtual)

$labelTreeInfo = New-Object System.Windows.Forms.Label
$labelTreeInfo.Text = "Marque os grupos e listas de usuários desejados abaixo de cada projeto."
$labelTreeInfo.AutoSize = $true
$labelTreeInfo.Location = New-Object System.Drawing.Point(15,72)
$groupSelecao.Controls.Add($labelTreeInfo)

$labelFiltroAcessos = New-Object System.Windows.Forms.Label
$labelFiltroAcessos.Text = "Filtrar projetos, grupos ou listas de usuários:"
$labelFiltroAcessos.AutoSize = $true
$labelFiltroAcessos.Location = New-Object System.Drawing.Point(15,102)
$groupSelecao.Controls.Add($labelFiltroAcessos)

$script:TextFiltroAcessos = New-Object System.Windows.Forms.TextBox
$script:TextFiltroAcessos.Location = New-Object System.Drawing.Point(285,98)
$script:TextFiltroAcessos.Size = New-Object System.Drawing.Size(730,25)
$script:TextFiltroAcessos.Font = $fontePadrao
$script:TextFiltroAcessos.Enabled = $false
$script:TextFiltroAcessos.ShortcutsEnabled = $true
$groupSelecao.Controls.Add($script:TextFiltroAcessos)

$script:BtnAplicarFiltro = New-Object System.Windows.Forms.Button
$script:BtnAplicarFiltro.Text = "Filtrar"
$script:BtnAplicarFiltro.Location = New-Object System.Drawing.Point(805,95)
$script:BtnAplicarFiltro.Size = New-Object System.Drawing.Size(90,30)
$script:BtnAplicarFiltro.Enabled = $false
$groupSelecao.Controls.Add($script:BtnAplicarFiltro)

$script:BtnFiltroAvancado = New-Object System.Windows.Forms.Button
$script:BtnFiltroAvancado.Text = "Filtro avançado"
$script:BtnFiltroAvancado.Location = New-Object System.Drawing.Point(905,95)
$script:BtnFiltroAvancado.Size = New-Object System.Drawing.Size(130,30)
$script:BtnFiltroAvancado.Enabled = $false
$groupSelecao.Controls.Add($script:BtnFiltroAvancado)

$script:BtnLimparFiltro = New-Object System.Windows.Forms.Button
$script:BtnLimparFiltro.Text = "Limpar filtro"
$script:BtnLimparFiltro.Location = New-Object System.Drawing.Point(1015,95)
$script:BtnLimparFiltro.Size = New-Object System.Drawing.Size(110,30)
$script:BtnLimparFiltro.Enabled = $false
$groupSelecao.Controls.Add($script:BtnLimparFiltro)

$script:LabelResultadoFiltro = New-Object System.Windows.Forms.Label
$script:LabelResultadoFiltro.Text = "Projetos exibidos: 0"
$script:LabelResultadoFiltro.AutoSize = $true
$script:LabelResultadoFiltro.Location = New-Object System.Drawing.Point(15,132)
$groupSelecao.Controls.Add($script:LabelResultadoFiltro)

$script:BtnSelecionarExibidos = New-Object System.Windows.Forms.Button
$script:BtnSelecionarExibidos.Text = "Selecionar exibidos"
$script:BtnSelecionarExibidos.Location = New-Object System.Drawing.Point(970,126)
$script:BtnSelecionarExibidos.Size = New-Object System.Drawing.Size(150,28)
$script:BtnSelecionarExibidos.Enabled = $false
$groupSelecao.Controls.Add($script:BtnSelecionarExibidos)

$script:TreeProjetos = New-Object System.Windows.Forms.TreeView
$script:TreeProjetos.Location = New-Object System.Drawing.Point(15,155)
$script:TreeProjetos.Size = New-Object System.Drawing.Size(1105,125)
$script:TreeProjetos.CheckBoxes = $true
$script:TreeProjetos.Font = $fontePadrao
$script:TreeProjetos.HideSelection = $false
$script:TreeProjetos.Enabled = $false
$groupSelecao.Controls.Add($script:TreeProjetos)

$groupProjetos = New-Object System.Windows.Forms.GroupBox
$groupProjetos.Text = "Etapa 3 - Resumo das seleções confirmadas"
$groupProjetos.Font = $fontePadrao
$groupProjetos.Location = New-Object System.Drawing.Point(20,570)
$groupProjetos.Size = New-Object System.Drawing.Size(1140,150)
$form.Controls.Add($groupProjetos)

$script:ListaResumo = New-Object System.Windows.Forms.ListBox
$script:ListaResumo.Location = New-Object System.Drawing.Point(15,28)
$script:ListaResumo.Size = New-Object System.Drawing.Size(1105,78)
$script:ListaResumo.Font = $fontePadrao
$groupProjetos.Controls.Add($script:ListaResumo)

$script:BtnResumo = New-Object System.Windows.Forms.Button
$script:BtnResumo.Text = "Ver resumo"
$script:BtnResumo.Location = New-Object System.Drawing.Point(15,112)
$script:BtnResumo.Size = New-Object System.Drawing.Size(120,28)
$script:BtnResumo.Enabled = $false
$groupProjetos.Controls.Add($script:BtnResumo)

$script:BtnLimparSelecoes = New-Object System.Windows.Forms.Button
$script:BtnLimparSelecoes.Text = "Limpar seleções"
$script:BtnLimparSelecoes.Location = New-Object System.Drawing.Point(150,112)
$script:BtnLimparSelecoes.Size = New-Object System.Drawing.Size(140,28)
$script:BtnLimparSelecoes.Enabled = $false
$groupProjetos.Controls.Add($script:BtnLimparSelecoes)

$script:BtnExecutar = New-Object System.Windows.Forms.Button
$script:BtnExecutar.Text = "Executar permissões"
$script:BtnExecutar.Location = New-Object System.Drawing.Point(945,112)
$script:BtnExecutar.Size = New-Object System.Drawing.Size(160,28)
$script:BtnExecutar.Enabled = $false
$groupProjetos.Controls.Add($script:BtnExecutar)

$groupLog = New-Object System.Windows.Forms.GroupBox
$groupLog.Text = "Acompanhamento"
$groupLog.Font = $fontePadrao
$groupLog.Location = New-Object System.Drawing.Point(20,730)
$groupLog.Size = New-Object System.Drawing.Size(1140,95)
$form.Controls.Add($groupLog)

$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Multiline = $true
$script:LogBox.ReadOnly = $true
$script:LogBox.ScrollBars = "Vertical"
$script:LogBox.Location = New-Object System.Drawing.Point(15,25)
$script:LogBox.Size = New-Object System.Drawing.Size(1105,55)
$script:LogBox.Font = New-Object System.Drawing.Font("Consolas",9)
$groupLog.Controls.Add($script:LogBox)

$btnFechar = New-Object System.Windows.Forms.Button
$btnFechar.Text = "Fechar"
$btnFechar.Location = New-Object System.Drawing.Point(1070,20)
$btnFechar.Size = New-Object System.Drawing.Size(90,28)
$btnFechar.Add_Click({ $form.Close() })
$form.Controls.Add($btnFechar)


function Ajustar-LayoutCampos {
    if (-not $groupUsuario -or -not $groupSelecao) {
        return
    }

    $margemDireita = 15
    $espacamento = 10

    if ($textEmail -and $btnBuscarUsuario) {
        $btnBuscarUsuario.Left = $groupUsuario.ClientSize.Width - $btnBuscarUsuario.Width - $margemDireita
        $textEmail.Width = [Math]::Max(220, $btnBuscarUsuario.Left - $textEmail.Left - $espacamento)
    }

    if ($script:ListaUsuariosSelecionados -and $script:BtnRemoverUsuarioSelecionado) {
        $script:BtnRemoverUsuarioSelecionado.Left = $groupUsuario.ClientSize.Width - $script:BtnRemoverUsuarioSelecionado.Width - $margemDireita
        if ($script:BtnImportarUsuarios) {
            $script:BtnImportarUsuarios.Left = $script:BtnRemoverUsuarioSelecionado.Left
        }
        $script:ListaUsuariosSelecionados.Width = [Math]::Max(320, $script:BtnRemoverUsuarioSelecionado.Left - $script:ListaUsuariosSelecionados.Left - $espacamento)
    }

    if ($script:TextFiltroAcessos -and $script:BtnAplicarFiltro -and $script:BtnLimparFiltro) {
        $script:BtnLimparFiltro.Left = $groupSelecao.ClientSize.Width - $script:BtnLimparFiltro.Width - $margemDireita
        if ($script:BtnFiltroAvancado) {
            $script:BtnFiltroAvancado.Left = $script:BtnLimparFiltro.Left - $script:BtnFiltroAvancado.Width - $espacamento
            $script:BtnAplicarFiltro.Left = $script:BtnFiltroAvancado.Left - $script:BtnAplicarFiltro.Width - $espacamento
        }
        else {
            $script:BtnAplicarFiltro.Left = $script:BtnLimparFiltro.Left - $script:BtnAplicarFiltro.Width - $espacamento
        }

        $filtroLeft = $script:TextFiltroAcessos.Left
        if ($labelFiltroAcessos) {
            $filtroLeft = $labelFiltroAcessos.Left + $labelFiltroAcessos.Width + 8
            $script:TextFiltroAcessos.Left = $filtroLeft
        }

        $script:TextFiltroAcessos.Width = [Math]::Max(260, $script:BtnAplicarFiltro.Left - $filtroLeft - $espacamento)
    }

    if ($script:BtnSelecionarExibidos) {
        $script:BtnSelecionarExibidos.Left = $groupSelecao.ClientSize.Width - $script:BtnSelecionarExibidos.Width - $margemDireita
    }

    if ($script:BtnImportarProjetosPerfis) {
        $script:BtnImportarProjetosPerfis.Left = $groupSelecao.ClientSize.Width - $script:BtnImportarProjetosPerfis.Width - $margemDireita
    }

    if ($script:BtnConfirmarSelecao) {
        $script:BtnConfirmarSelecao.Left = $groupSelecao.ClientSize.Width - $script:BtnConfirmarSelecao.Width - $margemDireita
        $script:BtnConfirmarSelecao.Top = $groupSelecao.ClientSize.Height - $script:BtnConfirmarSelecao.Height - 10
    }
}


$form.Add_Shown({ Ajustar-LayoutCampos })
$form.Add_Resize({ Ajustar-LayoutCampos })
$groupUsuario.Add_Resize({ Ajustar-LayoutCampos })
$groupSelecao.Add_Resize({ Ajustar-LayoutCampos })

# ---------------------------------------------------------
# AJUSTES DE REDIMENSIONAMENTO
# ---------------------------------------------------------
$groupUsuario.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$groupSelecao.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$groupProjetos.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$groupLog.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$textEmail.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$btnBuscarUsuario.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:CheckModoLoteUsuarios.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$labelUsuarioEncontrado.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:ListaUsuariosSelecionados.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnRemoverUsuarioSelecionado.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnImportarUsuarios.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$script:BtnLocalizarConcessao.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$script:BtnImportarProjetosPerfis.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnConfirmarSelecao.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$script:LabelConcessaoAtual.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TextFiltroAcessos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TextFiltroAcessos.BringToFront()

$script:BtnAplicarFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnFiltroAvancado.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnLimparFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnSelecionarExibidos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:LabelResultadoFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TreeProjetos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$script:ListaResumo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnResumo.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$script:BtnLimparSelecoes.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$script:BtnExecutar.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$script:LogBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$btnFechar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# ---------------------------------------------------------
# EVENTOS
# ---------------------------------------------------------
$script:TreeProjetos.Add_AfterCheck({
    param($sender, $e)

    if ($script:AtualizandoCheckTree) {
        return
    }

    try {
        $script:AtualizandoCheckTree = $true
        if ($e.Node.Nodes.Count -gt 0) {
            Marcar-FilhosNode -Node $e.Node -Checked $e.Node.Checked
        }
        Atualizar-CheckPais -Node $e.Node
    }
    finally {
        $script:AtualizandoCheckTree = $false
    }
})

$script:BtnAplicarFiltro.Add_Click({
    try {
        $script:FiltroExclusaoAvancado = ""
        Aplicar-FiltroProjetos
    }
    catch {
        Write-Log "Erro ao aplicar filtro. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnFiltroAvancado.Add_Click({
    try {
        Mostrar-ModalFiltroAvancado
    }
    catch {
        Write-Log "Erro ao abrir filtro avancado. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnLimparFiltro.Add_Click({
    try {
        Limpar-CamposFiltroProjetos
        Aplicar-FiltroProjetos
    }
    catch {
        Write-Log "Erro ao limpar filtro. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnSelecionarExibidos.Add_Click({
    try {
        $totalMarcados = Selecionar-AcessosExibidosNaArvore

        $temFiltroBusca = $script:TextFiltroAcessos -and -not [string]::IsNullOrWhiteSpace($script:TextFiltroAcessos.Text)
        $temFiltroExclusao = -not [string]::IsNullOrWhiteSpace($script:FiltroExclusaoAvancado)

        if ($temFiltroBusca -or $temFiltroExclusao) {
            Write-UiLog "Acessos exibidos pelo filtro selecionados: $totalMarcados"
        }
        else {
            Write-UiLog "Acessos exibidos selecionados: $totalMarcados"
        }

        Atualizar-EstadoInterface
    }
    catch {
        Write-Log "Erro ao selecionar acessos exibidos. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnImportarProjetosPerfis.Add_Click({
    try {
        if (-not $script:ConcessaoSelecionada -or -not $script:ProjetosCarregados -or $script:ProjetosCarregados.Count -eq 0) {
            throw "Selecione uma concessao antes de importar a lista de projetos."
        }

        $caminhoPlanilha = Selecionar-PlanilhaProjetos
        Write-UiLog "Importando lista de projetos: $caminhoPlanilha"

        $projetosPlanilha = @(Importar-NomesProjetos -CaminhoArquivo $caminhoPlanilha)
        $resumoImportacao = $script:ResumoImportacaoProjetos
        if ($resumoImportacao) {
            Write-UiLog "Linhas na planilha de projetos: $($resumoImportacao.LinhasPlanilha)"
            Write-UiLog "Projetos unicos importados: $($resumoImportacao.ProjetosUnicos) | Duplicados ignorados: $($resumoImportacao.DuplicadosIgnorados)"
            if (@($resumoImportacao.LinhasSemProjeto).Count -gt 0) {
                Write-UiLog "Linhas sem projeto ignoradas: $(@($resumoImportacao.LinhasSemProjeto) -join ', ')"
            }
        }

        $perfisSelecionados = @(Mostrar-ModalSelecaoPerfisProjeto)
        if ($perfisSelecionados.Count -eq 0) {
            Write-UiLog "Importacao de projetos cancelada: nenhum perfil selecionado."
            return
        }

        Write-UiLog "Perfis selecionados: $(@($perfisSelecionados | ForEach-Object { $_.Sufixo }) -join ', ')"
        $resumo = Aplicar-PerfisEmProjetosImportados -ProjetosPlanilha $projetosPlanilha -Perfis $perfisSelecionados

        Write-UiLog "Projetos encontrados: $(@($resumo.ProjetosEncontrados).Count) de $($resumo.ProjetosPlanilha)"
        Write-UiLog "Acessos marcados automaticamente: $($resumo.AcessosMarcados)"
        if (@($resumo.ProjetosNaoEncontrados).Count -gt 0) {
            Write-UiLog "Projetos nao encontrados: $(@($resumo.ProjetosNaoEncontrados).Count)"
        }
        if (@($resumo.PerfisNaoEncontrados).Count -gt 0) {
            Write-UiLog "Pendencias de perfil: $(@($resumo.PerfisNaoEncontrados).Count)"
        }

        Mostrar-ResumoImportacaoProjetosPerfis -Resumo $resumo
    }
    catch {
        Write-Log "Erro ao importar projetos/perfis. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:TextFiltroAcessos.Add_KeyDown({
    param($sender, $e)

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
        return
    }

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        try {
            $script:FiltroExclusaoAvancado = ""
            Aplicar-FiltroProjetos
            $e.SuppressKeyPress = $true
        }
        catch {
            Write-Log "Erro ao aplicar filtro por teclado. Detalhe: $($_.Exception.Message)" "ERROR"
            Show-UiError $_.Exception.Message
        }
    }
})

$script:CheckModoLoteUsuarios.Add_CheckedChanged({
    try {
        $script:ModoLoteUsuarios = $script:CheckModoLoteUsuarios.Checked
        if ($script:ModoLoteUsuarios) {
            $btnBuscarUsuario.Text = "Adicionar usuário(s)"
            $labelEmail.Text = "E-mail(s):"
            $textEmail.Height = 50
            Write-UiLog "Modo lote de usuarios ativado."
        }
        else {
            $btnBuscarUsuario.Text = "Buscar usuário"
            $labelEmail.Text = "E-mail:"
            $textEmail.Height = 25

            if ($script:UsuariosSelecionados.Count -gt 1) {
                $manter = @($script:UsuariosSelecionados | Select-Object -First 1)
                $script:UsuariosSelecionados = $manter
                $script:UsuarioEncontrado = $manter[0]
                Write-UiLog "Modo lote desativado. Mantido somente o primeiro usuario confirmado."
            }
        }

        Atualizar-ListaUsuariosSelecionados
        Atualizar-EstadoInterface
        Ajustar-LayoutCampos
    }
    catch {
        Write-Log "Erro ao alternar modo lote. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:ListaUsuariosSelecionados.Add_SelectedIndexChanged({
    Atualizar-EstadoInterface
})

$script:BtnRemoverUsuarioSelecionado.Add_Click({
    try {
        $indice = $script:ListaUsuariosSelecionados.SelectedIndex
        if ($indice -lt 0) {
            return
        }

        $usuarioRemovido = $script:UsuariosSelecionados[$indice]
        $novaLista = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $script:UsuariosSelecionados.Count; $i++) {
            if ($i -ne $indice) {
                $novaLista.Add($script:UsuariosSelecionados[$i])
            }
        }

        $script:UsuariosSelecionados = @($novaLista.ToArray())
        $script:UsuarioEncontrado = @($script:UsuariosSelecionados | Select-Object -First 1)[0]
        Write-UiLog "Usuario removido da lista: $(Obter-RotuloUsuario -Usuario $usuarioRemovido)"

        Atualizar-ListaUsuariosSelecionados
        Atualizar-EstadoInterface
    }
    catch {
        Write-Log "Erro ao remover usuario da lista. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnImportarUsuarios.Add_Click({
    try {
        if (-not $script:ModoLoteUsuarios) {
            $script:CheckModoLoteUsuarios.Checked = $true
            $script:ModoLoteUsuarios = $true
        }

        $caminhoPlanilha = Selecionar-PlanilhaUsuarios
        Write-UiLog "Importando usuarios da planilha: $caminhoPlanilha"

        $emails = @(Importar-EmailsUsuarios -CaminhoArquivo $caminhoPlanilha)
        $resumoImportacao = $script:ResumoImportacaoUsuarios
        if ($resumoImportacao) {
            Write-UiLog "Linhas na planilha: $($resumoImportacao.LinhasPlanilha)"
            Write-UiLog "E-mails encontrados: $($resumoImportacao.TotalEmailsEncontrados) | E-mails unicos: $($resumoImportacao.EmailsUnicos) | Duplicados ignorados: $($resumoImportacao.DuplicadosIgnorados)"

            if ($resumoImportacao.DuplicadosIgnorados -gt 0) {
                Write-UiLog "E-mails duplicados ignorados:"
                foreach ($emailDuplicado in @($resumoImportacao.EmailsDuplicados)) {
                    Write-UiLog " - $emailDuplicado"
                }
            }

            if (@($resumoImportacao.LinhasSemEmail).Count -gt 0) {
                Write-UiLog "Linhas sem e-mail ignoradas: $(@($resumoImportacao.LinhasSemEmail) -join ', ')"
            }
        }
        else {
            Write-UiLog "E-mails encontrados na planilha: $($emails.Count)"
        }

        $usuariosAntes = @($script:UsuariosSelecionados).Count
        $linhaPorEmail = $null
        if ($resumoImportacao) { $linhaPorEmail = $resumoImportacao.LinhaPorEmail }
        if (-not $linhaPorEmail) { $linhaPorEmail = @{} }
        $adicionados = Confirmar-E-AdicionarUsuariosPorEmail -Emails $emails -LinhaPorEmail $linhaPorEmail

        if (@($script:UsuariosSelecionados).Count -gt 0 -and $usuariosAntes -eq 0) {
            $script:ConcessaoSelecionada = $null
            $script:SelecoesProjetos = @()
            $script:ProjetosCarregados = @()
            Limpar-ArvoreProjetos
            Limpar-CamposFiltroProjetos
            Atualizar-ResumoNaTela
            $script:LabelConcessaoAtual.Text = "Concessão: -"
        }

        Atualizar-ListaUsuariosSelecionados
        Atualizar-EstadoInterface
        $mensagemResumo = "Importacao concluida.`r`n`r`n"
        if ($resumoImportacao) {
            $mensagemResumo += "Linhas na planilha: $($resumoImportacao.LinhasPlanilha)`r`n"
            $mensagemResumo += "E-mails encontrados: $($resumoImportacao.TotalEmailsEncontrados)`r`n"
            $mensagemResumo += "E-mails unicos importados: $($resumoImportacao.EmailsUnicos)`r`n"
            $mensagemResumo += "Duplicados ignorados: $($resumoImportacao.DuplicadosIgnorados)`r`n"
        }
        else {
            $mensagemResumo += "E-mails encontrados: $($emails.Count)`r`n"
        }
        $mensagemResumo += "Usuarios adicionados: $adicionados`r`n"
        $mensagemResumo += "Usuarios confirmados na lista: $($script:UsuariosSelecionados.Count)"

        $resolucao = $script:UltimaResolucaoUsuarios
        if ($resolucao) {
            if ($resolucao.Cancelado) {
                $mensagemResumo = "Importacao cancelada. Nenhum usuario foi adicionado.`r`n`r`n" + $mensagemResumo
            }

            $naoIncluidos = @($resolucao.Itens | Where-Object { $_.Status -in @("NaoEncontrado","Invalido","Erro","Ambiguo") })
            if ($naoIncluidos.Count -gt 0) {
                $mensagemResumo += "`r`n`r`nNAO incluidos ($($naoIncluidos.Count)):"
                foreach ($p in @($naoIncluidos | Select-Object -First 15)) {
                    $mensagemResumo += "`r`n - " + $(if ($p.Linha) { "linha $($p.Linha): " } else { "" }) + "$($p.Email) ($($script:RotulosStatusUsuario[$p.Status]))"
                }
                if ($naoIncluidos.Count -gt 15) {
                    $mensagemResumo += "`r`n ... e mais $($naoIncluidos.Count - 15)."
                }
            }
            if ($resolucao.CaminhoPendencias) {
                $mensagemResumo += "`r`n`r`nLista completa de pendencias salva em:`r`n$($resolucao.CaminhoPendencias)"
            }
        }
        if (@($resumoImportacao.LinhasSemEmail).Count -gt 0) {
            $mensagemResumo += "`r`n`r`nLinhas sem e-mail ignoradas: $(@($resumoImportacao.LinhasSemEmail) -join ', ')"
        }

        Show-UiInfo $mensagemResumo
    }
    catch {
        Write-Log "Erro ao importar usuarios. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$btnBuscarUsuario.Add_Click({
    try {
        $emails = @(Obter-EmailsInformados -Texto $textEmail.Text)
        if ($emails.Count -eq 0) {
            throw "Informe ao menos um e-mail."
        }

        if (-not $script:ModoLoteUsuarios -and $emails.Count -gt 1) {
            throw "Para buscar varios usuarios, marque a opcao 'Modo lote de usuários'."
        }

        $usuariosAntes = @($script:UsuariosSelecionados).Count

        $null = Confirmar-E-AdicionarUsuariosPorEmail -Emails $emails -SubstituirQuandoUnico:(!$script:ModoLoteUsuarios)

        if (@($script:UsuariosSelecionados).Count -gt 0 -and ($usuariosAntes -eq 0 -or -not $script:ModoLoteUsuarios)) {
            $script:ConcessaoSelecionada = $null
            $script:SelecoesProjetos = @()
            $script:ProjetosCarregados = @()
            Limpar-ArvoreProjetos
            Limpar-CamposFiltroProjetos
            Atualizar-ResumoNaTela
            $script:LabelConcessaoAtual.Text = "Concessão: -"
        }

        if ($script:ModoLoteUsuarios) {
            $textEmail.Text = ""
        }

        Atualizar-ListaUsuariosSelecionados
        Atualizar-EstadoInterface
    }
    catch {
        Write-Log "Erro ao buscar usuario. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnLocalizarConcessao.Add_Click({
    try {
        if (@(Obter-UsuariosParaExecucao).Count -eq 0) {
            throw "Nenhum usuario foi confirmado."
        }

        Write-UiLog "Selecionando concessao..."
        $pastasRaiz = Get-PWRootFoldersSafe
        $pastaEngenhariaRaiz = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $NomesPossiveisEngenhariaRaiz -Descricao "Pasta de engenharia na raiz"
        $idEngenharia = Obter-IdPasta -Pasta $pastaEngenhariaRaiz

        if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
            throw "Nao foi possivel obter o ID da pasta de engenharia na raiz."
        }

        $concessoes = Get-PWChildFoldersSafe -FolderId $idEngenharia
        $concessoes = Ordenar-ItensPorNome -Itens $concessoes

        $concessaoSelecionada = Selecionar-ItemDialog -Itens $concessoes -Titulo "Selecionar concessão" -Descricao "Escolha a concessão:"

        $script:SelecoesProjetos = @()

        $ignorarCache = $false
        $infoCache = Obter-InfoCacheConcessao -IdConcessao (Obter-IdPasta -Pasta $concessaoSelecionada)
        if ($infoCache) {
            $usarCache = Ask-UiYesNo -Titulo "Acessos salvos" -Mensagem (
                "Existem acessos desta concessão salvos neste computador em $($infoCache.LastWriteTime.ToString('dd/MM/yyyy HH:mm')).`r`n`r`n" +
                "Sim  = carregamento rápido (a lista de projetos é consultada no PW; só projetos novos ou com mais de $($script:DiasValidadeCacheAcessos) dias são relidos).`r`n" +
                "Não = reler os acessos de todos os projetos no ProjectWise (mais lento).`r`n`r`n" +
                "Use 'Não' se grupos ou user lists foram criados/alterados recentemente.")
            $ignorarCache = -not $usarCache
        }

        Carregar-ArvoreProjetosDaConcessao -Concessao $concessaoSelecionada -IgnorarCache:$ignorarCache
        Atualizar-ResumoNaTela
        Atualizar-EstadoInterface
    }
    catch {
        Write-Log "Erro ao localizar concessao. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnConfirmarSelecao.Add_Click({
    try {
        $usuariosExecucao = @(Obter-UsuariosParaExecucao)
        if ($usuariosExecucao.Count -eq 0) {
            throw "Nenhum usuario foi confirmado."
        }

        if (-not $script:ConcessaoSelecionada) {
            throw "Nenhuma concessao foi selecionada."
        }

        $selecoes = Obter-SelecoesDaArvore
        if (-not $selecoes -or $selecoes.Count -eq 0) {
            throw "Nenhum group ou user list foi marcado na arvore."
        }

        $selecoesMescladas = @(Mesclar-SelecoesProjetos -SelecoesExistentes $script:SelecoesProjetos -NovasSelecoes $selecoes)
        Mostrar-ResumoConsolidadoDialog -Usuarios $usuariosExecucao -SelecoesProjetos $selecoesMescladas

        $script:SelecoesProjetos = @($selecoesMescladas)
        Atualizar-ResumoNaTela
        Atualizar-EstadoInterface
        Write-UiLog "Selecoes confirmadas/atualizadas com sucesso."
    }
    catch {
        Write-Log "Erro ao confirmar selecoes. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnResumo.Add_Click({
    try {
        Mostrar-ResumoConsolidadoDialog -Usuarios @(Obter-UsuariosParaExecucao) -SelecoesProjetos $script:SelecoesProjetos
    }
    catch {
        Show-UiError $_.Exception.Message
    }
})

$script:BtnLimparSelecoes.Add_Click({
    if (Ask-UiYesNo -Mensagem "Deseja realmente limpar as selecoes confirmadas atuais?" -Titulo "Limpar seleções") {
        $script:SelecoesProjetos = @()
        Atualizar-ResumoNaTela
        Atualizar-EstadoInterface
        Write-UiLog "Selecoes confirmadas limpas."
    }
})

$script:BtnExecutar.Add_Click({
    try {
        $usuariosExecucao = @(Obter-UsuariosParaExecucao)
        if ($usuariosExecucao.Count -eq 0) {
            throw "Nenhum usuario foi confirmado."
        }

        if (-not $script:SelecoesProjetos -or $script:SelecoesProjetos.Count -eq 0) {
            throw "Nenhuma selecao confirmada foi encontrada."
        }

        $totalProjetos = $script:SelecoesProjetos.Count
        $totalAcessos = 0
        foreach ($selecao in $script:SelecoesProjetos) {
            $totalAcessos += @($selecao.AcessosSelecionados).Count
        }
        $totalUsuarios = $usuariosExecucao.Count
        $totalOperacoes = $totalUsuarios * $totalAcessos

        $executar = Ask-UiYesNo -Mensagem "Deseja executar agora a inclusao dos usuarios nos acessos selecionados?`r`n`r`nUsuarios : $totalUsuarios`r`nProjetos : $totalProjetos`r`nAcessos  : $totalAcessos`r`nOperacoes estimadas: $totalOperacoes" -Titulo "Executar permissões"
        if (-not $executar) {
            Write-UiLog "Execucao cancelada pelo operador."
            return
        }

        Write-UiLog "Executando inclusoes..."
        $script:ResultadosExecucao = @()
        foreach ($usuarioExecucao in $usuariosExecucao) {
            $rotuloUsuario = Obter-RotuloUsuario -Usuario $usuarioExecucao
            Write-UiLog "Executando inclusoes para: $rotuloUsuario"
            $script:ResultadosExecucao += @(Executar-InclusoesSelecionadas -Usuario $usuarioExecucao -SelecoesProjetos $script:SelecoesProjetos)
        }

        Mostrar-ResultadoExecucaoDialog -Resultados $script:ResultadosExecucao
        Mostrar-ResumoFinalExecucaoDialog -Resultados $script:ResultadosExecucao

        $sucesso = @($script:ResultadosExecucao | Where-Object { $_.Status -eq "SUCESSO" }).Count
        $jaExiste = @($script:ResultadosExecucao | Where-Object { $_.Status -eq "JA_EXISTE" }).Count
        $erro = @($script:ResultadosExecucao | Where-Object { $_.Status -eq "ERRO" }).Count

        Write-UiLog "Execucao concluida. Sucesso: $sucesso | Ja existe: $jaExiste | Erro: $erro"
        Show-UiInfo "Execucao concluida."
    }
    catch {
        Write-Log "Erro na execucao final. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$form.Add_Shown({
    if ($script:InicializacaoConcluida) {
        return
    }

    try {
        Inicializar-Aplicacao
        Atualizar-EstadoInterface
    }
    catch {
        Write-Log "Falha na inicializacao. Detalhe: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Falha ao iniciar a aplicacao.`r`n`r`n$($_.Exception.Message)",
            "Erro na inicializacao",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $form.Close()
    }
})

$form.Add_FormClosing({
    try {
        Encerrar-SessaoProjectWise -LoginPW $script:LoginPW
    }
    catch {
    }
})

# ---------------------------------------------------------
# INICIO
# ---------------------------------------------------------
Atualizar-EstadoInterface
[void]$form.ShowDialog()
