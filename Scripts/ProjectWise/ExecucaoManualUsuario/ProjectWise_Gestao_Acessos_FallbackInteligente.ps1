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
$script:ResumoImportacaoUsuarios = $null

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

    $cmd = Get-Command $NomeCmdlet -ErrorAction Stop
    return @($cmd.Parameters.Keys)
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

    $gruposPW = Obter-TodosGruposPWCache
    if (-not $gruposPW -or $gruposPW.Count -eq 0) {
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

    $resultado = @()
    foreach ($grupo in $gruposPW) {
        $nomeGrupo = Obter-NomeGrupoPW -Grupo $grupo
        if ([string]::IsNullOrWhiteSpace($nomeGrupo)) { continue }

        $nomeGrupoNorm = Normalizar-TextoBuscaGrupo $nomeGrupo
        if ($nomesExistentes.ContainsKey($nomeGrupoNorm)) { continue }

        $nomeGrupoCompacto = Normalizar-TextoCompactoBuscaGrupo $nomeGrupoNorm
        $match = $false

        foreach ($apelido in $apelidos) {
            $apelidoNorm = Normalizar-TextoBuscaGrupo $apelido
            $apelidoCompacto = Normalizar-TextoCompactoBuscaGrupo $apelidoNorm

            if ($apelidoCompacto.Length -lt 6) { continue }

            if ($nomeGrupoNorm -like "*$apelidoNorm*" -or $nomeGrupoCompacto -like "*$apelidoCompacto*") {
                $match = $true
                break
            }
        }

        if ($match) {
            $resultado += [PSCustomObject]@{
                Tipo     = "Group"
                Nome     = $nomeGrupo
                ObjetoPw = $grupo
                Origem   = "InferidoPorNome"
            }
            $nomesExistentes[$nomeGrupoNorm] = $true
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
        [string]$Descricao = ""
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
        $nome = Obter-NomeAmigavelItem -Item $item
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
function Buscar-UsuarioPW {
    param([string]$Email)

    Write-Log "Buscando usuario pelo email: $Email"

    $parametros = Obter-ParametrosCmdlet -NomeCmdlet "Get-PWUsersByMatch"
    Write-Log ("Parametros encontrados no cmdlet: " + ($parametros -join ", "))

    $usuarios = $null

    if ($parametros -contains "Email") {
        $usuarios = @(Get-PWUsersByMatch -Email $Email -ErrorAction SilentlyContinue)
    }
    elseif ($parametros -contains "Identity") {
        $usuarios = @(Get-PWUsersByMatch -Identity $Email -ErrorAction SilentlyContinue)
    }
    elseif ($parametros -contains "UserName") {
        $usuarios = @(Get-PWUsersByMatch -UserName $Email -ErrorAction SilentlyContinue)
    }
    elseif ($parametros -contains "UserId") {
        $usuarios = @(Get-PWUsersByMatch -UserId $Email -ErrorAction SilentlyContinue)
    }
    else {
        $usuarios = @(Get-PWUsersByMatch $Email -ErrorAction SilentlyContinue)
    }

    $usuarios = @($usuarios | Where-Object { $_ -ne $null })

    if (-not $usuarios -or $usuarios.Count -eq 0) {
        throw "Nenhum usuario encontrado para o e-mail informado."
    }

    $usuariosComEmailExato = @(
        $usuarios | Where-Object {
            $emailItem = Obter-ValorSeguroPropriedade -Objeto $_ -PossiveisNomes @("Email","EMail","Mail")
            $emailItem -and $emailItem.ToLower() -eq $Email.ToLower()
        }
    )

    if ($usuariosComEmailExato.Count -eq 1) {
        Write-Log "Usuario localizado com correspondencia exata de email."
        return $usuariosComEmailExato[0]
    }

    if ($usuarios.Count -eq 1) {
        return $usuarios[0]
    }

    throw "Mais de um usuario foi encontrado para este criterio."
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
        [hashtable]$MapaMarcados = $null
    )

    if ($null -eq $MapaMarcados) {
        $MapaMarcados = Obter-MapaAcessosMarcados
    }

    $filtroNormalizado = Normalizar-Texto $Filtro

    $script:AtualizandoCheckTree = $true
    try {
        Limpar-ArvoreProjetos

        foreach ($projetoData in $ProjetosData) {
            $nomeProjeto = $projetoData.ProjetoNome
            $projetoCombina = $false

            if ([string]::IsNullOrWhiteSpace($filtroNormalizado)) {
                $projetoCombina = $true
            }
            else {
                $projetoCombina = (Normalizar-Texto $nomeProjeto) -like "*$filtroNormalizado*"
            }

            if ($projetoCombina) {
                $gruposFiltrados = @($projetoData.Groups)
                $listasFiltradas = @($projetoData.UserLists)
                $outrosFiltrados = @($projetoData.OutrosAcessos)
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

        if ([string]::IsNullOrWhiteSpace($Filtro)) {
            $script:LabelResultadoFiltro.Text = "Projetos exibidos: $qtdProjetosVisiveis"
        }
        else {
            $script:LabelResultadoFiltro.Text = "Filtro ativo: $qtdProjetosVisiveis de $qtdProjetosOriginais projeto(s) exibidos"
        }
    }
}

function Aplicar-FiltroProjetos {
    $filtro = ""
    if ($script:TextFiltroAcessos) {
        $filtro = $script:TextFiltroAcessos.Text
    }

    Renderizar-ArvoreProjetos -ProjetosData $script:ProjetosCarregados -Filtro $filtro
    Atualizar-EstadoInterface
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

    return @(
        $Texto -split '[,;|\r\n]+' |
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
        throw "A planilha selecionada nao possui usuarios."
    }

    $emails = New-Object System.Collections.Generic.List[string]
    $emailsUnicos = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $duplicados = New-Object System.Collections.Generic.List[string]
    $linhasSemEmail = New-Object System.Collections.Generic.List[int]
    $totalEmailsEncontrados = 0
    $numeroLinha = 2

    foreach ($linha in $dados) {
        $valor = Obter-ValorCampoLinha -Linha $linha -Nomes @("E-mail", "Email", "Mail", "Usuario", "Usuário", "User")
        if ([string]::IsNullOrWhiteSpace($valor)) {
            $linhasSemEmail.Add($numeroLinha)
            $numeroLinha++
            continue
        }

        foreach ($email in @(Obter-EmailsInformados -Texto $valor)) {
            $totalEmailsEncontrados++
            if ($emailsUnicos.Add($email)) {
                $emails.Add($email)
            }
            else {
                $duplicados.Add($email)
            }
        }

        $numeroLinha++
    }

    if ($emails.Count -eq 0) {
        throw "Nenhum e-mail foi encontrado. Use uma coluna chamada E-mail, Email, Mail ou Usuario."
    }

    $script:ResumoImportacaoUsuarios = [PSCustomObject]@{
        LinhasPlanilha          = $dados.Count
        TotalEmailsEncontrados  = $totalEmailsEncontrados
        EmailsUnicos            = $emails.Count
        DuplicadosIgnorados     = $duplicados.Count
        EmailsDuplicados        = @($duplicados | Select-Object -Unique)
        LinhasSemEmail          = @($linhasSemEmail.ToArray())
    }

    return $emails.ToArray()
}

function Confirmar-E-AdicionarUsuariosPorEmail {
    param(
        [string[]]$Emails,
        [switch]$SubstituirQuandoUnico
    )

    $adicionados = 0
    foreach ($emailInformado in $Emails) {
        $email = Validar-Email -Email $emailInformado
        Write-UiLog "Buscando usuario: $email"
        $usuario = Buscar-UsuarioPW -Email $email

        $null = Validar-IdentificadoresUsuario -Usuario $usuario

        $nome = Obter-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("Name","UserName","FullName")
        $emailEncontrado = Obter-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("Email","EMail","Mail")
        $login = Obter-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("LoginName","Login","UserID")

        $confirmar = Ask-UiYesNo -Mensagem "Confirmar usuário encontrado?`r`n`r`nNome : $nome`r`nEmail: $emailEncontrado`r`nLogin: $login" -Titulo "Confirmar usuário"

        if (-not $confirmar) {
            Write-UiLog "Usuario nao confirmado: $email"
            continue
        }

        $adicionado = Adicionar-UsuarioSelecionado -Usuario $usuario -Substituir:$SubstituirQuandoUnico
        if ($adicionado) {
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
    $script:BtnConfirmarSelecao.Enabled = $temConcessao -and $temArvore
    $script:BtnResumo.Enabled = $temSelecoes
    $script:BtnExecutar.Enabled = $temSelecoes
    $script:BtnLimparSelecoes.Enabled = $temSelecoes
    $script:TreeProjetos.Enabled = $temConcessao
    $script:TextFiltroAcessos.Enabled = $temConcessao
    $script:BtnAplicarFiltro.Enabled = $temConcessao
    $script:BtnLimparFiltro.Enabled = $temConcessao
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

function Carregar-ArvoreProjetosDaConcessao {
    param([object]$Concessao)

    Limpar-ArvoreProjetos
    $script:ProjetosCarregados = @()

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

    foreach ($projeto in $projetos) {
        $nomeProjeto = Obter-RotuloProjeto -Projeto $projeto
        $nomeProjetoTecnico = Obter-ValorSeguroPropriedade -Objeto $projeto -PossiveisNomes @("Name","FolderName","ProjectName")
        $idProjeto = Obter-IdPasta -Pasta $projeto

        Write-UiLog "Lendo acessos do projeto '$nomeProjeto'..."
        $seguranca = Ler-SegurancaProjeto -PastaProjeto $projeto
        $acessos = Filtrar-AcessosProjeto -ObjetosSeguranca $seguranca["TodosAcessos"]

        $groups = @()
        $userLists = @()
        $outrosAcessos = @()

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

            if ($tipo -eq 'Group') {
                $groups += $itemAcesso
            }
            elseif ($tipo -eq 'UserList') {
                $userLists += $itemAcesso
            }
            else {
                $outrosAcessos += $itemAcesso
            }
        }

        # FALLBACK INTELIGENTE:
        # Se a API não retornou groups suficientes na segurança da pasta/projeto,
        # tenta localizar grupos existentes pelo padrão do nome do projeto.
        # Isso cobre cenários em que o ProjectWise Administrator mostra groups no nível de projeto,
        # mas o módulo pwps_dab não expõe um cmdlet de Project Security no ambiente.
        $usarFallbackGroups = ($groups.Count -lt 3)
        if ($usarFallbackGroups) {
            Write-Log "Projeto '$nomeProjeto' retornou apenas $($groups.Count) group(s) pela API. Tentando fallback por nome..." "WARN"
            $gruposInferidos = Obter-GruposInferidosPorProjeto `
                -ProjetoNome $nomeProjeto `
                -ProjetoNomeTecnico $nomeProjetoTecnico `
                -GroupsAtuais $groups

            if ($gruposInferidos.Count -gt 0) {
                $groups += @($gruposInferidos)
            }
        }

        if (($groups.Count + $userLists.Count + $outrosAcessos.Count) -eq 0) {
            Write-Log "Projeto '$nomeProjeto' ignorado: nenhum acesso com nome identificado." "WARN"
            continue
        }

        if (($groups.Count + $userLists.Count) -eq 0 -and $outrosAcessos.Count -gt 0) {
            Write-Log "Projeto '$nomeProjeto' possui apenas outros acessos. Sera exibido somente para diagnostico visual." "WARN"
        }

        Write-Log "Projeto '$nomeProjeto' carregado para arvore. Groups=$($groups.Count) | UserLists=$($userLists.Count) | Outros=$($outrosAcessos.Count)"

        $script:ProjetosCarregados += [PSCustomObject]@{
            ConcessaoNome     = $nomeConcessao
            ConcessaoObjeto   = $Concessao
            ProjetoNome       = $nomeProjeto
            ProjetoNomeTecnico = $nomeProjetoTecnico
            ProjetoObjeto     = $projeto
            ProjetoId         = $idProjeto
            Groups            = @($groups | Sort-Object Nome)
            UserLists         = @($userLists | Sort-Object Nome)
            OutrosAcessos     = @($outrosAcessos | Sort-Object Tipo, Nome)
        }
    }

    $script:ConcessaoSelecionada = $Concessao
    $script:LabelConcessaoAtual.Text = "Concessão: $nomeConcessao"
    if ($script:TextFiltroAcessos) {
        $script:TextFiltroAcessos.Text = ""
    }
    Renderizar-ArvoreProjetos -ProjetosData $script:ProjetosCarregados -Filtro ""
    Write-UiLog "Concessao '$nomeConcessao' carregada com sucesso."
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
$script:BtnAplicarFiltro.Location = New-Object System.Drawing.Point(915,95)
$script:BtnAplicarFiltro.Size = New-Object System.Drawing.Size(90,30)
$script:BtnAplicarFiltro.Enabled = $false
$groupSelecao.Controls.Add($script:BtnAplicarFiltro)

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
        $script:BtnAplicarFiltro.Left = $script:BtnLimparFiltro.Left - $script:BtnAplicarFiltro.Width - $espacamento

        $filtroLeft = $script:TextFiltroAcessos.Left
        if ($labelFiltroAcessos) {
            $filtroLeft = $labelFiltroAcessos.Left + $labelFiltroAcessos.Width + 8
            $script:TextFiltroAcessos.Left = $filtroLeft
        }

        $script:TextFiltroAcessos.Width = [Math]::Max(260, $script:BtnAplicarFiltro.Left - $filtroLeft - $espacamento)
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
$script:BtnConfirmarSelecao.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$script:LabelConcessaoAtual.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TextFiltroAcessos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:TextFiltroAcessos.BringToFront()

$script:BtnAplicarFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$script:BtnLimparFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
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
        Aplicar-FiltroProjetos
    }
    catch {
        Write-Log "Erro ao aplicar filtro. Detalhe: $($_.Exception.Message)" "ERROR"
        Show-UiError $_.Exception.Message
    }
})

$script:BtnLimparFiltro.Add_Click({
    try {
        $script:TextFiltroAcessos.Text = ""
        Aplicar-FiltroProjetos
    }
    catch {
        Write-Log "Erro ao limpar filtro. Detalhe: $($_.Exception.Message)" "ERROR"
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
        $adicionados = Confirmar-E-AdicionarUsuariosPorEmail -Emails $emails

        if (@($script:UsuariosSelecionados).Count -gt 0 -and $usuariosAntes -eq 0) {
            $script:ConcessaoSelecionada = $null
            $script:SelecoesProjetos = @()
            $script:ProjetosCarregados = @()
            Limpar-ArvoreProjetos
            if ($script:TextFiltroAcessos) { $script:TextFiltroAcessos.Text = "" }
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
            if ($script:TextFiltroAcessos) { $script:TextFiltroAcessos.Text = "" }
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
        Carregar-ArvoreProjetosDaConcessao -Concessao $concessaoSelecionada
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
