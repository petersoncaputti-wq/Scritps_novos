<#
.SYNOPSIS
    Cria usuarios do ProjectWise em lote a partir de CSV ou Excel.

.DESCRIPTION
    Colunas aceitas na planilha:
    - Name ou UserName                          Obrigatorio
    - Description                               Opcional
    - E-mail ou Email                           Obrigatorio

    O usuario sera criado sempre como Type = Bentley IMS.
    O campo Bentley IMS user name sera preenchido com o mesmo valor de E-mail.

.EXAMPLE
    .\ProjectWise_Criar_Usuarios_Lote.ps1 -GerarModelo

.EXAMPLE
    .\ProjectWise_Criar_Usuarios_Lote.ps1 -CaminhoArquivo .\usuarios_pw.csv -SomenteValidar

.EXAMPLE
    $senha = Read-Host "Senha PW" -AsSecureString
    .\ProjectWise_Criar_Usuarios_Lote.ps1 -CaminhoArquivo .\usuarios_pw.xlsx -DatasourceName "Servidor:Datasource" -UserName "admin" -Password $senha
#>

[CmdletBinding()]
param(
    [string]$CaminhoArquivo,
    [string]$DatasourceName,
    [string]$UserName,
    [securestring]$Password,
    [switch]$BentleyIMS,
    [switch]$AtualizarExistentes,
    [switch]$SomenteValidar,
    [switch]$NaoAdicionarAcessos,
    [switch]$GerarModelo,
    [switch]$NaoPausar,
    [string]$CaminhoModelo
)

Add-Type -AssemblyName System.Windows.Forms

$ErrorActionPreference = 'Stop'

$PastaLog = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path $PastaLog)) {
    New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArquivoLog = Join-Path $PastaLog "PW_CriarUsuariosLote_$TimeStamp.log"
$ArquivoRelatorio = Join-Path $PastaLog "PW_CriarUsuariosLote_Relatorio_$TimeStamp.csv"

function Write-Log {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    $linha = "{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Nivel.ToUpper(), $Mensagem
    [System.IO.File]::AppendAllText($ArquivoLog, $linha + [Environment]::NewLine)
}

function Write-Status {
    param(
        [string]$Mensagem,
        [string]$Nivel = "INFO"
    )

    Write-Log -Mensagem $Mensagem -Nivel $Nivel

    switch ($Nivel.ToUpper()) {
        "ERROR" { Write-Host $Mensagem -ForegroundColor Red }
        "WARN"  { Write-Host $Mensagem -ForegroundColor Yellow }
        "OK"    { Write-Host $Mensagem -ForegroundColor Green }
        default { Write-Host $Mensagem }
    }
}

function Pausar-Final {
    param([string]$Mensagem = "Processo finalizado. Pressione ENTER para fechar esta janela.")

    if ($NaoPausar) {
        return
    }

    try {
        Write-Host ""
        Read-Host $Mensagem | Out-Null
    }
    catch {
    }
}

function Require-Command {
    param([string[]]$Nomes)

    foreach ($nome in $Nomes) {
        if (-not (Get-Command $nome -ErrorAction SilentlyContinue)) {
            throw "Cmdlet obrigatorio nao encontrado: $nome. Verifique se o modulo PWPS_DAB esta instalado/carregado."
        }
    }
}

function Get-ValorCampo {
    param(
        [object]$Linha,
        [string[]]$Nomes
    )

    foreach ($nome in $Nomes) {
        $prop = @($Linha.PSObject.Properties | Where-Object { $_.Name -ieq $nome } | Select-Object -First 1)
        if ($prop.Count -gt 0 -and $null -ne $prop[0].Value) {
            $valor = ([string]$prop[0].Value).Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
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
        $prop = $Objeto.PSObject.Properties[$nome]
        if ($prop -and $null -ne $prop.Value) {
            $valor = ([string]$prop.Value).Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
}

function ConvertTo-BoolSeguro {
    param([string]$Valor)

    if ([string]::IsNullOrWhiteSpace($Valor)) {
        return $false
    }

    $normalizado = $Valor.Trim().ToLowerInvariant()
    return @("1", "true", "sim", "s", "yes", "y", "ims", "federado") -contains $normalizado
}

function Split-Lista {
    param([string]$Valor)

    if ([string]::IsNullOrWhiteSpace($Valor)) {
        return @()
    }

    return @(
        $Valor -split '[;|,]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -ne "" } |
            Select-Object -Unique
    )
}

function Test-Email {
    param([string]$Email)

    if ([string]::IsNullOrWhiteSpace($Email)) {
        return $true
    }

    return ($Email.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Selecionar-ArquivoEntrada {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
    $dialog.Title = "Selecione a planilha de usuarios do ProjectWise"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    throw "Nenhum arquivo foi selecionado."
}

function Importar-DadosUsuarios {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Arquivo nao encontrado: $Path"
    }

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in @(".xlsx", ".xlsm")) {
        Import-Module ImportExcel -ErrorAction Stop
        $dados = @(Import-Excel -Path $Path)
        $linhaInicial = 2
    }
    elseif ($ext -eq ".csv") {
        $primeiraLinha = Get-Content -Path $Path -TotalCount 1
        $qtdPontoVirgula = ([regex]::Matches($primeiraLinha, ';')).Count
        $qtdVirgula = ([regex]::Matches($primeiraLinha, ',')).Count
        $delimitador = ','
        if ($qtdPontoVirgula -gt $qtdVirgula) {
            $delimitador = ';'
        }

        $dados = @(Import-Csv -Path $Path -Delimiter $delimitador)
        $linhaInicial = 2
    }
    else {
        throw "Extensao nao suportada: $ext. Use .csv, .xlsx ou .xlsm."
    }

    $indice = 0
    foreach ($linha in $dados) {
        $linha | Add-Member -NotePropertyName "__LinhaOriginal" -NotePropertyValue ($linhaInicial + $indice) -Force
        $indice++
    }

    return $dados
}

function New-RegistroUsuario {
    param([object]$Linha)

    $userName = Get-ValorCampo -Linha $Linha -Nomes @("Name", "UserName", "Usuario", "UsuarioPW", "Login", "NomeUsuario")
    $descricao = Get-ValorCampo -Linha $Linha -Nomes @("Description", "Descrition", "Descricao", "NomeCompleto", "Nome Completo")
    $email = Get-ValorCampo -Linha $Linha -Nomes @("Email", "E-mail", "Mail")
    $securityProvider = Get-ValorCampo -Linha $Linha -Nomes @("SecurityProvider", "ProvedorSeguranca", "Provedor", "Dominio", "Domain")
    $senhaTexto = Get-ValorCampo -Linha $Linha -Nomes @("Password", "Senha")
    $imsUser = $true
    $grupos = Split-Lista (Get-ValorCampo -Linha $Linha -Nomes @("Groups", "Grupos", "Grupo"))
    $userLists = Split-Lista (Get-ValorCampo -Linha $Linha -Nomes @("UserLists", "ListasUsuario", "User Lists", "UserList", "ListaUsuario"))

    return [PSCustomObject]@{
        Linha            = $Linha.__LinhaOriginal
        UserName         = $userName
        Description      = $descricao
        Email            = $email
        SecurityProvider = $securityProvider
        PasswordText     = $senhaTexto
        IMSUser          = $imsUser
        Groups           = $grupos
        UserLists        = $userLists
    }
}

function Validar-RegistroUsuario {
    param([object]$Registro)

    $erros = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($Registro.UserName)) {
        $erros.Add("Name vazio")
    }

    if (-not (Test-Email -Email $Registro.Email)) {
        $erros.Add("Email em formato invalido")
    }

    if ([string]::IsNullOrWhiteSpace($Registro.Email)) {
        $erros.Add("E-mail vazio")
    }

    return @($erros)
}

function Conectar-ProjectWise {
    Require-Command -Nomes @("New-PWLogin", "Undo-PWLogin", "Get-PWCurrentUser")

    if ($SomenteValidar) {
        Write-Status "Modo SomenteValidar: login no ProjectWise nao sera executado."
        return $null
    }

    Write-Status "Abrindo login no ProjectWise..."

    if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
        $login = New-PWLogin -UseGui -ErrorAction Stop
    }
    else {
        $parametros = @{
            DatasourceName = $DatasourceName
            ErrorAction    = "Stop"
        }

        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            $parametros.UserName = $UserName
        }

        if ($Password) {
            $parametros.Password = $Password
        }

        if ($BentleyIMS) {
            $parametros.BentleyIMS = $true
        }

        $login = New-PWLogin @parametros
    }

    $usuarioAtual = Get-PWCurrentUser -ErrorAction SilentlyContinue
    $nomeAtual = Get-ValorSeguroPropriedade -Objeto $usuarioAtual -PossiveisNomes @("Name", "UserName", "LoginName")
    Write-Status "Login realizado. Usuario atual: $nomeAtual" "OK"

    return $login
}

function Buscar-UsuarioPW {
    param([object]$Registro)

    $resultado = Test-UsuarioExistePW -Registro $Registro
    if ($resultado.Existe) {
        return $resultado.Usuario
    }

    return $null
}

function Test-UsuarioExistePW {
    param([object]$Registro)

    $resultado = [PSCustomObject]@{
        Existe       = $false
        Usuario      = $null
        EncontradoPor = ""
    }

    if (-not [string]::IsNullOrWhiteSpace($Registro.UserName)) {
        $usuariosPorNome = @(Get-PWUsersByMatch -UserName $Registro.UserName -ErrorAction SilentlyContinue)
        $usuariosPorNome = @($usuariosPorNome | Where-Object {
            (Get-ValorSeguroPropriedade -Objeto $_ -PossiveisNomes @("Name", "UserName", "LoginName")) -ieq $Registro.UserName
        })

        if ($usuariosPorNome.Count -gt 0) {
            $resultado.Existe = $true
            $resultado.Usuario = $usuariosPorNome[0]
            $resultado.EncontradoPor = "Name"
            return $resultado
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Registro.Email)) {
        $usuariosPorEmail = @(Get-PWUsersByMatch -Email $Registro.Email -ErrorAction SilentlyContinue)
        $usuariosPorEmail = @($usuariosPorEmail | Where-Object {
            (Get-ValorSeguroPropriedade -Objeto $_ -PossiveisNomes @("Email", "EMail", "Mail")) -ieq $Registro.Email
        })

        if ($usuariosPorEmail.Count -gt 0) {
            $resultado.Existe = $true
            $resultado.Usuario = $usuariosPorEmail[0]
            $resultado.EncontradoPor = "E-mail"
            return $resultado
        }
    }

    return $resultado
}

function Criar-UsuarioPW {
    param([object]$Registro)

    $parametros = @{
        UserNames   = @($Registro.UserName)
        Email       = $Registro.Email
        IMSUser     = $true
        ErrorAction = "Stop"
    }

    if (-not [string]::IsNullOrWhiteSpace($Registro.Description)) {
        $parametros.Description = $Registro.Description
    }

    New-PWUserSimple @parametros | Out-Null
    $usuario = Buscar-UsuarioPW -Registro $Registro

    if ($usuario) {
        Set-IdentidadeBentleyIMS -Usuario $usuario -Email $Registro.Email
    }

    return $usuario
}

function Atualizar-UsuarioPW {
    param(
        [object]$Usuario,
        [object]$Registro
    )

    $parametros = @{
        InputUser   = $Usuario
        ErrorAction = "Stop"
    }

    if (-not [string]::IsNullOrWhiteSpace($Registro.Description)) {
        $parametros.Description = $Registro.Description
    }

    if (-not [string]::IsNullOrWhiteSpace($Registro.Email)) {
        $parametros.Email = $Registro.Email
    }

    if ($parametros.Keys.Count -gt 2) {
        Update-PWUserProperties @parametros | Out-Null
    }

    Set-IdentidadeBentleyIMS -Usuario $Usuario -Email $Registro.Email
}

function Set-IdentidadeBentleyIMS {
    param(
        [object]$Usuario,
        [string]$Email
    )

    if ([string]::IsNullOrWhiteSpace($Email)) {
        throw "Nao foi possivel definir Bentley IMS user name porque o e-mail esta vazio."
    }

    Set-PWUserIdentity -Users @($Usuario) -Identity $Email -ErrorAction Stop | Out-Null
}

function Test-UsuarioEmColecao {
    param(
        [object]$Usuario,
        [object[]]$Membros
    )

    $userId = Get-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("UserID", "UserId", "ID", "Id")
    $userName = Get-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Name", "UserName", "LoginName")
    $email = Get-ValorSeguroPropriedade -Objeto $Usuario -PossiveisNomes @("Email", "EMail", "Mail")

    foreach ($membro in $Membros) {
        $membroId = Get-ValorSeguroPropriedade -Objeto $membro -PossiveisNomes @("UserID", "UserId", "ID", "Id")
        $membroNome = Get-ValorSeguroPropriedade -Objeto $membro -PossiveisNomes @("Name", "UserName", "LoginName")
        $membroEmail = Get-ValorSeguroPropriedade -Objeto $membro -PossiveisNomes @("Email", "EMail", "Mail")

        if ($userId -ne "" -and $membroId -ne "" -and $userId -eq $membroId) {
            return $true
        }

        if ($userName -ne "" -and $membroNome -ne "" -and $userName -ieq $membroNome) {
            return $true
        }

        if ($email -ne "" -and $membroEmail -ne "" -and $email -ieq $membroEmail) {
            return $true
        }
    }

    return $false
}

function Add-UsuarioEmAcessos {
    param(
        [object]$Usuario,
        [object]$Registro
    )

    $resultados = New-Object System.Collections.Generic.List[string]

    if ($NaoAdicionarAcessos) {
        return "Ignorado por parametro NaoAdicionarAcessos"
    }

    foreach ($grupo in $Registro.Groups) {
        try {
            $membros = @(Get-PWUsersInGroup -GroupName $grupo -ErrorAction Stop)
            if (Test-UsuarioEmColecao -Usuario $Usuario -Membros $membros) {
                $resultados.Add("Grupo '$grupo': ja existe")
                continue
            }

            Add-PWUserToGroup -InputUser @($Usuario) -GroupName $grupo -ErrorAction Stop | Out-Null
            $resultados.Add("Grupo '$grupo': adicionado")
        }
        catch {
            $resultados.Add("Grupo '$grupo': erro - $($_.Exception.Message)")
        }
    }

    foreach ($userList in $Registro.UserLists) {
        try {
            $membros = @(Get-PWUsersInUserList -UserList $userList -ErrorAction Stop)
            if (Test-UsuarioEmColecao -Usuario $Usuario -Membros $membros) {
                $resultados.Add("UserList '$userList': ja existe")
                continue
            }

            Add-PWUserToUserList -InputUser @($Usuario) -UserList $userList -ErrorAction Stop | Out-Null
            $resultados.Add("UserList '$userList': adicionado")
        }
        catch {
            $resultados.Add("UserList '$userList': erro - $($_.Exception.Message)")
        }
    }

    if ($resultados.Count -eq 0) {
        return "Sem grupos/user lists"
    }

    return ($resultados -join " | ")
}

function New-Resultado {
    param(
        [object]$Registro,
        [string]$Status,
        [string]$Acao,
        [string]$Mensagem,
        [string]$Acessos = ""
    )

    return [PSCustomObject]@{
        Linha            = $Registro.Linha
        UserName         = $Registro.UserName
        Email            = $Registro.Email
        SecurityProvider = $Registro.SecurityProvider
        IMSUser          = $Registro.IMSUser
        Status           = $Status
        Acao             = $Acao
        Mensagem         = $Mensagem
        Acessos          = $Acessos
    }
}

function Exportar-Modelo {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Join-Path $PSScriptRoot "Modelo_Criacao_Usuarios_ProjectWise.xlsx"
    }

    $modelo = @(
        [PSCustomObject]@{
            Name        = "nome.sobrenome"
            Description = "Nome Sobrenome"
            "E-mail"    = "nome.sobrenome@empresa.com"
        }
    )

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in @(".xlsx", ".xlsm")) {
        Import-Module ImportExcel -ErrorAction Stop
        $modelo | Export-Excel -Path $Path -WorksheetName "Usuarios" -AutoSize -ClearSheet
    }
    else {
        $modelo | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    }

    Write-Status "Modelo gerado em: $Path" "OK"
}

function Processar-Registros {
    param([object[]]$Registros)

    Require-Command -Nomes @(
        "Get-PWUsersByMatch",
        "New-PWUserSimple",
        "Update-PWUserProperties",
        "Set-PWUserIdentity"
    )

    $precisaAcessos = @($Registros | Where-Object { @($_.Groups).Count -gt 0 -or @($_.UserLists).Count -gt 0 }).Count -gt 0
    if ($precisaAcessos) {
        Require-Command -Nomes @(
            "Add-PWUserToGroup",
            "Add-PWUserToUserList",
            "Get-PWUsersInGroup",
            "Get-PWUsersInUserList"
        )
    }

    $resultados = New-Object System.Collections.Generic.List[object]

    foreach ($registro in $Registros) {
        try {
            $erros = @(Validar-RegistroUsuario -Registro $registro)
            if ($erros.Count -gt 0) {
                $mensagem = $erros -join "; "
                Write-Status "Linha $($registro.Linha): invalida - $mensagem" "WARN"
                $resultados.Add((New-Resultado -Registro $registro -Status "INVALIDO" -Acao "Validacao" -Mensagem $mensagem))
                continue
            }

            if ($SomenteValidar) {
                Write-Status "Linha $($registro.Linha): valida - $($registro.UserName)" "OK"
                $resultados.Add((New-Resultado -Registro $registro -Status "VALIDO" -Acao "SomenteValidar" -Mensagem "Registro validado. Nenhuma alteracao feita."))
                continue
            }

            $verificacaoUsuario = Test-UsuarioExistePW -Registro $registro
            $usuario = $verificacaoUsuario.Usuario
            $acao = ""

            if ($verificacaoUsuario.Existe) {
                if ($AtualizarExistentes) {
                    Atualizar-UsuarioPW -Usuario $usuario -Registro $registro
                    $usuario = Buscar-UsuarioPW -Registro $registro
                    $acao = "Atualizado"
                    Write-Status "Linha $($registro.Linha): usuario existente atualizado - $($registro.UserName) (encontrado por $($verificacaoUsuario.EncontradoPor))" "OK"
                }
                else {
                    $acao = "JaExistia"
                    Write-Status "Linha $($registro.Linha): usuario ja existe - $($registro.UserName) (encontrado por $($verificacaoUsuario.EncontradoPor))" "WARN"
                }
            }
            else {
                $usuario = Criar-UsuarioPW -Registro $registro
                if (-not $usuario) {
                    throw "O cmdlet de criacao executou, mas o usuario nao foi localizado em seguida."
                }

                $acao = "Criado"
                Write-Status "Linha $($registro.Linha): usuario criado - $($registro.UserName)" "OK"
            }

            $statusAcessos = Add-UsuarioEmAcessos -Usuario $usuario -Registro $registro
            $resultados.Add((New-Resultado -Registro $registro -Status "OK" -Acao $acao -Mensagem "Processado com sucesso." -Acessos $statusAcessos))
        }
        catch {
            $mensagemErro = $_.Exception.Message
            Write-Status "Linha $($registro.Linha): erro - $mensagemErro" "ERROR"
            $resultados.Add((New-Resultado -Registro $registro -Status "ERRO" -Acao "Processamento" -Mensagem $mensagemErro))
        }
    }

    return $resultados.ToArray()
}

function Mostrar-ResumoFinal {
    param(
        [object[]]$Resultados,
        [string]$ArquivoRelatorioFinal
    )

    $total = @($Resultados).Count
    $criados = @($Resultados | Where-Object { $_.Status -eq "OK" -and $_.Acao -eq "Criado" }).Count
    $atualizados = @($Resultados | Where-Object { $_.Status -eq "OK" -and $_.Acao -eq "Atualizado" }).Count
    $jaExistiam = @($Resultados | Where-Object { $_.Status -eq "OK" -and $_.Acao -eq "JaExistia" }).Count
    $validos = @($Resultados | Where-Object { $_.Status -eq "VALIDO" }).Count
    $invalidos = @($Resultados | Where-Object { $_.Status -eq "INVALIDO" }).Count
    $erros = @($Resultados | Where-Object { $_.Status -eq "ERRO" }).Count

    Write-Status ""
    Write-Status "================ RESUMO FINAL ================"
    Write-Status "Total de linhas processadas : $total"
    Write-Status "Usuarios criados            : $criados" "OK"
    Write-Status "Usuarios atualizados        : $atualizados" "OK"
    Write-Status "Usuarios ja existentes      : $jaExistiam" "WARN"
    Write-Status "Linhas apenas validadas     : $validos"
    Write-Status "Linhas invalidas            : $invalidos" $(if ($invalidos -gt 0) { "WARN" } else { "OK" })
    Write-Status "Erros de processamento      : $erros" $(if ($erros -gt 0) { "ERROR" } else { "OK" })
    Write-Status "Relatorio completo          : $ArquivoRelatorioFinal" "OK"

    $itensComProblema = @($Resultados | Where-Object { $_.Status -in @("INVALIDO", "ERRO") })
    if ($itensComProblema.Count -gt 0) {
        Write-Status ""
        Write-Status "Linhas que precisam de atencao:" "WARN"
        foreach ($item in $itensComProblema) {
            Write-Status ("- Linha {0} | {1} | {2} | {3}" -f $item.Linha, $item.Status, $item.UserName, $item.Mensagem) "WARN"
        }
    }

    $itensExistentes = @($Resultados | Where-Object { $_.Status -eq "OK" -and $_.Acao -eq "JaExistia" })
    if ($itensExistentes.Count -gt 0) {
        Write-Status ""
        Write-Status "Usuarios que ja existiam e nao foram recriados:" "WARN"
        foreach ($item in $itensExistentes) {
            Write-Status ("- Linha {0} | {1} | {2}" -f $item.Linha, $item.UserName, $item.Email) "WARN"
        }
    }

    Write-Status "=============================================="
}

try {
    Write-Status "ProjectWise - criacao de usuarios em lote"
    Write-Status "Log: $ArquivoLog"

    if ($GerarModelo) {
        Exportar-Modelo -Path $CaminhoModelo
        Pausar-Final
        return
    }

    Import-Module PWPS_DAB -ErrorAction Stop

    $login = Conectar-ProjectWise

    if ([string]::IsNullOrWhiteSpace($CaminhoArquivo)) {
        $CaminhoArquivo = Selecionar-ArquivoEntrada
    }

    Write-Status "Arquivo de entrada: $CaminhoArquivo"
    $linhas = @(Importar-DadosUsuarios -Path $CaminhoArquivo)
    if ($linhas.Count -eq 0) {
        throw "A planilha nao possui registros."
    }

    $registros = @($linhas | ForEach-Object { New-RegistroUsuario -Linha $_ })
    Write-Status "Registros carregados: $($registros.Count)"

    try {
        $resultados = @(Processar-Registros -Registros $registros)
    }
    finally {
        if ($login -and -not $SomenteValidar) {
            try {
                Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null
                Write-Status "Sessao ProjectWise encerrada."
            }
            catch {
                Write-Log "Falha ao encerrar sessao: $($_.Exception.Message)" "WARN"
            }
        }
    }

    $resultados | Export-Csv -Path $ArquivoRelatorio -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    Mostrar-ResumoFinal -Resultados $resultados -ArquivoRelatorioFinal $ArquivoRelatorio
    Pausar-Final
}
catch {
    Write-Status "Falha geral: $($_.Exception.Message)" "ERROR"
    Pausar-Final -Mensagem "O processo terminou com erro. Pressione ENTER para fechar esta janela."
    throw
}
