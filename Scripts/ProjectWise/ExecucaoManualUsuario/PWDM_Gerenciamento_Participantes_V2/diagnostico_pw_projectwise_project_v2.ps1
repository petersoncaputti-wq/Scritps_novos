<#
.SYNOPSIS
    Diagnostica propriedades "ProjectWise Project" de uma pasta/projeto ProjectWise.

.DESCRIPTION
    Script somente leitura. Conecta no ProjectWise, localiza um projeto por FolderPath
    ou FolderId e exporta todas as propriedades retornadas pelo modulo PWPS_DAB.

    O objetivo e descobrir quais campos correspondem aos dados da aba
    "ProjectWise Project" do Explorer, especialmente ID/Name usados pelo PW Web/PWDM.

.EXAMPLE
    .\diagnostico_pw_projectwise_project_v2.ps1 -FolderPath "Engenharia\Concessao\Projetos\Projeto X"

.EXAMPLE
    .\diagnostico_pw_projectwise_project_v2.ps1 -FolderId 12345
#>

[CmdletBinding(DefaultParameterSetName = "PorCaminho")]
param(
    [Parameter(ParameterSetName = "PorCaminho", Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(ParameterSetName = "PorId", Mandatory = $true)]
    [string]$FolderId,

    [string]$Saida,
    [switch]$NaoDesconectar
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if ([string]::IsNullOrWhiteSpace($Saida)) {
    $Saida = Join-Path $PastaLogs "pw_projectwise_project_diagnostico_$TimeStamp.json"
}

$ArquivoLog = Join-Path $PastaLogs "pw_projectwise_project_diagnostico_$TimeStamp.log"
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
    foreach ($modulo in @("pwps_dab", "PWPS_DAB", "Bentley.PowerShell.ProjectWise", "ProjectWisePowerShell")) {
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

    throw "Nao foi possivel localizar/importar o modulo do ProjectWise."
}

function Conectar-ProjectWise {
    foreach ($nomeCmdlet in @("New-PWLogin", "Get-PWLogin", "Open-PWConnection")) {
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
            Write-Log "Sessao encerrada via '$nomeCmdlet'."
            return
        }
        catch {
            Write-Log "Falha ao encerrar sessao via '$nomeCmdlet': $($_.Exception.Message)" "WARN"
        }
    }
}

function Converter-ValorSeguro {
    param(
        [object]$Valor,
        [int]$Profundidade = 0
    )

    if ($null -eq $Valor) {
        return $null
    }

    if ($Profundidade -ge 3) {
        return $Valor.ToString()
    }

    if ($Valor -is [string] -or $Valor.GetType().IsPrimitive -or $Valor -is [decimal]) {
        return $Valor
    }

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
            if ($contador -ge 30) {
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

    if ($objeto.Count -eq 0) {
        return $Valor.ToString()
    }

    return $objeto
}

function Obter-ValorPropriedade {
    param(
        [object]$Objeto,
        [string[]]$Nomes
    )

    foreach ($nome in $Nomes) {
        $prop = $Objeto.PSObject.Properties[$nome]
        if ($prop -and $null -ne $prop.Value -and $prop.Value.ToString().Trim() -ne "") {
            return $prop.Value.ToString().Trim()
        }
    }

    return ""
}

function Obter-ProjetoAlvo {
    if ($PSCmdlet.ParameterSetName -eq "PorCaminho") {
        Write-Log "Buscando pasta por caminho: $FolderPath"
        $cmd = Get-Command Get-PWFolders -ErrorAction Stop
        $params = @($cmd.Parameters.Keys)
        if ($params -contains "JustOne" -and $params -contains "Slow") {
            return Get-PWFolders -FolderPath $FolderPath -JustOne -Slow -ErrorAction Stop
        }
        if ($params -contains "JustOne") {
            return Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
        }
        return @(Get-PWFolders -FolderPath $FolderPath -ErrorAction Stop | Select-Object -First 1)[0]
    }

    Write-Log "Buscando pasta por FolderID: $FolderId"
    $cmdChildren = Get-Command Get-PWFoldersImmediateChildren -ErrorAction SilentlyContinue
    $cmdFolders = Get-Command Get-PWFolders -ErrorAction SilentlyContinue

    if ($cmdFolders) {
        $params = @($cmdFolders.Parameters.Keys)
        foreach ($nomeParam in @("FolderID", "FolderId", "ProjectID", "ProjectId")) {
            if ($params -contains $nomeParam) {
                $parametros = @{ $nomeParam = $FolderId; ErrorAction = "Stop" }
                return @(& $cmdFolders @parametros | Select-Object -First 1)[0]
            }
        }
    }

    if ($cmdChildren) {
        throw "Get-PWFolders nao aceitou busca por ID neste ambiente. Informe -FolderPath para o teste."
    }

    throw "Nao encontrei cmdlet adequado para buscar projeto."
}

function Obter-CmdletsRelacionados {
    $padroes = @("*PW*Project*", "*PW*Folder*Propert*", "*PW*Project*Propert*", "*PW*WorkArea*")
    $cmdlets = @()
    foreach ($padrao in $padroes) {
        $cmdlets += @(Get-Command -Name $padrao -ErrorAction SilentlyContinue)
    }

    return @(
        $cmdlets |
            Sort-Object Name -Unique |
            ForEach-Object {
                [ordered]@{
                    name = $_.Name
                    module = $_.ModuleName
                    parameters = @($_.Parameters.Keys)
                }
            }
    )
}

function Testar-CmdletsPropriedades {
    param([object]$Projeto)

    $resultados = @()
    $idProjeto = Obter-ValorPropriedade -Objeto $Projeto -Nomes @("ProjectID", "ProjectId", "FolderID", "FolderId", "Id", "ID")

    foreach ($nomeCmdlet in @(
        "Get-PWProjectProperties",
        "Get-PWProjectProperty",
        "Get-PWFolderProperties",
        "Get-PWFolderProperty",
        "Get-PWProjectDetails",
        "Get-PWWorkAreaProperties"
    )) {
        $cmd = Get-Command $nomeCmdlet -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        $params = @($cmd.Parameters.Keys)
        $tentativas = @()
        if ($params -contains "InputFolder") {
            $tentativas += @{ InputFolder = $Projeto; ErrorAction = "Stop" }
        }
        if ($params -contains "InputProject") {
            $tentativas += @{ InputProject = $Projeto; ErrorAction = "Stop" }
        }
        if ($idProjeto -and $params -contains "FolderID") {
            $tentativas += @{ FolderID = $idProjeto; ErrorAction = "Stop" }
        }
        if ($idProjeto -and $params -contains "ProjectID") {
            $tentativas += @{ ProjectID = $idProjeto; ErrorAction = "Stop" }
        }

        foreach ($tentativa in $tentativas) {
            try {
                $valor = & $nomeCmdlet @tentativa
                $resultados += [ordered]@{
                    cmdlet = $nomeCmdlet
                    parametrosUsados = @($tentativa.Keys)
                    sucesso = $true
                    resultado = Converter-ValorSeguro -Valor $valor
                }
                break
            }
            catch {
                $resultados += [ordered]@{
                    cmdlet = $nomeCmdlet
                    parametrosUsados = @($tentativa.Keys)
                    sucesso = $false
                    erro = $_.Exception.Message
                }
            }
        }
    }

    return $resultados
}

function Encontrar-Guids {
    param([object]$Valor)

    $texto = ($Valor | ConvertTo-Json -Depth 20 -Compress)
    $regex = [regex]"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
    return @($regex.Matches($texto) | ForEach-Object { $_.Value.ToLowerInvariant() } | Select-Object -Unique)
}

try {
    Write-Log "Iniciando diagnostico ProjectWise Project."
    Importar-ModuloProjectWise
    Conectar-ProjectWise

    $projeto = Obter-ProjetoAlvo
    if ($null -eq $projeto) {
        throw "Projeto nao encontrado."
    }

    $propriedades = Converter-ValorSeguro -Valor $projeto
    $cmdletsRelacionados = Obter-CmdletsRelacionados
    $resultadosCmdlets = Testar-CmdletsPropriedades -Projeto $projeto
    $guids = Encontrar-Guids -Valor @{
        propriedades = $propriedades
        resultadosCmdlets = $resultadosCmdlets
    }

    $urlsPossiveis = @(
        $guids | ForEach-Object {
            "https://pwdm.bentley.com/$_/ProjectSettings/$_/View#PARTICIPANTS"
        }
    )

    $dados = [ordered]@{
        timestamp = (Get-Date).ToString("s")
        entrada = [ordered]@{
            parameterSet = $PSCmdlet.ParameterSetName
            folderPath = $FolderPath
            folderId = $FolderId
        }
        propriedadesProjeto = $propriedades
        guidsEncontrados = $guids
        urlsPwdmPossiveis = $urlsPossiveis
        cmdletsRelacionados = $cmdletsRelacionados
        resultadosCmdletsPropriedades = $resultadosCmdlets
    }

    $json = $dados | ConvertTo-Json -Depth 30
    [System.IO.File]::WriteAllText($Saida, $json, [System.Text.Encoding]::UTF8)

    Write-Host ""
    Write-Host "[OK] Diagnostico salvo."
    Write-Host "JSON: $Saida"
    Write-Host "Log : $ArquivoLog"
    if ($urlsPossiveis.Count -gt 0) {
        Write-Host ""
        Write-Host "URLs PWDM possiveis:"
        $urlsPossiveis | ForEach-Object { Write-Host "- $_" }
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    Encerrar-SessaoProjectWise
}
