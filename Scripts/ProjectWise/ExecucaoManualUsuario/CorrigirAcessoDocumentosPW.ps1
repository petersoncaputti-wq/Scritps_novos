param(
    [string]$NomeGrupo = "PW-CORRECAO-ACESSO-DOCUMENTOS",
    [string]$NomeConcessao,
    [string[]]$Projetos,
    [switch]$AplicarWorkflowStates
)

# Script corretivo para aplicar acesso somente leitura/download em pastas e
# documentos de projetos existentes no ProjectWise.
# Nao remove permissoes existentes, nao altera workflow/state e nao concede
# permissao de edicao, criacao, exclusao ou escrita.

Import-Module PWPS_DAB -ErrorAction Stop

$folderAccess = @("r")
$documentAccess = @("r", "fr")

$pastasProjeto = @(
    "",
    "0 - Documentos Gerais",
    "1 - Area de Trabalho",
    "2 - Unidade",
    "Previsao de Documentos (LD)",
    "Area de Transferencia"
)

$workflowStatesConhecidos = @(
    @{
        Workflow = "Workflow - Engenharia"
        States = @(
            "Em elaboracao",
            "Em analise do Assistente",
            "Em analise da Engenharia",
            "Em analise Especifica",
            "Em consolidacao pela Engenharia",
            "Em Envio para Projetista",
            "Em revisao pelo Projetista",
            "Em Envio para Unidade",
            "Concluido",
            "Superado"
        )
    },
    @{
        Workflow = "Workflow - Engenharia - Unidade"
        States = @(
            "Em analise pela Unidade",
            "Em analise pela Engenharia",
            "Enviado ao Poder Concedente",
            "Concluido pela Unidade",
            "Nova emissao sendo analisada pela Engenharia",
            "Enviado ao Poder Concedente - Nova emissao em analise Eng",
            "Concluido - Nova emissao em analise Eng",
            "Superado"
        )
    }
)

$resumo = [ordered]@{
    ProjetosProcessados = 0
    PastasEncontradas = 0
    PastasNaoEncontradas = 0
    PermissoesPastaAplicadas = 0
    PermissoesDocumentoAplicadas = 0
    PermissoesWorkflowStateAplicadas = 0
    Erros = 0
}

#-------------------------------------------------------
# Funcoes auxiliares
#-------------------------------------------------------
function Normalizar-Texto {
    param([object]$Valor)

    if ($null -eq $Valor) {
        return ''
    }

    return ([string]$Valor).Trim()
}

function Escrever-LogParametro {
    param(
        [string]$Nome,
        [object]$Valor
    )

    Write-Host ("{0,-35}: [{1}]" -f $Nome, ([string]$Valor))
}

function Incrementar-Contador {
    param(
        [string]$Nome,
        [int]$Valor = 1
    )

    $script:resumo[$Nome] = [int]$script:resumo[$Nome] + $Valor
}

function SolicitarNomeGrupo {
    param([string]$ValorPadrao)

    $ValorPadrao = Normalizar-Texto $ValorPadrao

    Write-Host ""
    Write-Host "===== GRUPO GLOBAL DE CORRECAO ====="
    Write-Host "O nome do grupo sempre deve ser informado/confirmado pelo usuario."
    $nomeInformado = Read-Host "Nome do grupo [$ValorPadrao]"
    $nomeInformado = Normalizar-Texto $nomeInformado

    if ([string]::IsNullOrWhiteSpace($nomeInformado)) {
        return $ValorPadrao
    }

    return $nomeInformado
}

function MontarCaminhoProjeto {
    param(
        [string]$NomeConcessao,
        [string]$Projeto,
        [string]$PastaRelativa
    )

    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projeto = Normalizar-Texto $Projeto
    $PastaRelativa = Normalizar-Texto $PastaRelativa

    $raizProjeto = "ENGENHARIA\$NomeConcessao\Projetos\$Projeto"

    if ([string]::IsNullOrWhiteSpace($PastaRelativa)) {
        return $raizProjeto
    }

    return "$raizProjeto\$PastaRelativa"
}

function ExibirResumoFinal {
    Write-Host ""
    Write-Host "===== RESUMO FINAL ====="
    Escrever-LogParametro -Nome "Projetos processados"                -Valor $resumo.ProjetosProcessados
    Escrever-LogParametro -Nome "Pastas encontradas"                  -Valor $resumo.PastasEncontradas
    Escrever-LogParametro -Nome "Pastas nao encontradas"              -Valor $resumo.PastasNaoEncontradas
    Escrever-LogParametro -Nome "Permissoes de pasta aplicadas"       -Valor $resumo.PermissoesPastaAplicadas
    Escrever-LogParametro -Nome "Permissoes de documento aplicadas"   -Valor $resumo.PermissoesDocumentoAplicadas
    Escrever-LogParametro -Nome "Permissoes workflow/state aplicadas" -Valor $resumo.PermissoesWorkflowStateAplicadas
    Escrever-LogParametro -Nome "Erros"                               -Valor $resumo.Erros
    Write-Host "========================="
    Write-Host ""
}

#-------------------------------------------------------
# Funcoes de grupo
#-------------------------------------------------------
function GarantirGrupoCorrecao {
    param([string]$NomeGrupo)

    $NomeGrupo = Normalizar-Texto $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($NomeGrupo)) {
        throw "Nome do grupo nao pode ficar vazio."
    }

    Write-Host ""
    Write-Host "Validando grupo global de correcao: $NomeGrupo"

    try {
        $grupo = Get-PWGroups -GroupName $NomeGrupo -ErrorAction Stop

        if (-not $grupo) {
            Write-Host "Grupo nao encontrado. Criando grupo: $NomeGrupo"
            $grupo = New-PWGroupByName -GroupName $NomeGrupo -GroupDescription "Grupo global para correcao de acesso somente leitura/download em documentos de projetos." -ErrorAction Stop
            Write-Host "Grupo criado com sucesso: $NomeGrupo"
        }
        else {
            Write-Host "Grupo ja existe e sera reutilizado: $NomeGrupo"
        }

        return $grupo
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao garantir grupo [$NomeGrupo]: $($_.Exception.Message)"
        throw
    }
}

#-------------------------------------------------------
# Funcoes de permissao
#-------------------------------------------------------
function TestarPastaPW {
    param([string]$FolderPath)

    $FolderPath = Normalizar-Texto $FolderPath

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        Incrementar-Contador -Nome "PastasNaoEncontradas"
        Write-Host "FolderPath vazio. Permissoes nao serao aplicadas."
        return $false
    }

    try {
        $pasta = Get-PWFolders -FolderPath $FolderPath -PopulatePaths -JustOne -ErrorAction Stop

        if (-not $pasta) {
            Incrementar-Contador -Nome "PastasNaoEncontradas"
            Write-Host "Pasta nao encontrada: $FolderPath"
            return $false
        }

        Incrementar-Contador -Nome "PastasEncontradas"
        Write-Host "Pasta encontrada: $FolderPath"
        return $true
    }
    catch {
        Incrementar-Contador -Nome "PastasNaoEncontradas"
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao validar pasta [$FolderPath]: $($_.Exception.Message)"
        return $false
    }
}

function AplicarFolderSecurity {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    try {
        Write-Host "Aplicando FolderSecurity para grupo [$NomeGrupo] em [$FolderPath] com acesso: $($folderAccess -join ', ')"

        Update-PWFolderSecurity `
            -Verbose `
            -InputFolder $FolderPath `
            -FolderSecurity `
            -MemberType Group `
            -MemberName $NomeGrupo `
            -MemberAccess $folderAccess `
            -ErrorAction Stop

        Incrementar-Contador -Nome "PermissoesPastaAplicadas"
        Write-Host "FolderSecurity aplicada com sucesso."
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao aplicar FolderSecurity em [$FolderPath]: $($_.Exception.Message)"
    }
}

function AplicarDocumentSecurityBasica {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    try {
        Write-Host "Aplicando DocumentSecurity para grupo [$NomeGrupo] em [$FolderPath] com acesso: $($documentAccess -join ', ')"

        Update-PWFolderSecurity `
            -Verbose `
            -InputFolder $FolderPath `
            -DocumentSecurity `
            -MemberType Group `
            -MemberName $NomeGrupo `
            -MemberAccess $documentAccess `
            -ErrorAction Stop

        Incrementar-Contador -Nome "PermissoesDocumentoAplicadas"
        Write-Host "DocumentSecurity aplicada com sucesso."
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao aplicar DocumentSecurity em [$FolderPath]: $($_.Exception.Message)"
    }
}

function AplicarDocumentSecurityWorkflowState {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    foreach ($workflowConfig in $workflowStatesConhecidos) {
        $workflow = $workflowConfig.Workflow

        foreach ($state in $workflowConfig.States) {
            try {
                Write-Host "Aplicando DocumentSecurity por workflow/state em [$FolderPath]"
                Escrever-LogParametro -Nome "Workflow" -Valor $workflow
                Escrever-LogParametro -Nome "State"    -Valor $state

                Update-PWFolderSecurity `
                    -Verbose `
                    -InputFolder $FolderPath `
                    -DocumentSecurity `
                    -WorkFlowName $workflow `
                    -StateName $state `
                    -MemberType Group `
                    -MemberName $NomeGrupo `
                    -MemberAccess $documentAccess `
                    -ErrorAction Stop

                Incrementar-Contador -Nome "PermissoesWorkflowStateAplicadas"
                Write-Host "DocumentSecurity workflow/state aplicada com sucesso."
            }
            catch {
                Incrementar-Contador -Nome "Erros"
                Write-Host "Erro ao aplicar workflow/state [$workflow / $state] em [$FolderPath]: $($_.Exception.Message)"
            }
        }
    }
}

#-------------------------------------------------------
# Funcao principal por projeto
#-------------------------------------------------------
function CorrigirAcessoProjeto {
    param(
        [string]$NomeConcessao,
        [string]$Projeto,
        [string]$NomeGrupo,
        [switch]$AplicarWorkflowStates
    )

    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projeto = Normalizar-Texto $Projeto
    $NomeGrupo = Normalizar-Texto $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($Projeto)) {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Projeto vazio recebido. Item ignorado."
        return
    }

    Incrementar-Contador -Nome "ProjetosProcessados"

    Write-Host ""
    Write-Host "===== INICIO DA CORRECAO DO PROJETO ====="
    Escrever-LogParametro -Nome "NomeConcessao" -Valor $NomeConcessao
    Escrever-LogParametro -Nome "Projeto"       -Valor $Projeto
    Escrever-LogParametro -Nome "Grupo"         -Valor $NomeGrupo
    Write-Host "========================================="

    foreach ($pastaRelativa in $pastasProjeto) {
        $folderPath = MontarCaminhoProjeto -NomeConcessao $NomeConcessao -Projeto $Projeto -PastaRelativa $pastaRelativa

        try {
            Write-Host ""
            Write-Host "----- Pasta alvo -----"
            Write-Host $folderPath

            if (-not (TestarPastaPW -FolderPath $folderPath)) {
                continue
            }

            AplicarFolderSecurity -FolderPath $folderPath -NomeGrupo $NomeGrupo
            AplicarDocumentSecurityBasica -FolderPath $folderPath -NomeGrupo $NomeGrupo

            if ($AplicarWorkflowStates) {
                AplicarDocumentSecurityWorkflowState -FolderPath $folderPath -NomeGrupo $NomeGrupo
            }
        }
        catch {
            Incrementar-Contador -Nome "Erros"
            Write-Host "Erro inesperado na pasta [$folderPath]: $($_.Exception.Message)"
            continue
        }
    }
}

#-------------------------------------------------------
# Inicio da execucao
#-------------------------------------------------------
$login = $null

try {
    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projetos = @($Projetos | ForEach-Object { Normalizar-Texto $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    Write-Host ""
    Write-Host "===== PARAMETROS RECEBIDOS ====="
    Escrever-LogParametro -Nome "NomeConcessao"          -Valor $NomeConcessao
    Escrever-LogParametro -Nome "Projetos"               -Valor ($Projetos -join ", ")
    Escrever-LogParametro -Nome "AplicarWorkflowStates"  -Valor $AplicarWorkflowStates
    Write-Host "================================"

    if ([string]::IsNullOrWhiteSpace($NomeConcessao)) {
        throw "Parametro NomeConcessao e obrigatorio."
    }

    if (-not $Projetos -or $Projetos.Count -eq 0) {
        throw "Parametro Projetos e obrigatorio e deve possuir pelo menos um projeto."
    }

    $NomeGrupo = SolicitarNomeGrupo -ValorPadrao $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($NomeGrupo)) {
        throw "Nome do grupo nao pode ficar vazio."
    }

    $login = New-PWLogin -ErrorAction Stop

    if (-not $login) {
        Write-Host "Login no ProjectWise nao realizado."
        return
    }

    $null = GarantirGrupoCorrecao -NomeGrupo $NomeGrupo

    foreach ($projeto in $Projetos) {
        CorrigirAcessoProjeto `
            -NomeConcessao $NomeConcessao `
            -Projeto $projeto `
            -NomeGrupo $NomeGrupo `
            -AplicarWorkflowStates:$AplicarWorkflowStates
    }
}
catch {
    Incrementar-Contador -Nome "Erros"
    Write-Host "Erro geral na execucao: $($_.Exception.Message)"
}
finally {
    if ($login) {
        Undo-PWLogin
    }

    ExibirResumoFinal
}
