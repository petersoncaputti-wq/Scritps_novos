Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# Renomear arquivos com revisão numérica no final do nome
# Exemplo:
#   DE-SP0000150-059.062-122-P02-101 1.pdf
# vira:
#   DE-SP0000150-059.062-122-P02-101-R00a.pdf
#
# Regra da revisão:
#   1  -> R00a
#   2  -> R00b
#   26 -> R00z
#   27 -> R01a
#   28 -> R01b
# ============================================================

# =========================
# Configurações
# =========================
$script:ExtensoesPermitidas = @('.dwg', '.pdf')
$script:UsarFiltroExtensao = $false
$script:ExibirDetalhadoNoConsole = $false
$script:ExportarCsv = $true
$script:HabilitarLogs = $true
$script:CaminhoLog = ''
$script:ArquivoLog = ''
$script:FalhaInicializacaoLog = $false

# Validação simples da base antes da revisão.
# Mantém o script independente das tabelas ANTT.
# Base esperada no exemplo: DE-SP0000150-059.062-122-P02-101
$script:ValidarEstruturaBase = $true

# =========================
# Funções de Logging
# =========================
function Obter-Timestamp {
    return Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
}

function Inicializar-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaminhoBase
    )

    if (-not $script:HabilitarLogs) { return }
    if ([string]::IsNullOrWhiteSpace($CaminhoBase)) {
        $script:FalhaInicializacaoLog = $true
        Write-Host "Aviso: pasta base vazia. O arquivo de log nao sera criado." -ForegroundColor Yellow
        return
    }

    $dataHora = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:CaminhoLog = Join-Path $CaminhoBase 'logs_renomeacao'
    $script:ArquivoLog = Join-Path $script:CaminhoLog "log_renomeacao_$dataHora.txt"

    try {
        if (-not (Test-Path -Path $script:CaminhoLog -PathType Container)) {
            New-Item -ItemType Directory -Path $script:CaminhoLog -Force | Out-Null
        }

        "" | Out-File -FilePath $script:ArquivoLog -Encoding UTF8
        Registrar-LogInfo "===== INICIANDO SCRIPT DE RENOMEAÇÃO ====="
        Registrar-LogInfo "Data/Hora: $(Obter-Timestamp)"
        Registrar-LogInfo "Computador: $env:COMPUTERNAME"
        Registrar-LogInfo "Usuário: $env:USERNAME"
    }
    catch {
        Write-Host "Aviso: Não foi possível criar arquivo de log." -ForegroundColor Yellow
    }
}

function Registrar-Log {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'AVISO', 'ERRO', 'SUCESSO')]
        [string]$Nivel,

        [Parameter(Mandatory = $true)]
        [string]$Mensagem,

        [System.ConsoleColor]$Cor = [System.ConsoleColor]::White
    )

    if (-not $script:HabilitarLogs) { return }

    $timestamp = Obter-Timestamp
    $linhaLog = "[$timestamp] [$Nivel] $Mensagem"

    if ($script:ArquivoLog) {
        try {
            Add-Content -Path $script:ArquivoLog -Value $linhaLog -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            if (-not $script:FalhaInicializacaoLog) {
                $script:FalhaInicializacaoLog = $true
                Write-Host "Aviso: falha ao gravar no arquivo de log: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    Write-Host $linhaLog -ForegroundColor $Cor
}

function Registrar-LogInfo {
    param([string]$Mensagem)
    Registrar-Log -Nivel 'INFO' -Mensagem $Mensagem -Cor Cyan
}

function Registrar-LogAviso {
    param([string]$Mensagem)
    Registrar-Log -Nivel 'AVISO' -Mensagem $Mensagem -Cor Yellow
}

function Registrar-LogErro {
    param(
        [string]$Mensagem,
        [System.Management.Automation.ErrorRecord]$Erro
    )

    if ($Erro) {
        $detalhe = " | Tipo: $($Erro.Exception.GetType().FullName) | Detalhe: $($Erro.Exception.Message)"
        Registrar-Log -Nivel 'ERRO' -Mensagem ($Mensagem + $detalhe) -Cor Red
        return
    }

    Registrar-Log -Nivel 'ERRO' -Mensagem $Mensagem -Cor Red
}

function Registrar-LogSucesso {
    param([string]$Mensagem)
    Registrar-Log -Nivel 'SUCESSO' -Mensagem $Mensagem -Cor Green
}

# =========================
# Medição de tempo
# =========================
function Medir-Tempo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Nome,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Bloco
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        return & $Bloco
    }
    finally {
        $sw.Stop()
        $tempo = [math]::Round($sw.Elapsed.TotalSeconds, 2)
        $mensagem = "{0}: {1}s" -f $Nome, $tempo
        Write-Host $mensagem
        Registrar-LogInfo $mensagem
    }
}

# =========================
# Normalizar caminho informado manualmente
# =========================
function Normalizar-CaminhoManual {
    param([string]$Caminho)

    if ([string]::IsNullOrWhiteSpace($Caminho)) { return $null }
    return $Caminho.Trim().Trim('"')
}

function Testar-ArtefatoInternoScript {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$Arquivo
    )

    if ($Arquivo.DirectoryName -match '\\logs_renomeacao(\\|$)') {
        return $true
    }

    if ($Arquivo.Name -like 'resultado_renomeacao_revisao_*.csv') {
        return $true
    }

    if ($Arquivo.Name -like 'log_renomeacao_*.txt' -or $Arquivo.Name -like 'logs_renomeacao_*.txt') {
        return $true
    }

    return $false
}

# =========================
# Selecionar pasta com fallback para rede/UNC
# =========================
function Selecionar-Pasta {
    Write-Host ""
    Write-Host "Seleção da pasta para análise"
    Write-Host "----------------------------------------"
    Write-Host "1 - Selecionar pelo navegador de pastas"
    Write-Host "2 - Informar caminho manualmente"
    Write-Host ""

    do {
        $opcao = Read-Host "Escolha uma opção (1/2)"
        $opcao = $opcao.Trim()
    } while ($opcao -notin @('1', '2'))

    if ($opcao -eq '1') {
        $dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialogo.Description = "Selecione a pasta que será analisada"
        $dialogo.ShowNewFolderButton = $false

        if ($dialogo.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selecionado = $dialogo.SelectedPath

            if (Test-Path -Path $selecionado -PathType Container) {
                return $selecionado
            }

            Write-Host ""
            Write-Host "A pasta selecionada não pôde ser acessada."
        }

        Write-Host ""
        Write-Host "Caso a pasta de rede não apareça no seletor, informe o caminho manualmente."
    }

    Write-Host ""
    Write-Host "Informe o caminho completo da pasta."
    Write-Host "Exemplos:"
    Write-Host "- Z:\Projetos\Documentos"
    Write-Host "- \\servidor\compartilhamento\pasta"
    Write-Host ""

    $tentativas = 0

    do {
        $tentativas++
        $caminhoManual = Normalizar-CaminhoManual -Caminho (Read-Host "Caminho da pasta")

        if (-not [string]::IsNullOrWhiteSpace($caminhoManual) -and (Test-Path -Path $caminhoManual -PathType Container)) {
            return $caminhoManual
        }

        Write-Host "Caminho inválido ou inacessível. Verifique se a pasta existe e se você tem permissão."
    } while ($tentativas -lt 3)

    return $null
}

# =========================
# Obter arquivos com subpastas
# =========================
function Obter-Arquivos {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Caminho,

        [string[]]$Extensoes = @('.dwg', '.pdf'),

        [bool]$FiltrarExtensoes = $true
    )

    if (-not (Test-Path -Path $Caminho -PathType Container)) {
        Registrar-LogErro "Caminho inexistente ou inacessivel durante a busca: $Caminho"
        return @()
    }

    $errosBusca = @()
    $arquivos = Get-ChildItem -Path $Caminho -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +errosBusca
    $arquivos = @($arquivos | Where-Object { -not (Testar-ArtefatoInternoScript -Arquivo $_) })

    if ($errosBusca.Count -gt 0) {
        Registrar-LogAviso "A busca encontrou $($errosBusca.Count) aviso(s)/erro(s) de acesso. Alguns arquivos ou pastas podem nao ter sido analisados."
        foreach ($erroBusca in $errosBusca) {
            Registrar-LogErro "Falha durante busca em '$Caminho'" -Erro $erroBusca
        }
    }

    if ($FiltrarExtensoes -and $Extensoes.Count -gt 0) {
        $extHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ext in $Extensoes) {
            [void]$extHash.Add($ext)
        }

        $arquivos = $arquivos | Where-Object { $extHash.Contains($_.Extension) }
    }

    $arquivosOrdenados = @($arquivos | Sort-Object FullName)
    Registrar-LogInfo "Busca concluida em '$Caminho'. Arquivos retornados: $($arquivosOrdenados.Count)"

    return $arquivosOrdenados
}



# =========================
# Contar arquivos totais e filtrados para diagnóstico
# =========================
function Obter-DiagnosticoBusca {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Caminho,

        [string[]]$Extensoes = @('.dwg', '.pdf')
    )

    $errosDiagnostico = @()
    $todos = @(Get-ChildItem -Path $Caminho -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +errosDiagnostico)
    $todos = @($todos | Where-Object { -not (Testar-ArtefatoInternoScript -Arquivo $_) })

    if ($errosDiagnostico.Count -gt 0) {
        Registrar-LogAviso "O diagnostico encontrou $($errosDiagnostico.Count) aviso(s)/erro(s) de acesso."
        foreach ($erroDiagnostico in $errosDiagnostico) {
            Registrar-LogErro "Falha durante diagnostico em '$Caminho'" -Erro $erroDiagnostico
        }
    }

    $extHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($ext in $Extensoes) {
        [void]$extHash.Add($ext)
    }

    $consideradosPeloFiltro = @($todos | Where-Object { $extHash.Contains($_.Extension) })
    $foraDoFiltro = @($todos | Where-Object { -not $extHash.Contains($_.Extension) })

    $diagnostico = [PSCustomObject]@{
        TotalNaPastaRecursivo = $todos.Count
        TotalComExtensaoPermitida = $consideradosPeloFiltro.Count
        TotalForaDoFiltro = $foraDoFiltro.Count
        ForaDoFiltroPorExtensao = @(
            $foraDoFiltro |
                Group-Object Extension |
                Sort-Object Count -Descending |
                ForEach-Object {
                    [PSCustomObject]@{
                        Extensao = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(sem extensão)' } else { $_.Name }
                        Quantidade = $_.Count
                    }
                }
        )
    }

    Registrar-LogInfo "Diagnostico concluido. Total: $($diagnostico.TotalNaPastaRecursivo) | Extensoes permitidas: $($diagnostico.TotalComExtensaoPermitida) | Fora do filtro: $($diagnostico.TotalForaDoFiltro)"

    return $diagnostico
}

# =========================
# Converter número para revisão R00a, R00b ... R01a
# =========================
function Converter-RevisaoNumerica {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Numero
    )

    if ($Numero -lt 1) {
        throw "Revisão numérica inválida: $Numero. O valor deve ser maior ou igual a 1."
    }

    $indice = $Numero - 1
    $numeroRevisao = [math]::Floor($indice / 26)
    $indiceLetra = $indice % 26
    $letra = [char]([int][char]'a' + $indiceLetra)

    return ("R{0:00}{1}" -f $numeroRevisao, $letra)
}

function Converter-RevisaoNumeroLetra {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Numero,

        [Parameter(Mandatory = $true)]
        [string]$Letra
    )

    if ($Numero -lt 0) {
        throw "Revisao numero/letra invalida: $Numero$Letra. O numero deve ser maior ou igual a 0."
    }

    if ($Letra -notmatch '^[a-zA-Z]$') {
        throw "Revisao numero/letra invalida: $Numero$Letra. A letra deve estar entre A e Z."
    }

    $numeroRevisao = $Numero - 1
    if ($numeroRevisao -lt 0) {
        $numeroRevisao = 0
    }

    return ("R{0:00}{1}" -f $numeroRevisao, $Letra.ToLower())
}

# =========================
# Montar objeto de retorno
# =========================
function Novo-ResultadoAnalise {
    param(
        [string]$NomeOriginal,
        [string]$NomeSugerido,
        [string]$CaminhoCompleto,
        [string]$Diretorio,
        [string]$Status,
        [array]$Erros,
        [array]$Avisos,
        [array]$Motivos,
        [bool]$PrecisaAjuste,
        [bool]$Cancelado,
        $Campos
    )

    return [PSCustomObject]@{
        NomeOriginal    = $NomeOriginal
        NomeSugerido    = $NomeSugerido
        CaminhoCompleto = $CaminhoCompleto
        Diretorio       = $Diretorio
        Status          = $Status
        Erros           = @($Erros)
        Avisos          = @($Avisos)
        Motivos         = @($Motivos)
        PrecisaAjuste   = $PrecisaAjuste
        Cancelado       = $Cancelado
        Campos          = $Campos
    }
}

# =========================
# Análise simples, sem padrão ANTT
# =========================
function Analisar-ArquivoRevisao {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$Arquivo
    )

    $nomeArquivo = $Arquivo.Name
    $caminhoCompleto = $Arquivo.FullName
    $diretorio = $Arquivo.DirectoryName
    $ext = $Arquivo.Extension
    $nomeBaseOriginal = [System.IO.Path]::GetFileNameWithoutExtension($nomeArquivo).Trim()

    $erros = [System.Collections.Generic.List[string]]::new()
    $avisos = [System.Collections.Generic.List[string]]::new()
    $motivos = [System.Collections.Generic.List[string]]::new()

    $base = ''
    $revisaoNumerica = $null
    $revisaoFormatada = ''
    $nomeBaseNovo = $nomeBaseOriginal
    $cancelado = $false

    # Caso já esteja no padrão final, considera válido.
    if ($nomeBaseOriginal -match '^(.+)-R\d{2}[a-z]$') {
        $base = $matches[1]
        $revisaoFormatada = ($nomeBaseOriginal -replace '^.+-(R\d{2}[a-z])$', '$1')
        $nomeBaseNovo = $nomeBaseOriginal
        $avisos.Add('Arquivo já está com revisão no padrão esperado.')
    }
    # Caso esteja com revisão cancelada, mantém sem alteração.
    elseif ($nomeBaseOriginal -match '^(.+)-RCA$') {
        $base = $matches[1]
        $revisaoFormatada = 'RCA'
        $nomeBaseNovo = $nomeBaseOriginal
        $cancelado = $true
        $avisos.Add('Arquivo cancelado. Não será renomeado.')
    }
    # Caso principal: base + espaço + número no final.
    elseif ($nomeBaseOriginal -match '^(.+?)_(\d{3})-(\d{1,2})$') {
        $base = "$($matches[1].Trim())-$($matches[2])"
        $revisaoNumerica = [int]$matches[3]

        try {
            $numeroParaConverter = $revisaoNumerica
            if ($numeroParaConverter -eq 0) {
                $numeroParaConverter = 1
                $avisos.Add("Revisao numerica '0' interpretada como primeira revisao.")
            }

            $revisaoFormatada = Converter-RevisaoNumerica -Numero $numeroParaConverter
            $nomeBaseNovo = "$base-$revisaoFormatada"
            $motivos.Add("Padrao com '_' e revisao por hifen normalizado. Revisao '$($matches[3])' convertida para '$revisaoFormatada'.")
        }
        catch {
            $erros.Add($_.Exception.Message)
        }
    }
    elseif ($nomeBaseOriginal -match '^(.+?)_(\d{3})\s+(\d+)$') {
        $base = "$($matches[1].Trim())-$($matches[2])"
        $revisaoNumerica = [int]$matches[3]

        try {
            $numeroParaConverter = $revisaoNumerica
            if ($numeroParaConverter -eq 0) {
                $numeroParaConverter = 1
                $avisos.Add("Revisao numerica '0' interpretada como primeira revisao.")
            }

            $revisaoFormatada = Converter-RevisaoNumerica -Numero $numeroParaConverter
            $nomeBaseNovo = "$base-$revisaoFormatada"
            $motivos.Add("Padrao com '_' normalizado. Revisao '$revisaoNumerica' convertida para '$revisaoFormatada'.")
        }
        catch {
            $erros.Add($_.Exception.Message)
        }
    }
    elseif ($nomeBaseOriginal -match '^(.+?)[\s-]+(\d+)([a-zA-Z])$') {
        $base = $matches[1].Trim()
        $revisaoNumerica = [int]$matches[2]
        $letraRevisao = $matches[3]

        try {
            $revisaoFormatada = Converter-RevisaoNumeroLetra -Numero $revisaoNumerica -Letra $letraRevisao
            $nomeBaseNovo = "$base-$revisaoFormatada"
            $motivos.Add("Revisao numero/letra '$revisaoNumerica$letraRevisao' convertida para '$revisaoFormatada'.")
        }
        catch {
            $erros.Add($_.Exception.Message)
        }
    }
    elseif ($nomeBaseOriginal -match '^(.+?)-(\d{1,2})$') {
        $base = $matches[1].Trim()
        $revisaoNumerica = [int]$matches[2]

        try {
            $numeroParaConverter = $revisaoNumerica
            if ($numeroParaConverter -eq 0) {
                $numeroParaConverter = 1
                $avisos.Add("Revisao numerica '0' interpretada como primeira revisao.")
            }

            $revisaoFormatada = Converter-RevisaoNumerica -Numero $numeroParaConverter
            $nomeBaseNovo = "$base-$revisaoFormatada"
            $motivos.Add("Revisao numerica com hifen '$($matches[2])' convertida para '$revisaoFormatada'.")
        }
        catch {
            $erros.Add($_.Exception.Message)
        }
    }
    elseif ($nomeBaseOriginal -match '^(.+?)\s+(\d+)$') {
        $base = $matches[1].Trim()
        $revisaoNumerica = [int]$matches[2]

        try {
            $numeroParaConverter = $revisaoNumerica
            if ($numeroParaConverter -eq 0) {
                $numeroParaConverter = 1
                $avisos.Add("Revisao numerica '0' interpretada como primeira revisao.")
            }

            $revisaoFormatada = Converter-RevisaoNumerica -Numero $numeroParaConverter
            $nomeBaseNovo = "$base-$revisaoFormatada"
            $motivos.Add("Revisão numérica '$revisaoNumerica' convertida para '$revisaoFormatada'.")
        }
        catch {
            $erros.Add($_.Exception.Message)
        }
    }
    else {
        $erros.Add('Não foi encontrada revisão numérica no final do nome. O esperado é: BASE 1, BASE 2, BASE 27 etc.')
    }

    # Validação leve da base, sem depender de tabelas ANTT.
    if ($script:ValidarEstruturaBase -and -not [string]::IsNullOrWhiteSpace($base)) {
        $partesBase = $base -split '-'

        if ($partesBase.Count -ne 6) {
            $erros.Add("Estrutura da base diferente do esperado. Foram encontrados $($partesBase.Count) blocos separados por '-', mas o exemplo esperado possui 6 blocos antes da revisão.")
        }

        if ($partesBase.Count -ge 6 -and $partesBase[5] -notmatch '^\d{3}$') {
            $erros.Add('Último bloco da base inválido. O esperado é uma sequência com 3 dígitos, como 101.')
        }
    }

    $nomeSugeridoCompleto = $nomeBaseNovo + $ext
    $precisaAjuste = ($nomeArquivo -cne $nomeSugeridoCompleto)

    if ($cancelado) {
        $status = 'Cancelado'
        $precisaAjuste = $false
        $nomeSugeridoCompleto = $nomeArquivo
    }
    elseif ($erros.Count -gt 0) {
        $status = 'Inválido'
        $precisaAjuste = $false
    }
    elseif ($precisaAjuste) {
        $status = 'Ajustável'
    }
    else {
        $status = 'Válido'
    }

    return Novo-ResultadoAnalise `
        -NomeOriginal $nomeArquivo `
        -NomeSugerido $nomeSugeridoCompleto `
        -CaminhoCompleto $caminhoCompleto `
        -Diretorio $diretorio `
        -Status $status `
        -Erros @($erros) `
        -Avisos @($avisos) `
        -Motivos @($motivos) `
        -PrecisaAjuste $precisaAjuste `
        -Cancelado $cancelado `
        -Campos ([PSCustomObject]@{
            Base             = $base
            RevisaoNumerica  = $revisaoNumerica
            RevisaoFormatada = $revisaoFormatada
        })
}

# =========================
# Processar análise
# =========================
function Processar-Analise {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Arquivos
    )

    $resultados = [System.Collections.Generic.List[object]]::new()

    foreach ($arquivo in $Arquivos) {
        $resultado = Analisar-ArquivoRevisao -Arquivo $arquivo
        $resultados.Add($resultado)
    }

    $resultadosArray = @($resultados)
    $gruposStatus = @($resultadosArray | Group-Object Status | Sort-Object Name)
    foreach ($grupo in $gruposStatus) {
        Registrar-LogInfo "Analise - Status '$($grupo.Name)': $($grupo.Count)"
    }

    $invalidos = @($resultadosArray | Where-Object { $_.Status -eq 'InvÃ¡lido' })
    foreach ($invalido in $invalidos) {
        Registrar-LogAviso "Arquivo invalido: $($invalido.CaminhoCompleto) | Erros: $($invalido.Erros -join ' | ')"
    }

    return $resultadosArray
}

# =========================
# Exibir resultado detalhado
# =========================
function Exibir-Resultado {
    param(
        [Parameter(Mandatory = $true)]
        $Resultado,

        [Parameter(Mandatory = $true)]
        [string]$PastaRaiz
    )

    $relativo = $Resultado.CaminhoCompleto.Replace($PastaRaiz, '').TrimStart('\')

    Write-Host "Arquivo: $relativo"
    Write-Host "Status : $($Resultado.Status)"
    Write-Host "Atual  : $($Resultado.NomeOriginal)"
    Write-Host "Novo   : $($Resultado.NomeSugerido)"
    Write-Host "Base   : $($Resultado.Campos.Base)"
    Write-Host "Revisão: $($Resultado.Campos.RevisaoFormatada)"

    if ($Resultado.Motivos.Count -gt 0) {
        Write-Host "Motivos do ajuste:"
        foreach ($motivo in $Resultado.Motivos) { Write-Host "- $motivo" }
    }

    if ($Resultado.Avisos.Count -gt 0) {
        Write-Host "Observações:"
        foreach ($aviso in $Resultado.Avisos) { Write-Host "- $aviso" }
    }

    if ($Resultado.Erros.Count -gt 0) {
        Write-Host "Erros:"
        foreach ($erro in $Resultado.Erros) { Write-Host "- $erro" }
    }

    Write-Host "----------------------------------------"
}

# =========================
# Exibir resumo
# =========================
function Exibir-Resumo {
    param([Parameter(Mandatory = $true)][array]$Resultados)

    $validos = ($Resultados | Where-Object { $_.Status -eq 'Válido' }).Count
    $ajustaveis = ($Resultados | Where-Object { $_.Status -eq 'Ajustável' }).Count
    $cancelados = ($Resultados | Where-Object { $_.Status -eq 'Cancelado' }).Count
    $invalidos = ($Resultados | Where-Object { $_.Status -eq 'Inválido' }).Count
    $precisamAjuste = ($Resultados | Where-Object { $_.PrecisaAjuste -eq $true }).Count

    Write-Host ""
    Write-Host "========== RESUMO =========="
    Write-Host "Total de arquivos analisados : $($Resultados.Count)"
    Write-Host "Válidos                     : $validos"
    Write-Host "Ajustáveis                  : $ajustaveis"
    Write-Host "Cancelados                  : $cancelados"
    Write-Host "Inválidos                   : $invalidos"
    Write-Host "Precisam de ajuste          : $precisamAjuste"
    Write-Host "============================"
    Write-Host ""

    Registrar-LogInfo "Resumo da análise - Total: $($Resultados.Count) | Válidos: $validos | Ajustáveis: $ajustaveis | Cancelados: $cancelados | Inválidos: $invalidos"
}

# =========================
# Preparar dados para exportação CSV
# =========================
function Converter-ResultadosParaCsv {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Resultados,

        [Parameter(Mandatory = $true)]
        [string]$PastaRaiz
    )

    foreach ($r in $Resultados) {
        $relativo = $r.CaminhoCompleto.Replace($PastaRaiz, '').TrimStart('\')

        [PSCustomObject]@{
            NomeOriginal      = $r.NomeOriginal
            NomeSugerido      = $r.NomeSugerido
            CaminhoRelativo   = $relativo
            Diretorio         = $r.Diretorio
            Status            = $r.Status
            PrecisaAjuste     = $r.PrecisaAjuste
            Cancelado         = $r.Cancelado
            Base              = $r.Campos.Base
            RevisaoNumerica   = $r.Campos.RevisaoNumerica
            RevisaoFormatada  = $r.Campos.RevisaoFormatada
            Motivos           = ($r.Motivos -join ' | ')
            Avisos            = ($r.Avisos -join ' | ')
            Erros             = ($r.Erros -join ' | ')
            CaminhoCompleto   = $r.CaminhoCompleto
        }
    }
}

# =========================
# Exportar CSV
# =========================
function Exportar-ResultadoCsv {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Resultados,

        [Parameter(Mandatory = $true)]
        [string]$PastaRaiz
    )

    $dataHora = Get-Date -Format 'yyyyMMdd_HHmmss'
    $arquivoCsv = Join-Path $PastaRaiz "resultado_renomeacao_revisao_$dataHora.csv"

    $dadosCsv = Converter-ResultadosParaCsv -Resultados $Resultados -PastaRaiz $PastaRaiz
    try {
        $dadosCsv | Export-Csv -Path $arquivoCsv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Registrar-LogSucesso "CSV exportado com sucesso: $arquivoCsv"
    }
    catch {
        Registrar-LogErro "Falha ao exportar CSV: $arquivoCsv" -Erro $_
        throw
    }

    Write-Host "Resultado exportado para:"
    Write-Host $arquivoCsv
    Write-Host ""

    return $arquivoCsv
}

# =========================
# Renomear arquivos ajustáveis
# =========================
function Renomear-ArquivosAjustaveis {
    param([Parameter(Mandatory = $true)][array]$Resultados)

    $arquivosParaRenomear = @(
        $Resultados | Where-Object {
            $_.PrecisaAjuste -eq $true -and
            $_.Cancelado -eq $false -and
            $_.Status -eq 'Ajustável' -and
            $_.NomeOriginal -cne $_.NomeSugerido
        }
    )

    if ($arquivosParaRenomear.Count -eq 0) {
        Write-Host ""
        Write-Host "Nenhum arquivo precisa ser renomeado."
        Registrar-LogInfo "Nenhum arquivo requer renomeação."
        return
    }

    Write-Host ""
    Write-Host "Arquivos que serão renomeados: $($arquivosParaRenomear.Count)"
    Write-Host "----------------------------------------"
    Registrar-LogInfo "Preparando renomeação de $($arquivosParaRenomear.Count) arquivo(s)."

    foreach ($item in $arquivosParaRenomear) {
        Write-Host "Pasta: $($item.Diretorio)"
        Write-Host "De  : $($item.NomeOriginal)"
        Write-Host "Para: $($item.NomeSugerido)"
        Write-Host ""
    }

    do {
        $confirmacaoRenomear = Read-Host "Deseja aplicar os ajustes diretamente nos arquivos encontrados? (S/N)"
        $confirmacaoRenomear = $confirmacaoRenomear.ToUpper().Trim()
    } while ($confirmacaoRenomear -notin @('S','N'))

    if ($confirmacaoRenomear -eq 'N') {
        Registrar-LogAviso "Renomeação cancelada pelo usuário. Nenhum arquivo foi alterado."
        Write-Host "Renomeação cancelada pelo usuário. Nenhum arquivo foi alterado."
        return
    }

    Write-Host ""
    Write-Host "Iniciando renomeação..."
    Write-Host "----------------------------------------"
    Registrar-LogInfo "Iniciando renomeação de $($arquivosParaRenomear.Count) arquivo(s)."

    $renomeados = 0
    $falhas = 0

    foreach ($item in $arquivosParaRenomear) {
        $caminhoOriginal = $item.CaminhoCompleto
        $caminhoNovo = Join-Path $item.Diretorio $item.NomeSugerido

        if (-not (Test-Path -LiteralPath $caminhoOriginal -PathType Leaf)) {
            $mensagem = "Arquivo não encontrado: $caminhoOriginal"
            Write-Host $mensagem
            Registrar-LogErro $mensagem
            $falhas++
            continue
        }

        if (Test-Path -LiteralPath $caminhoNovo -PathType Leaf) {
            $mensagem = "Não foi possível renomear '$($item.NomeOriginal)' porque já existe '$($item.NomeSugerido)' na mesma pasta."
            Write-Host $mensagem
            Registrar-LogErro $mensagem
            $falhas++
            continue
        }

        try {
            Rename-Item -LiteralPath $caminhoOriginal -NewName $item.NomeSugerido -ErrorAction Stop
            $mensagem = "Renomeado: $($item.NomeOriginal) -> $($item.NomeSugerido)"
            Write-Host $mensagem
            Registrar-LogSucesso $mensagem
            $renomeados++
        }
        catch {
            $mensagem = "Erro ao renomear '$($item.NomeOriginal)': $($_.Exception.Message)"
            Write-Host $mensagem
            Registrar-LogErro $mensagem -Erro $_
            $falhas++
        }
    }

    Write-Host ""
    Write-Host "Resumo da renomeação"
    Write-Host "----------------------------------------"
    Write-Host "Renomeados com sucesso: $renomeados"
    Write-Host "Falhas                : $falhas"
    Registrar-LogInfo "Resumo da renomeação - Sucesso: $renomeados | Falhas: $falhas"
}

# =========================
# Execução
# =========================
try {
    $pastaSelecionada = Selecionar-Pasta

    if (-not $pastaSelecionada) {
        $mensagem = "Nenhuma pasta válida foi selecionada ou informada."
        Write-Host $mensagem
        Registrar-LogErro $mensagem
        Read-Host "Pressione Enter para sair"
        exit
    }

    Inicializar-Log -CaminhoBase $pastaSelecionada
    Registrar-LogInfo "Pasta selecionada: $pastaSelecionada"

    $diagnosticoBusca = Medir-Tempo 'Diagnóstico da busca' {
        Obter-DiagnosticoBusca `
            -Caminho $pastaSelecionada `
            -Extensoes $script:ExtensoesPermitidas
    }

    $arquivos = Medir-Tempo 'Busca de arquivos' {
        Obter-Arquivos `
            -Caminho $pastaSelecionada `
            -Extensoes $script:ExtensoesPermitidas `
            -FiltrarExtensoes $script:UsarFiltroExtensao
    }

    Write-Host "----------------------------------------"
    Write-Host "Pasta selecionada:"
    Write-Host $pastaSelecionada
    Write-Host "----------------------------------------"
    Write-Host "Total real de arquivos na pasta e subpastas: $($diagnosticoBusca.TotalNaPastaRecursivo)"
    Write-Host "Filtro de extensão ativo: $script:UsarFiltroExtensao"

    if ($script:UsarFiltroExtensao) {
        Write-Host "Extensões consideradas: $($script:ExtensoesPermitidas -join ', ')"
        Write-Host "Arquivos dentro do filtro: $($diagnosticoBusca.TotalComExtensaoPermitida)"
        Write-Host "Arquivos fora do filtro   : $($diagnosticoBusca.TotalForaDoFiltro)"
        Registrar-LogInfo "Filtro de extensão ativo: $($script:ExtensoesPermitidas -join ', ') | Dentro do filtro: $($diagnosticoBusca.TotalComExtensaoPermitida) | Fora: $($diagnosticoBusca.TotalForaDoFiltro)"

        if ($diagnosticoBusca.TotalForaDoFiltro -gt 0) {
            Write-Host "Extensões fora do filtro:"
            foreach ($grupo in $diagnosticoBusca.ForaDoFiltroPorExtensao) {
                Write-Host ("- {0}: {1}" -f $grupo.Extensao, $grupo.Quantidade)
            }
        }
    }
    else {
        Write-Host "Extensões consideradas: todas"
        Registrar-LogInfo "Filtro de extensão desativado - analisando todas as extensões"
    }

    Write-Host "Quantidade de arquivos encontrados para análise: $($arquivos.Count)"
    Write-Host "----------------------------------------"

    if ($arquivos.Count -eq 0) {
        $mensagem = "Nenhum arquivo encontrado para análise."
        Write-Host $mensagem
        Registrar-LogAviso $mensagem
        Read-Host "Pressione Enter para encerrar"
        exit
    }

    Registrar-LogInfo "Arquivos encontrados: $($arquivos.Count)"

    do {
        $confirmacao = Read-Host "Deseja continuar para a validação e simulação dos ajustes? (S/N)"
        $confirmacao = $confirmacao.ToUpper().Trim()
    } while ($confirmacao -notin @('S', 'N'))

    if ($confirmacao -eq 'N') {
        Registrar-LogAviso "Processo cancelado pelo usuário na etapa de validação."
        Write-Host "Processo cancelado pelo usuário."
        Read-Host "Pressione Enter para encerrar"
        exit
    }

    Registrar-LogInfo "Usuário confirmou para prosseguir com a análise dos arquivos."

    $resultados = Medir-Tempo 'Análise dos arquivos' {
        Processar-Analise -Arquivos $arquivos
    }

    if ($script:ExibirDetalhadoNoConsole) {
        foreach ($resultado in $resultados) {
            Exibir-Resultado -Resultado $resultado -PastaRaiz $pastaSelecionada
        }
    }
    else {
        Write-Host "Exibição detalhada no console desativada. Use o CSV gerado para analisar arquivo por arquivo."
        Write-Host ""
    }

    Exibir-Resumo -Resultados $resultados

    if ($script:ExportarCsv) {
        $arquivoCsv = Exportar-ResultadoCsv -Resultados $resultados -PastaRaiz $pastaSelecionada
        Registrar-LogSucesso "Resultado exportado para CSV: $arquivoCsv"
    }

    Renomear-ArquivosAjustaveis -Resultados $resultados
}
catch {
    $mensagemErro = "Ocorreu um erro durante a execução do script: $($_.Exception.Message)"
    Write-Host "Ocorreu um erro durante a execução do script."
    Write-Host $_.Exception.Message
    Registrar-LogErro $mensagemErro -Erro $_
}
finally {
    if ($script:HabilitarLogs -and $script:ArquivoLog) {
        Registrar-LogInfo "===== FINALIZANDO SCRIPT ====="
        Write-Host ""
        Write-Host "Arquivo de log salvo em: $($script:ArquivoLog)" -ForegroundColor Green
    }
    Read-Host "Pressione Enter para encerrar"
}
