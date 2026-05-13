Add-Type -AssemblyName System.Windows.Forms

# =========================
# Configurações
# =========================
$script:ExtensoesPermitidas = @('.dwg', '.pdf')
$script:UsarFiltroExtensao = $true
$script:ExibirDetalhadoNoConsole = $false
$script:ExportarCsv = $true
$script:AplicarParalelismo = $false
$script:ThrottleLimit = 6

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
        Write-Host ("{0}: {1}s" -f $Nome, [math]::Round($sw.Elapsed.TotalSeconds, 2))
    }
}

# =========================
# Normalizar caminho informado manualmente
# =========================
function Normalizar-CaminhoManual {
    param(
        [string]$Caminho
    )

    if ([string]::IsNullOrWhiteSpace($Caminho)) { return $null }

    $caminhoLimpo = $Caminho.Trim().Trim('"')

    return $caminhoLimpo
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
        return @()
    }

    $arquivos = Get-ChildItem -Path $Caminho -File -Recurse -ErrorAction SilentlyContinue

    if ($FiltrarExtensoes -and $Extensoes.Count -gt 0) {
        $extHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ext in $Extensoes) {
            [void]$extHash.Add($ext)
        }

        $arquivos = $arquivos | Where-Object { $extHash.Contains($_.Extension) }
    }

    return $arquivos | Sort-Object FullName
}

# =========================
# Tabelas ANTT
# =========================
function Obter-TabelasANTT {
    return @{
        Concessionarias = @(
            'DUT','TER','CRT','TPA','ECS','APS','ALS','ARB','AFD','AFL',
            'TBR','RAC','VBA','EC1','EC5','TRA','CRO','MSV','V40','ECP',
            'VSL','ECC','VCO','ECA','RSP','VBR','ERM','ARA','ELP','EVM',
            'CRN','CRC','ZEB'
        )

        TiposObra = @(
            'ACE','ALG','ACA','ANE','BAL','BSO','CCO','CIC','CNT','CTO',
            'CTR','DES','SEG','DDE','DNI','DRE','DPL','EDI','TAL','TRA',
            'FAD','ILU','INT','MOP','MON','EME','OAC','OAE','PFA','PIN',
            'PAS','AMB','PAV','PON','PCR','POR','PFR','PFI','PIF','PPV',
            'PDA','SIN','REC','RET','SAU','ITS','TER','TRV','PRF','VAR',
            'MAR','MAN','MAS'
        )

        TiposProjetoFixos = @(
            'ANT','ASB','DUP','EVT','EXO','EXE','FUN','OUT','PGT','PIT'
        )

        Classes = @(
            'BO','CR','DE','EM','ES','ET','EE','GR','IA','ID','IP','IT',
            'LM','MC','MD','NS','PL','PP','RV','RA','RM','RC','RF','RT'
        )

        Disciplinas = @(
            'A0','Y1','Z0','X9','X6','X5','X3','A2','I1','N1','X1','K1',
            'E1','T1','D2','D1','H1','A3','C1','Q2','C2','X4','L1','C3',
            'X2','Z9','X7','P5','P1','P6','O1','P7','L4','P4','B2','D3',
            'B1','T2','Q1','O2','H2','P2','N3','L3','L2','F1','P3','K2',
            'P8','M1','J2','N2','I2','O3','J1','K3','G1','D4','A1','V1',
            'V2','V3','V4','V5'
        )
    }
}

# =========================
# Converter tabelas para HashSet
# =========================
function Converter-TabelasParaHashSet {
    param(
        [Parameter(Mandatory = $true)]
        $Tabelas
    )

    $resultado = @{}

    foreach ($chave in $Tabelas.Keys) {
        $hash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($valor in $Tabelas[$chave]) {
            [void]$hash.Add($valor)
        }

        $resultado[$chave] = $hash
    }

    return $resultado
}

# =========================
# Ajuste de revisão / normalização
# =========================
function Ajustar-Revisao {
    param (
        [Parameter(Mandatory = $true)]
        [string]$NomeBase
    )

    $novoNome = $NomeBase
    $motivos = [System.Collections.Generic.List[string]]::new()
    $alterado = $false

    # _RCA -> -RCA
    if ($novoNome -match '_RCA$') {
        $novoNome = $novoNome -replace '_RCA$', '-RCA'
        $motivos.Add('Separador "_" antes da revisão cancelada ajustado para "-"')
        $alterado = $true
    }

    # Se terminar com RCA, não aplicar outras regras
    if ($novoNome -match '-RCA$') {
        return [PSCustomObject]@{
            NomeAjustado = $novoNome
            Alterado     = $alterado
            Motivos      = @($motivos)
        }
    }

    # _R00a / _R00A / _R00 -> -R...
    if ($novoNome -match '_R\d{2}[A-Za-z]?$') {
        $novoNome = $novoNome -replace '_R', '-R'
        $motivos.Add('Separador "_" antes da revisão ajustado para "-"')
        $alterado = $true
    }

    # -R00 -> -R00a
    if ($novoNome -match '-R\d{2}$') {
        $novoNome = $novoNome + 'a'
        $motivos.Add('Revisão sem letra final ajustada para "R00a"')
        $alterado = $true
    }

    # -R00A -> -R00a
    if ($novoNome -match '-R\d{2}[A-Z]$') {
        $novoNome = [regex]::Replace(
            $novoNome,
            '-R(\d{2})([A-Z])$',
            {
                param($m)
                "-R$($m.Groups[1].Value)$($m.Groups[2].Value.ToLower())"
            }
        )
        $motivos.Add('Letra final da revisão ajustada para minúscula')
        $alterado = $true
    }

    return [PSCustomObject]@{
        NomeAjustado = $novoNome
        Alterado     = $alterado
        Motivos      = @($motivos)
    }
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
# Análise ANTT
# =========================
function Analisar-ArquivoANTT {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$Arquivo,

        [Parameter(Mandatory = $true)]
        $Tabelas
    )

    $nomeArquivo = $Arquivo.Name
    $caminhoCompleto = $Arquivo.FullName
    $diretorio = $Arquivo.DirectoryName
    $ext = $Arquivo.Extension
    $nomeBaseOriginal = [System.IO.Path]::GetFileNameWithoutExtension($nomeArquivo).Trim()

    # Ajusta primeiro, valida depois
    $ajuste = Ajustar-Revisao -NomeBase $nomeBaseOriginal
    $nomeBase = $ajuste.NomeAjustado
    $nomeSugeridoCompleto = $nomeBase + $ext

    # Só considera que precisa ajuste se o nome realmente muda
    $precisaAjusteReal = ($nomeArquivo -cne $nomeSugeridoCompleto)

    $erros = [System.Collections.Generic.List[string]]::new()
    $avisos = [System.Collections.Generic.List[string]]::new()
    $status = 'Válido'
    $cancelado = $false

    $concessionaria = ''
    $rodoviaUF = ''
    $localizacao = ''
    $tipoObra = ''
    $tipoProjeto = ''
    $classe = ''
    $disciplina = ''
    $sequencia = ''
    $revisao = ''

    $partes = $nomeBase -split '-'

    if ($partes.Count -eq 9) {
        $concessionaria = $partes[0].ToUpper()
        $rodoviaUF      = $partes[1].ToUpper()
        $localizacao    = $partes[2]
        $tipoObra       = $partes[3].ToUpper()
        $tipoProjeto    = $partes[4].ToUpper()
        $classe         = $partes[5].ToUpper()
        $disciplina     = $partes[6].ToUpper()
        $sequencia      = $partes[7]
        $revisao        = $partes[8]
    }
    elseif ($partes.Count -eq 10) {
        $concessionaria = $partes[0].ToUpper()
        $rodoviaUF      = $partes[1].ToUpper()
        $localizacao    = "$($partes[2])-$($partes[3])"
        $tipoObra       = $partes[4].ToUpper()
        $tipoProjeto    = $partes[5].ToUpper()
        $classe         = $partes[6].ToUpper()
        $disciplina     = $partes[7].ToUpper()
        $sequencia      = $partes[8]
        $revisao        = $partes[9]
    }
    else {
        return Novo-ResultadoAnalise `
            -NomeOriginal $nomeArquivo `
            -NomeSugerido $nomeSugeridoCompleto `
            -CaminhoCompleto $caminhoCompleto `
            -Diretorio $diretorio `
            -Status 'Inválido' `
            -Erros @('Estrutura inválida. O nome deve seguir o padrão ANTT com 9 blocos.') `
            -Avisos @() `
            -Motivos $(if ($precisaAjusteReal) { $ajuste.Motivos } else { @() }) `
            -PrecisaAjuste $precisaAjusteReal `
            -Cancelado $false `
            -Campos ([PSCustomObject]@{
                Concessionaria = ''
                RodoviaUF      = ''
                Localizacao    = ''
                TipoObra       = ''
                TipoProjeto    = ''
                Classe         = ''
                Disciplina     = ''
                Sequencia      = ''
                Revisao        = ''
            })
    }

    if ($concessionaria.Length -ne 3) { $erros.Add('Concessionária deve ter 3 caracteres.') }
    if ($rodoviaUF.Length -ne 5)      { $erros.Add('Rodovia/UF deve ter 5 caracteres.') }
    if ($localizacao.Length -ne 7)    { $erros.Add('Localização deve ter 7 caracteres.') }
    if ($tipoObra.Length -ne 3)       { $erros.Add('Tipo de obra deve ter 3 caracteres.') }
    if ($tipoProjeto.Length -ne 3)    { $erros.Add('Tipo de projeto deve ter 3 caracteres.') }
    if ($classe.Length -ne 2)         { $erros.Add('Classe deve ter 2 caracteres.') }
    if ($disciplina.Length -ne 2)     { $erros.Add('Disciplina deve ter 2 caracteres.') }
    if ($sequencia.Length -ne 3)      { $erros.Add('Sequência deve ter 3 caracteres.') }

    if (-not $Tabelas.Concessionarias.Contains($concessionaria)) {
        $erros.Add('Concessionária inválida.')
    }

    if ($rodoviaUF -notmatch '^\d{3}[A-Z]{2}$') {
        $erros.Add('Rodovia/UF inválida. Deve seguir o padrão 000UF.')
    }

    if ($localizacao -notmatch '^\d{3}[-+]\d{3}$') {
        $erros.Add('Localização inválida. Deve seguir 000-000 ou 000+000.')
    }
    else {
        if ($localizacao -match '^\d{3}-\d{3}$') {
            $avisos.Add('Localização em trecho: indica extensão entre os quilômetros informados.')
        }
        elseif ($localizacao -match '^\d{3}\+\d{3}$') {
            $avisos.Add('Localização pontual: indica ponto específico conforme os quilômetros informados.')
        }
    }

    if (-not $Tabelas.TiposObra.Contains($tipoObra)) {
        $erros.Add('Tipo de obra inválido.')
    }

    $tipoProjetoValido = $false
    if ($Tabelas.TiposProjetoFixos.Contains($tipoProjeto)) {
        $tipoProjetoValido = $true
    }
    elseif ($tipoProjeto -match '^AL\d$') {
        $tipoProjetoValido = $true
    }

    if (-not $tipoProjetoValido) {
        $erros.Add('Tipo de projeto inválido.')
    }

    if (-not $Tabelas.Classes.Contains($classe)) {
        $erros.Add('Classe inválida.')
    }

    if (-not $Tabelas.Disciplinas.Contains($disciplina)) {
        $erros.Add('Disciplina inválida.')
    }

    if ($sequencia -notmatch '^\d{3}$') {
        $erros.Add('Sequência inválida. Deve conter 3 dígitos numéricos.')
    }

    if ($revisao -eq 'RCA') {
        $cancelado = $true
        $status = 'Cancelado'
        $avisos.Add('Arquivo cancelado. Não deve ser alterado.')
    }
    elseif ($revisao -notmatch '^R\d{2}[a-z]$') {
        $erros.Add('Revisão inválida. O padrão esperado é como R00a, com apenas a última letra minúscula.')
    }

    if ($erros.Count -gt 0) {
        $status = 'Inválido'
    }
    elseif ($precisaAjusteReal -and -not $cancelado) {
        $status = 'Ajustável'
    }
    elseif ($cancelado) {
        $status = 'Cancelado'
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
        -Motivos $(if ($precisaAjusteReal) { $ajuste.Motivos } else { @() }) `
        -PrecisaAjuste $precisaAjusteReal `
        -Cancelado $cancelado `
        -Campos ([PSCustomObject]@{
            Concessionaria = $concessionaria
            RodoviaUF      = $rodoviaUF
            Localizacao    = $localizacao
            TipoObra       = $tipoObra
            TipoProjeto    = $tipoProjeto
            Classe         = $classe
            Disciplina     = $disciplina
            Sequencia      = $sequencia
            Revisao        = $revisao
        })
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
            NomeOriginal    = $r.NomeOriginal
            NomeSugerido    = $r.NomeSugerido
            CaminhoRelativo = $relativo
            Diretorio       = $r.Diretorio
            Status          = $r.Status
            PrecisaAjuste   = $r.PrecisaAjuste
            Cancelado       = $r.Cancelado
            Concessionaria  = $r.Campos.Concessionaria
            RodoviaUF       = $r.Campos.RodoviaUF
            Localizacao     = $r.Campos.Localizacao
            TipoObra        = $r.Campos.TipoObra
            TipoProjeto     = $r.Campos.TipoProjeto
            Classe          = $r.Campos.Classe
            Disciplina      = $r.Campos.Disciplina
            Sequencia       = $r.Campos.Sequencia
            Revisao         = $r.Campos.Revisao
            Motivos         = ($r.Motivos -join ' | ')
            Avisos          = ($r.Avisos -join ' | ')
            Erros           = ($r.Erros -join ' | ')
            CaminhoCompleto = $r.CaminhoCompleto
        }
    }
}

# =========================
# Exibir resultado detalhado
# =========================
function Exibir-Resultado {
    param (
        [Parameter(Mandatory = $true)]
        $Resultado,

        [Parameter(Mandatory = $true)]
        [string]$PastaRaiz
    )

    $caminhoRelativo = $Resultado.CaminhoCompleto.Replace($PastaRaiz, '').TrimStart('\')

    Write-Host "Arquivo: $($Resultado.NomeOriginal)"
    Write-Host "Local : $caminhoRelativo"
    Write-Host "Status: $($Resultado.Status)"

    if ($Resultado.PrecisaAjuste) {
        Write-Host "Novo nome sugerido: $($Resultado.NomeSugerido)"
    }

    Write-Host "Concessionária: $($Resultado.Campos.Concessionaria)"
    Write-Host "Rodovia/UF     : $($Resultado.Campos.RodoviaUF)"
    Write-Host "Localização    : $($Resultado.Campos.Localizacao)"
    Write-Host "Tipo de Obra   : $($Resultado.Campos.TipoObra)"
    Write-Host "Tipo de Projeto: $($Resultado.Campos.TipoProjeto)"
    Write-Host "Classe         : $($Resultado.Campos.Classe)"
    Write-Host "Disciplina     : $($Resultado.Campos.Disciplina)"
    Write-Host "Sequência      : $($Resultado.Campos.Sequencia)"
    Write-Host "Revisão        : $($Resultado.Campos.Revisao)"

    if ($Resultado.Motivos.Count -gt 0) {
        Write-Host "Motivos do ajuste:"
        foreach ($motivo in $Resultado.Motivos) {
            Write-Host "- $motivo"
        }
    }

    if ($Resultado.Avisos.Count -gt 0) {
        Write-Host "Observações:"
        foreach ($aviso in $Resultado.Avisos) {
            Write-Host "- $aviso"
        }
    }

    if ($Resultado.Erros.Count -gt 0) {
        Write-Host "Erros:"
        foreach ($erro in $Resultado.Erros) {
            Write-Host "- $erro"
        }
    }

    Write-Host "----------------------------------------"
}

# =========================
# Exibir resumo
# =========================
function Exibir-Resumo {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Resultados
    )

    Write-Host ""
    Write-Host "========== RESUMO =========="
    Write-Host "Total de arquivos analisados : $($Resultados.Count)"
    Write-Host "Válidos                     : $(($Resultados | Where-Object { $_.Status -eq 'Válido' }).Count)"
    Write-Host "Ajustáveis                  : $(($Resultados | Where-Object { $_.Status -eq 'Ajustável' }).Count)"
    Write-Host "Cancelados                  : $(($Resultados | Where-Object { $_.Status -eq 'Cancelado' }).Count)"
    Write-Host "Inválidos                   : $(($Resultados | Where-Object { $_.Status -eq 'Inválido' }).Count)"
    Write-Host "Precisam de ajuste          : $(($Resultados | Where-Object { $_.PrecisaAjuste -eq $true }).Count)"
    Write-Host "============================"
    Write-Host ""
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
    $arquivoCsv = Join-Path $PastaRaiz "resultado_analise_antt_$dataHora.csv"

    $dadosCsv = Converter-ResultadosParaCsv -Resultados $Resultados -PastaRaiz $PastaRaiz
    $dadosCsv | Export-Csv -Path $arquivoCsv -NoTypeInformation -Encoding UTF8

    Write-Host "Resultado exportado para:"
    Write-Host $arquivoCsv
    Write-Host ""

    return $arquivoCsv
}

# =========================
# Renomear arquivos ajustáveis
# =========================
function Renomear-ArquivosAjustaveis {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Resultados
    )

    $arquivosParaRenomear = @(
        $Resultados | Where-Object {
            $_.PrecisaAjuste -eq $true -and
            $_.Cancelado -eq $false -and
            $_.NomeOriginal -cne $_.NomeSugerido
        }
    )

    if ($arquivosParaRenomear.Count -eq 0) {
        Write-Host ""
        Write-Host "Nenhum arquivo precisa ser renomeado."
        return
    }

    Write-Host ""
    Write-Host "Arquivos que serão renomeados: $($arquivosParaRenomear.Count)"
    Write-Host "----------------------------------------"

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
        Write-Host "Renomeação cancelada pelo usuário."
        return
    }

    Write-Host ""
    Write-Host "Iniciando renomeação..."
    Write-Host "----------------------------------------"

    $renomeados = 0
    $falhas = 0

    foreach ($item in $arquivosParaRenomear) {
        $caminhoOriginal = $item.CaminhoCompleto
        $caminhoNovo = Join-Path $item.Diretorio $item.NomeSugerido

        if (-not (Test-Path -Path $caminhoOriginal -PathType Leaf)) {
            Write-Host "Arquivo não encontrado: $caminhoOriginal"
            $falhas++
            continue
        }

        if (Test-Path -Path $caminhoNovo -PathType Leaf) {
            Write-Host "Não foi possível renomear '$($item.NomeOriginal)' porque já existe '$($item.NomeSugerido)' na mesma pasta."
            $falhas++
            continue
        }

        try {
            Rename-Item -Path $caminhoOriginal -NewName $item.NomeSugerido -ErrorAction Stop
            Write-Host "Renomeado: $($item.NomeOriginal) -> $($item.NomeSugerido)"
            $renomeados++
        }
        catch {
            Write-Host "Erro ao renomear '$($item.NomeOriginal)': $($_.Exception.Message)"
            $falhas++
        }
    }

    Write-Host ""
    Write-Host "Resumo da renomeação"
    Write-Host "----------------------------------------"
    Write-Host "Renomeados com sucesso: $renomeados"
    Write-Host "Falhas                : $falhas"
}

# =========================
# Processar análise
# =========================
function Processar-Analise {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Arquivos,

        [Parameter(Mandatory = $true)]
        $Tabelas
    )

    if ($script:AplicarParalelismo -and $PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "Processando com paralelismo. ThrottleLimit: $script:ThrottleLimit"

        return @(
            $Arquivos | ForEach-Object -Parallel {
                function Ajustar-Revisao {
                    param ([Parameter(Mandatory = $true)][string]$NomeBase)

                    $novoNome = $NomeBase
                    $motivos = [System.Collections.Generic.List[string]]::new()
                    $alterado = $false

                    if ($novoNome -match '_RCA$') {
                        $novoNome = $novoNome -replace '_RCA$', '-RCA'
                        $motivos.Add('Separador "_" antes da revisão cancelada ajustado para "-"')
                        $alterado = $true
                    }

                    if ($novoNome -match '-RCA$') {
                        return [PSCustomObject]@{ NomeAjustado = $novoNome; Alterado = $alterado; Motivos = @($motivos) }
                    }

                    if ($novoNome -match '_R\d{2}[A-Za-z]?$') {
                        $novoNome = $novoNome -replace '_R', '-R'
                        $motivos.Add('Separador "_" antes da revisão ajustado para "-"')
                        $alterado = $true
                    }

                    if ($novoNome -match '-R\d{2}$') {
                        $novoNome = $novoNome + 'a'
                        $motivos.Add('Revisão sem letra final ajustada para "R00a"')
                        $alterado = $true
                    }

                    if ($novoNome -match '-R\d{2}[A-Z]$') {
                        $novoNome = [regex]::Replace($novoNome, '-R(\d{2})([A-Z])$', { param($m) "-R$($m.Groups[1].Value)$($m.Groups[2].Value.ToLower())" })
                        $motivos.Add('Letra final da revisão ajustada para minúscula')
                        $alterado = $true
                    }

                    return [PSCustomObject]@{ NomeAjustado = $novoNome; Alterado = $alterado; Motivos = @($motivos) }
                }

                function Novo-ResultadoAnalise {
                    param($NomeOriginal, $NomeSugerido, $CaminhoCompleto, $Diretorio, $Status, $Erros, $Avisos, $Motivos, $PrecisaAjuste, $Cancelado, $Campos)

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

                function Analisar-ArquivoANTTParallel {
                    param($Arquivo, $Tabelas)

                    $nomeArquivo = $Arquivo.Name
                    $caminhoCompleto = $Arquivo.FullName
                    $diretorio = $Arquivo.DirectoryName
                    $ext = $Arquivo.Extension
                    $nomeBaseOriginal = [System.IO.Path]::GetFileNameWithoutExtension($nomeArquivo).Trim()
                    $ajuste = Ajustar-Revisao -NomeBase $nomeBaseOriginal
                    $nomeBase = $ajuste.NomeAjustado
                    $nomeSugeridoCompleto = $nomeBase + $ext
                    $precisaAjusteReal = ($nomeArquivo -cne $nomeSugeridoCompleto)

                    $erros = [System.Collections.Generic.List[string]]::new()
                    $avisos = [System.Collections.Generic.List[string]]::new()
                    $status = 'Válido'
                    $cancelado = $false

                    $concessionaria = ''; $rodoviaUF = ''; $localizacao = ''; $tipoObra = ''; $tipoProjeto = ''; $classe = ''; $disciplina = ''; $sequencia = ''; $revisao = ''
                    $partes = $nomeBase -split '-'

                    if ($partes.Count -eq 9) {
                        $concessionaria = $partes[0].ToUpper(); $rodoviaUF = $partes[1].ToUpper(); $localizacao = $partes[2]; $tipoObra = $partes[3].ToUpper(); $tipoProjeto = $partes[4].ToUpper(); $classe = $partes[5].ToUpper(); $disciplina = $partes[6].ToUpper(); $sequencia = $partes[7]; $revisao = $partes[8]
                    }
                    elseif ($partes.Count -eq 10) {
                        $concessionaria = $partes[0].ToUpper(); $rodoviaUF = $partes[1].ToUpper(); $localizacao = "$($partes[2])-$($partes[3])"; $tipoObra = $partes[4].ToUpper(); $tipoProjeto = $partes[5].ToUpper(); $classe = $partes[6].ToUpper(); $disciplina = $partes[7].ToUpper(); $sequencia = $partes[8]; $revisao = $partes[9]
                    }
                    else {
                        return Novo-ResultadoAnalise $nomeArquivo $nomeSugeridoCompleto $caminhoCompleto $diretorio 'Inválido' @('Estrutura inválida. O nome deve seguir o padrão ANTT com 9 blocos.') @() $(if ($precisaAjusteReal) { $ajuste.Motivos } else { @() }) $precisaAjusteReal $false ([PSCustomObject]@{ Concessionaria=''; RodoviaUF=''; Localizacao=''; TipoObra=''; TipoProjeto=''; Classe=''; Disciplina=''; Sequencia=''; Revisao='' })
                    }

                    if ($concessionaria.Length -ne 3) { $erros.Add('Concessionária deve ter 3 caracteres.') }
                    if ($rodoviaUF.Length -ne 5) { $erros.Add('Rodovia/UF deve ter 5 caracteres.') }
                    if ($localizacao.Length -ne 7) { $erros.Add('Localização deve ter 7 caracteres.') }
                    if ($tipoObra.Length -ne 3) { $erros.Add('Tipo de obra deve ter 3 caracteres.') }
                    if ($tipoProjeto.Length -ne 3) { $erros.Add('Tipo de projeto deve ter 3 caracteres.') }
                    if ($classe.Length -ne 2) { $erros.Add('Classe deve ter 2 caracteres.') }
                    if ($disciplina.Length -ne 2) { $erros.Add('Disciplina deve ter 2 caracteres.') }
                    if ($sequencia.Length -ne 3) { $erros.Add('Sequência deve ter 3 caracteres.') }
                    if (-not $Tabelas.Concessionarias.Contains($concessionaria)) { $erros.Add('Concessionária inválida.') }
                    if ($rodoviaUF -notmatch '^\d{3}[A-Z]{2}$') { $erros.Add('Rodovia/UF inválida. Deve seguir o padrão 000UF.') }

                    if ($localizacao -notmatch '^\d{3}[-+]\d{3}$') {
                        $erros.Add('Localização inválida. Deve seguir 000-000 ou 000+000.')
                    }
                    elseif ($localizacao -match '^\d{3}-\d{3}$') {
                        $avisos.Add('Localização em trecho: indica extensão entre os quilômetros informados.')
                    }
                    elseif ($localizacao -match '^\d{3}\+\d{3}$') {
                        $avisos.Add('Localização pontual: indica ponto específico conforme os quilômetros informados.')
                    }

                    if (-not $Tabelas.TiposObra.Contains($tipoObra)) { $erros.Add('Tipo de obra inválido.') }
                    $tipoProjetoValido = $Tabelas.TiposProjetoFixos.Contains($tipoProjeto) -or ($tipoProjeto -match '^AL\d$')
                    if (-not $tipoProjetoValido) { $erros.Add('Tipo de projeto inválido.') }
                    if (-not $Tabelas.Classes.Contains($classe)) { $erros.Add('Classe inválida.') }
                    if (-not $Tabelas.Disciplinas.Contains($disciplina)) { $erros.Add('Disciplina inválida.') }
                    if ($sequencia -notmatch '^\d{3}$') { $erros.Add('Sequência inválida. Deve conter 3 dígitos numéricos.') }

                    if ($revisao -eq 'RCA') {
                        $cancelado = $true; $status = 'Cancelado'; $avisos.Add('Arquivo cancelado. Não deve ser alterado.')
                    }
                    elseif ($revisao -notmatch '^R\d{2}[a-z]$') {
                        $erros.Add('Revisão inválida. O padrão esperado é como R00a, com apenas a última letra minúscula.')
                    }

                    if ($erros.Count -gt 0) { $status = 'Inválido' }
                    elseif ($precisaAjusteReal -and -not $cancelado) { $status = 'Ajustável' }
                    elseif ($cancelado) { $status = 'Cancelado' }
                    else { $status = 'Válido' }

                    return Novo-ResultadoAnalise $nomeArquivo $nomeSugeridoCompleto $caminhoCompleto $diretorio $status @($erros) @($avisos) $(if ($precisaAjusteReal) { $ajuste.Motivos } else { @() }) $precisaAjusteReal $cancelado ([PSCustomObject]@{ Concessionaria=$concessionaria; RodoviaUF=$rodoviaUF; Localizacao=$localizacao; TipoObra=$tipoObra; TipoProjeto=$tipoProjeto; Classe=$classe; Disciplina=$disciplina; Sequencia=$sequencia; Revisao=$revisao })
                }

                Analisar-ArquivoANTTParallel -Arquivo $_ -Tabelas $using:Tabelas
            } -ThrottleLimit $script:ThrottleLimit
        )
    }

    $resultados = [System.Collections.Generic.List[object]]::new()

    foreach ($arquivo in $Arquivos) {
        $resultado = Analisar-ArquivoANTT -Arquivo $arquivo -Tabelas $Tabelas
        $resultados.Add($resultado)
    }

    return @($resultados)
}

# =========================
# Execução
# =========================
try {
    $pastaSelecionada = Selecionar-Pasta

    if (-not $pastaSelecionada) {
        Write-Host "Nenhuma pasta válida foi selecionada ou informada."
        Read-Host "Pressione Enter para sair"
        exit
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
    Write-Host "Filtro de extensão ativo: $script:UsarFiltroExtensao"

    if ($script:UsarFiltroExtensao) {
        Write-Host "Extensões consideradas: $($script:ExtensoesPermitidas -join ', ')"
    }

    Write-Host "Quantidade de arquivos encontrados: $($arquivos.Count)"
    Write-Host "----------------------------------------"

    if ($arquivos.Count -eq 0) {
        Write-Host "Nenhum arquivo encontrado para análise."
        Read-Host "Pressione Enter para encerrar"
        exit
    }

    do {
        $confirmacao = Read-Host "Deseja continuar para a validação e análise de ajuste? (S/N)"
        $confirmacao = $confirmacao.ToUpper().Trim()
    } while ($confirmacao -notin @('S', 'N'))

    if ($confirmacao -eq 'N') {
        Write-Host "Processo cancelado pelo usuário."
        Read-Host "Pressione Enter para encerrar"
        exit
    }

    $tabelas = Obter-TabelasANTT
    $tabelasHash = Converter-TabelasParaHashSet -Tabelas $tabelas

    $resultados = Medir-Tempo 'Análise dos arquivos' {
        Processar-Analise -Arquivos $arquivos -Tabelas $tabelasHash
    }

    if ($script:ExibirDetalhadoNoConsole) {
        foreach ($resultado in $resultados) {
            Exibir-Resultado -Resultado $resultado -PastaRaiz $pastaSelecionada
        }
    }
    else {
        Write-Host "Exibição detalhada no console desativada para melhorar performance."
        Write-Host "Use o CSV gerado para analisar arquivo por arquivo."
        Write-Host ""
    }

    Exibir-Resumo -Resultados $resultados

    if ($script:ExportarCsv) {
        Exportar-ResultadoCsv -Resultados $resultados -PastaRaiz $pastaSelecionada | Out-Null
    }

    Renomear-ArquivosAjustaveis -Resultados $resultados
}
catch {
    Write-Host "Ocorreu um erro durante a execução do script."
    Write-Host $_.Exception.Message
}
finally {
    Read-Host "Pressione Enter para encerrar"
}
