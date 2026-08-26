# Script ProjectWise otimizado
# Objetivo: reduzir chamadas repetidas ao ProjectWise usando cache e buscas mais específicas.

#-------------------------------------------------------
# Caches globais da execução
#-------------------------------------------------------
$script:CacheDisciplina = @{}
$script:CacheFaseProjeto = @{}
$script:CachePastas = @{}
$script:CacheRichProject = @{}
$script:CacheDocsPorPastaNumero = @{}

#-------------------------------------------------------
# Utilitários
#-------------------------------------------------------
function Medir-Tempo {
    param(
        [string]$Nome,
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

function Get-Attr0 {
    param($documento)

    if ($null -eq $documento) { return $null }
    if ($null -eq $documento.Attributes) { return $null }
    if ($null -eq $documento.Attributes[0]) { return $null }

    return $documento.Attributes[0]
}

function TestaPastaExisteCache {
    param([string]$pastaDestino)

    if ([string]::IsNullOrWhiteSpace($pastaDestino)) { return $false }

    $key = $pastaDestino.TrimEnd('\').ToLowerInvariant()

    if ($script:CachePastas.ContainsKey($key)) {
        return $script:CachePastas[$key]
    }

    $existe = [bool](Get-PWFolders -FolderPath $pastaDestino -WarningAction SilentlyContinue)
    $script:CachePastas[$key] = $existe

    return $existe
}

function ObterRichProjectCache {
    param($documento)

    if ($null -eq $documento) { return $null }

    $documentId = $null
    foreach ($propName in @('DocumentGUID', 'DocumentId', 'DocumentID', 'ObjectId', 'ObjectID')) {
        $prop = $documento.PSObject.Properties[$propName]
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            $documentId = "$propName|$($prop.Value)"
            break
        }
    }

    $key = if ($documentId) { $documentId } else { "FolderPath|$($documento.FolderPath)|Name|$($documento.Name)|FileName|$($documento.FileName)" }

    if ($script:CacheRichProject.ContainsKey($key)) {
        return $script:CacheRichProject[$key]
    }

    $result = Get-PWRichProjectForDocument -InputDocument $documento
    $script:CacheRichProject[$key] = $result

    return $result
}

function ConverterParaInteiroSeguro {
    param($valor)

    $saida = 0
    if ([int]::TryParse([string]$valor, [ref]$saida)) {
        return $saida
    }

    return $null
}

function Set-StateSeguro {
    param(
        $Documentos,
        [string]$State,
        [switch]$SilenciarErroManualChange
    )

    if (-not $Documentos -or [string]::IsNullOrWhiteSpace($State)) { return }

    try {
        Set-PWDocumentState -InputDocuments $Documentos -State $State -Force -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        $mensagem = $_.Exception.Message

        if ($SilenciarErroManualChange -and $mensagem -match 'Manual change state not allowed') {
            return
        }

        Write-Warning ("Falha ao alterar state para '{0}': {1}" -f $State, $mensagem)
    }
}

#-------------------------------------------------------
# Busca inicial
#-------------------------------------------------------
function ObterDocumentosValidados {
    Get-PWDocumentsBySearch -SearchName 'Scripts\Docs GRD Validados' -GetAttributes
}

#-------------------------------------------------------
# Regras de cadastro
#-------------------------------------------------------
function DefineDisciplinaPai {
    param ($cadastroDisciplina, $poderConcedente)

    $attr = Get-Attr0 $cadastroDisciplina
    if ($null -eq $attr -or [string]::IsNullOrWhiteSpace($attr.Relacionamento)) { return $null }

    if ($poderConcedente -eq "ARTESP") {
        $partes = $attr.Relacionamento.Split(';')
        if ($partes.Count -lt 2) { return $null }
        return $partes[1]
    }

    return $attr.Relacionamento
}

function DefineFaseProjeto {
    param ($cadastroDisciplina, $cadastroFaseProjeto, $poderConcedente)

    $attrDisciplina = Get-Attr0 $cadastroDisciplina
    if ($null -eq $attrDisciplina -or [string]::IsNullOrWhiteSpace($attrDisciplina.Relacionamento)) { return $null }

    if ($poderConcedente -eq "ARTESP") {
        $partes = $attrDisciplina.Relacionamento.Split(';')
        if ($partes.Count -lt 1) { return $null }
        return $partes[0]
    }

    $attrFase = Get-Attr0 $cadastroFaseProjeto
    if ($null -eq $attrFase -or [string]::IsNullOrWhiteSpace($attrFase.Relacionamento)) { return $null }

    return $attrFase.Relacionamento
}

function ObtemCadastroDisciplina {
    param ($documentoBase, $poderConcedente)

    $attr = Get-Attr0 $documentoBase
    if ($null -eq $attr -or [string]::IsNullOrWhiteSpace($attr.Disciplina)) { return $null }

    $disciplina = $attr.Disciplina
    $tipoRegistro = if ($poderConcedente -eq "ARTESP") { 'Disciplinas ARTESP' } else { 'Disciplinas ANTT' }
    $key = "$tipoRegistro|$disciplina"

    if ($script:CacheDisciplina.ContainsKey($key)) {
        return $script:CacheDisciplina[$key]
    }

    $result = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{
        TipoRegistro = $tipoRegistro
        Codigo       = $disciplina
    } -GetAttributes

    $script:CacheDisciplina[$key] = $result
    return $result
}

function ObtemCadastroFaseProjeto {
    param ($documentoBase, $poderConcedente)

    if ($poderConcedente -eq "ARTESP") { return $null }

    $attr = Get-Attr0 $documentoBase
    if ($null -eq $attr -or [string]::IsNullOrWhiteSpace($attr.FaseProjeto)) { return $null }

    $faseProjeto = $attr.FaseProjeto
    $key = "Tipos de Projeto ANTT|$faseProjeto"

    if ($script:CacheFaseProjeto.ContainsKey($key)) {
        return $script:CacheFaseProjeto[$key]
    }

    $result = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{
        TipoRegistro = 'Tipos de Projeto ANTT'
        Codigo       = $faseProjeto
    } -GetAttributes

    $script:CacheFaseProjeto[$key] = $result
    return $result
}

function TemAtributosParaRoteamento {
    param ($doc, $poderConcedente)

    $a = Get-Attr0 $doc
    if ($null -eq $a) { return $false }

    if ($poderConcedente -eq 'ARTESP') {
        return [bool]$a.Disciplina
    }

    return ([bool]$a.Disciplina -and [bool]$a.FaseProjeto)
}

function ValidaSeModeloFederadoAutoral {
    param($documento)

    $attr = Get-Attr0 $documento
    if ($null -eq $attr) { return $false }

    $sequencial = ConverterParaInteiroSeguro $attr.Sequencial

    if ($attr.PoderConcedente -eq 'ANTT' -and $null -ne $sequencial -and $sequencial -in 800..999) { return $true }
    if ($attr.PoderConcedente -eq 'ARTESP' -and $attr.TipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

#-------------------------------------------------------
# Relacionamentos / hospedeiro
#-------------------------------------------------------
function TentarEncontrarHospedeiro {
    param ($documento)

    $rels = $null

    try {
        $rels = Get-PWDocumentRelationships -InputDocuments $documento -ReferencedBy -ErrorAction Stop
    }
    catch {
        try {
            $rels = Get-PWDocumentReferences -InputDocuments $documento -ReferencedBy -ErrorAction Stop
        }
        catch {
            $rels = $null
        }
    }

    if (-not $rels) { return $null }

    $hospedeiros = $rels | Where-Object {
        $_.FileName -match '\.(dgn|dwg)$'
    }

    if ($hospedeiros) { return ($hospedeiros | Select-Object -First 1) }
    return ($rels | Select-Object -First 1)
}

#-------------------------------------------------------
# Cálculo de pasta destino
#-------------------------------------------------------
function CalculaPastaDestino {
    param ($documento, [ref]$documentoBaseUsado)

    $attrDocumento = Get-Attr0 $documento
    if ($null -eq $attrDocumento) { return $null }

    $poderConcedente = $attrDocumento.PoderConcedente
    if ([string]::IsNullOrWhiteSpace($poderConcedente)) { return $null }

    $docBase = $documento

    if (-not (TemAtributosParaRoteamento -doc $documento -poderConcedente $poderConcedente)) {
        $hosp = TentarEncontrarHospedeiro -documento $documento
        if ($hosp -and (TemAtributosParaRoteamento -doc $hosp -poderConcedente $poderConcedente)) {
            $docBase = $hosp
        }
        else {
            return $null
        }
    }

    $documentoBaseUsado.Value = $docBase
    $attrBase = Get-Attr0 $docBase
    if ($null -eq $attrBase) { return $null }

    $cadastroDisciplina = ObtemCadastroDisciplina -documentoBase $docBase -poderConcedente $poderConcedente
    $disciplina = DefineDisciplinaPai -cadastroDisciplina $cadastroDisciplina -poderConcedente $poderConcedente
    if ([string]::IsNullOrWhiteSpace($disciplina)) { return $null }

    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -documentoBase $docBase -poderConcedente $poderConcedente
    $faseProjeto = DefineFaseProjeto -cadastroDisciplina $cadastroDisciplina -cadastroFaseProjeto $cadastroFaseProjeto -poderConcedente $poderConcedente
    if ([string]::IsNullOrWhiteSpace($faseProjeto)) { return $null }

    $volume = $attrBase.Volume

    $pastaRaizProjeto = ObterRichProjectCache -documento $docBase
    if (-not $pastaRaizProjeto) { return $null }

    $caminhoEspecialModeloAutoral = ''

    if ($attrDocumento.Disciplina -ne 'U4' -and (ValidaSeModeloFederadoAutoral $documento)) {
        $caminhoEspecialModeloAutoral = 'Modelo BIM\Modelos Autorais'
    }

    if ($attrDocumento.PoderConcedente -eq 'ARTESP' -and $attrDocumento.TipoDocumento -eq 'MI') {
        $caminhoEspecialModeloAutoral = 'Modelo BIM'
        $disciplina = ''
    }

    $pastaDestino = @(
        $pastaRaizProjeto.FullPath,
        '1 - Area de Trabalho',
        $faseProjeto,
        $volume,
        $caminhoEspecialModeloAutoral,
        $disciplina
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $pastaDestino = $pastaDestino -join '\'

    if (TestaPastaExisteCache -pastaDestino $pastaDestino) {
        return $pastaDestino
    }

    $pastaRaizProjetoSemCache = Get-PWRichProjectForDocument -InputDocument $docBase
    if ($pastaRaizProjetoSemCache -and $pastaRaizProjetoSemCache.FullPath -ne $pastaRaizProjeto.FullPath) {
        $pastaDestinoSemCache = @(
            $pastaRaizProjetoSemCache.FullPath,
            '1 - Area de Trabalho',
            $faseProjeto,
            $volume,
            $caminhoEspecialModeloAutoral,
            $disciplina
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $pastaDestinoSemCache = $pastaDestinoSemCache -join '\'

        if ([bool](Get-PWFolders -FolderPath $pastaDestinoSemCache -WarningAction SilentlyContinue)) {
            return $pastaDestinoSemCache
        }
    }

    Write-Warning "Pasta destino calculada não encontrada: $pastaDestino"

    return $null
}

#-------------------------------------------------------
# Busca de documentos existentes / anteriores
#-------------------------------------------------------
function ObtemDocumentoExistenteNaPastaDestino {
    param ($pastaDestino, $documento)

    if ([string]::IsNullOrWhiteSpace($pastaDestino) -or $null -eq $documento) { return $null }

    $resultados = @()

    if (-not [string]::IsNullOrWhiteSpace($documento.FileName)) {
        $resultados += Get-PWDocumentsBySearch `
            -FolderPath $pastaDestino `
            -JustThisFolder `
            -FileName $documento.FileName `
            -GetAttributes `
            -WarningAction SilentlyContinue
    }

    if (-not $resultados -and -not [string]::IsNullOrWhiteSpace($documento.Name)) {
        $resultados += Get-PWDocumentsBySearch `
            -FolderPath $pastaDestino `
            -JustThisFolder `
            -DocumentName $documento.Name `
            -GetAttributes `
            -WarningAction SilentlyContinue
    }

    if (-not $resultados) { return $null }

    return $resultados | Where-Object {
        $_.Name -eq $documento.Name -or $_.FileName -eq $documento.FileName
    } | Select-Object -First 1
}

function ObtemDocumentosAnterioresNaPastaDestino {
    param ($pastaDestino, $numeroDocumento, $sequencialEmissao)

    if ([string]::IsNullOrWhiteSpace($pastaDestino)) { return $null }
    if ([string]::IsNullOrWhiteSpace($numeroDocumento)) { return $null }

    $seqAtual = ConverterParaInteiroSeguro $sequencialEmissao
    if ($null -eq $seqAtual) { return $null }

    $key = "$($pastaDestino.TrimEnd('\').ToLowerInvariant())|$numeroDocumento"

    if ($script:CacheDocsPorPastaNumero.ContainsKey($key)) {
        $documentos = $script:CacheDocsPorPastaNumero[$key]
    }
    else {
        $documentos = Get-PWDocumentsBySearch `
            -FolderPath $pastaDestino `
            -JustThisFolder `
            -Attributes @{ NumeroPoderConcedente = $numeroDocumento } `
            -GetAttributes `
            -WarningAction SilentlyContinue

        $script:CacheDocsPorPastaNumero[$key] = $documentos
    }

    if (-not $documentos) { return $null }

    return $documentos | Where-Object {
        $seqDoc = ConverterParaInteiroSeguro $_.Attributes[0].SequencialEmissao
        $null -ne $seqDoc -and $seqDoc -lt $seqAtual
    }
}

#-------------------------------------------------------
# Movimentação de superados / atualização de estados
#-------------------------------------------------------
function MovimentaDocumentosAnterioresParaPastaSuperado {
    param ($pastaDestino, $documento)

    $attr = Get-Attr0 $documento
    if ($null -eq $attr) { return }

    $numeroDocumento = $attr.NumeroPoderConcedente
    $sequencialEmissao = $attr.SequencialEmissao

    if ([string]::IsNullOrWhiteSpace($numeroDocumento) -or [string]::IsNullOrWhiteSpace($sequencialEmissao)) { return }

    $documentosAnteriores = ObtemDocumentosAnterioresNaPastaDestino `
        -pastaDestino $pastaDestino `
        -numeroDocumento $numeroDocumento `
        -sequencialEmissao $sequencialEmissao

    if (-not $documentosAnteriores) { return }

    $pastaSuperados = $pastaDestino + '\Superados'

    if (-not (TestaPastaExisteCache -pastaDestino $pastaSuperados)) {
        Write-Warning "Pasta Superados não encontrada: $pastaSuperados"
        return
    }

    $documentosMovimentados = Move-PWDocumentsToFolder `
        -InputDocument $documentosAnteriores `
        -TargetFolderPath $pastaSuperados

    if ($documentosMovimentados) {
        Set-StateSeguro -Documentos $documentosMovimentados -State 'Superado' -SilenciarErroManualChange
    }
}

function CalculaState {
    param ($stateAtual)

    switch ($stateAtual) {
        'Emitido pela Engenharia' { 'Nova emissao sendo analisada pela Engenharia' }
        'Solicitado reanalise da Engenharia' { 'Nova emissao sendo analisada pela Engenharia' }
        'Enviado ao Poder Concedente' { 'Enviado ao Poder Concedente - Nova emissao em analise Eng' }
        'Concluido pela Unidade' { 'Concluido - Nova emissao em analise Eng' }
        Default { $null }
    }
}

function AtualizaStateDocumentosEmitidosParaUnidade {
    param ($documento, $pastaDestino)

    $attr = Get-Attr0 $documento
    if ($null -eq $attr) { return }

    $pastaUnidade = $pastaDestino.Replace('1 - Area de Trabalho', '2 - Unidade')
    $numeroDocumento = $attr.NumeroPoderConcedente
    $seqDocumentoAtual = ConverterParaInteiroSeguro $attr.SequencialEmissao

    if ([string]::IsNullOrWhiteSpace($numeroDocumento) -or $null -eq $seqDocumentoAtual) { return }

    $documentosUnidade = Get-PWDocumentsBySearch `
        -FolderPath $pastaUnidade `
        -JustThisFolder `
        -Attributes @{ NumeroPoderConcedente = $numeroDocumento } `
        -GetAttributes `
        -WarningAction SilentlyContinue

    if (-not $documentosUnidade) { return }

    $documentoUnidade = $documentosUnidade | Select-Object -First 1
    $attrUnidade = Get-Attr0 $documentoUnidade
    if ($null -eq $attrUnidade) { return }

    $seqUnidade = ConverterParaInteiroSeguro $attrUnidade.SequencialEmissao
    if ($null -eq $seqUnidade -or $seqUnidade -ge $seqDocumentoAtual) { return }

    $stateAtual = $documentoUnidade.WorkflowState
    $proximoState = CalculaState -stateAtual $stateAtual

    if ($proximoState) {
        Set-StateSeguro -Documentos $documentosUnidade -State $proximoState -SilenciarErroManualChange
    }
}

#-------------------------------------------------------
# Cópia alternativa sem arquivo físico
#-------------------------------------------------------
function Copiar-DocumentoSemArquivo {
    param($documento, [string]$pastaDestino)

    $dup = ObtemDocumentoExistenteNaPastaDestino -pastaDestino $pastaDestino -documento $documento
    if ($dup) { return $dup }

    if ([string]::IsNullOrWhiteSpace($documento.FileName)) {
        Write-Warning "Documento sem FileName. Não foi possível criar arquivo temporário para: $($documento.Name)"
        return $null
    }

    $temp = Join-Path $env:TEMP ("PWEmpty_" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $temp | Out-Null
    $tempFile = Join-Path $temp $documento.FileName
    New-Item -Path $tempFile -ItemType File | Out-Null

    try {
        $novo = $null
        $newCmd = Get-Command -Name New-PWDocument -ErrorAction SilentlyContinue

        if ($newCmd -and $newCmd.Parameters.ContainsKey('FilePath') -and $newCmd.Parameters.ContainsKey('FolderPath')) {
            $params = @{
                FolderPath = $pastaDestino
                FilePath   = $tempFile
            }

            if ($newCmd.Parameters.ContainsKey('Description') -and $documento.Description) { $params['Description'] = $documento.Description }
            if ($newCmd.Parameters.ContainsKey('Name')) { $params['Name'] = $documento.Name }

            $novo = New-PWDocument @params -ErrorAction Stop
        }
        else {
            $importFromFolder = Get-Command -Name Import-PWDocumentsFromFolder -ErrorAction SilentlyContinue

            if ($importFromFolder) {
                $novo = Import-PWDocumentsFromFolder `
                    -InputFolder $temp `
                    -ProjectWiseFolder $pastaDestino `
                    -JustOneLevel `
                    -ErrorAction Stop `
                    -WarningAction SilentlyContinue | Select-Object -First 1
            }
            else {
                $importPlanilha = Get-Command -Name Import-PWDocuments -ErrorAction SilentlyContinue

                if ($importPlanilha) {
                    $novo = Import-PWDocuments `
                        -InputFolder $temp `
                        -FolderPath $pastaDestino `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue | Select-Object -First 1
                }
                else {
                    throw 'Nenhum cmdlet disponível para importar/criar documento com arquivo.'
                }
            }
        }

        if (-not $novo) {
            $novo = Get-PWDocumentsBySearch `
                -FolderPath $pastaDestino `
                -JustThisFolder `
                -FileName $documento.FileName `
                -WarningAction SilentlyContinue | Select-Object -First 1
        }

        return $novo
    }
    finally {
        try { Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

#-------------------------------------------------------
# Cópia principal com fallback
#-------------------------------------------------------
function CopiarDocumentoParaDestino {
    param($documento, [string]$pastaDestino)

    $documentoCopiado = $null

    try {
        $documentoCopiado = Copy-PWDocumentsToFolder `
            -InputDocument $documento `
            -TargetFolderPath $pastaDestino `
            -ErrorAction Stop `
            -WarningAction Stop
    }
    catch {
        Write-Warning ("Copy-PWDocumentsToFolder falhou para {0}: {1}" -f $documento.FileName, $_.Exception.Message)
    }

    if ($documentoCopiado) { return $documentoCopiado }

    if ([string]::IsNullOrWhiteSpace($documento.FileName)) {
        return Copiar-DocumentoSemArquivo -documento $documento -pastaDestino $pastaDestino
    }

    $temp = Join-Path $env:TEMP ("PWCopy_" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $temp | Out-Null
    $caminhoArquivo = Join-Path $temp $documento.FileName
    $exportOk = $false

    try {
        try {
            $null = Export-PWDocuments `
                -InputDocuments $documento `
                -OutputFolder $temp `
                -ErrorAction SilentlyContinue `
                -WarningAction SilentlyContinue

            $exportOk = Test-Path $caminhoArquivo
        }
        catch {
            $exportOk = $false
        }

        if ($exportOk) {
            try {
                $importFromFolder = Get-Command -Name Import-PWDocumentsFromFolder -ErrorAction SilentlyContinue

                if ($importFromFolder) {
                    $documentoCopiado = Import-PWDocumentsFromFolder `
                        -InputFolder $temp `
                        -ProjectWiseFolder $pastaDestino `
                        -JustOneLevel `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue
                }
                else {
                    $importPlanilha = Get-Command -Name Import-PWDocuments -ErrorAction SilentlyContinue

                    if ($importPlanilha) {
                        $documentoCopiado = Import-PWDocuments `
                            -InputFolder $temp `
                            -FolderPath $pastaDestino `
                            -ErrorAction Stop `
                            -WarningAction SilentlyContinue
                    }
                    else {
                        $documentoCopiado = New-PWDocument `
                            -FolderPath $pastaDestino `
                            -Name $documento.Name `
                            -FilePath $caminhoArquivo `
                            -ErrorAction Stop
                    }
                }
            }
            catch {
                Write-Warning ("Falha ao importar/criar a partir do arquivo exportado: " + $_.Exception.Message)
            }

            if ($documentoCopiado -is [System.Array]) {
                $documentoCopiado = $documentoCopiado | Where-Object {
                    $_.FolderPath -eq $pastaDestino -and $_.FileName -eq $documento.FileName
                } | Select-Object -First 1
            }

            if (-not $documentoCopiado) {
                $documentoCopiado = Get-PWDocumentsBySearch `
                    -FolderPath $pastaDestino `
                    -JustThisFolder `
                    -FileName $documento.FileName `
                    -WarningAction SilentlyContinue

                if ($documentoCopiado -is [System.Array]) {
                    $documentoCopiado = $documentoCopiado | Select-Object -First 1
                }
            }
        }
        else {
            $documentoCopiado = Copiar-DocumentoSemArquivo -documento $documento -pastaDestino $pastaDestino
        }

        return $documentoCopiado
    }
    finally {
        try { Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

#-------------------------------------------------------
# Execução principal
#-------------------------------------------------------
try {
    if ([string]::IsNullOrWhiteSpace($env:ECORODOVIAS_PW_PASSWORD)) { throw 'Defina ECORODOVIAS_PW_PASSWORD para executar esta rotina automática.' }
    $SecurePassword = ConvertTo-SecureString $env:ECORODOVIAS_PW_PASSWORD -AsPlainText -Force

    New-PWLogin `
        -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' `
        -Password $SecurePassword `
        -UserName 'admin'

    $documentos = Medir-Tempo 'Busca documentos validados' {
        ObterDocumentosValidados
    }

    if (-not $documentos) {
        Write-Host 'Nenhum documento validado encontrado.'
        return
    }

    $documentos = @($documentos)
    $totalDocumentos = $documentos.Count
    $indiceDocumento = 0

    Write-Host ("Quantidade de documentos a movimentar: {0}" -f $totalDocumentos)

    foreach ($documento in $documentos) {
        $indiceDocumento++
        $nomeLog = if ($documento.FileName) { $documento.FileName } else { $documento.Name }
        $percentual = [math]::Round(($indiceDocumento / $totalDocumentos) * 100, 0)

        Write-Progress `
            -Activity 'Movimentando documentos validados' `
            -Status ("{0}/{1} - {2}" -f $indiceDocumento, $totalDocumentos, $nomeLog) `
            -PercentComplete $percentual

        Write-Host ("Processando [{0}/{1} - {2}%]: {3}" -f $indiceDocumento, $totalDocumentos, $percentual, $nomeLog)

        $docBaseRef = New-Object PSObject

        $pastaDestino = Medir-Tempo "Calcula pasta destino - $nomeLog" {
            CalculaPastaDestino -documento $documento -documentoBaseUsado ([ref]$docBaseRef)
        }

        if (-not $pastaDestino) {
            Set-StateSeguro `
                -Documentos $documento `
                -State 'Validado pelo Sistema - Pasta destino nao localizada' `
                -SilenciarErroManualChange
            continue
        }

        $srcFolder = ($documento.FolderPath.TrimEnd('\')).ToLowerInvariant()
        $dstFolder = ($pastaDestino.TrimEnd('\')).ToLowerInvariant()

        if ($srcFolder -eq $dstFolder) {
            Set-StateSeguro `
                -Documentos $documento `
                -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
                -SilenciarErroManualChange

            AtualizaStateDocumentosEmitidosParaUnidade -documento $documento -pastaDestino $pastaDestino
            continue
        }

        # Melhoria Modelos BIM
        if ((ValidaSeModeloFederadoAutoral -documento $documento)) {

            $docExiste = ObtemDocumentoExistenteNaPastaDestino -pastaDestino $pastaDestino -documento $documento
            $criarVersao = $null -ne $docExiste

            $documentoCopiado = Copy-PWDocumentsToFolder -InputDocument $documento -TargetFolderPath $pastaDestino -CreateVersion:$criarVersao

            if (-not $documentoCopiado) {
                Write-Warning ("Cópia não concluída para: " + $nomeLog)
                continue
            }

            $stateDocumentoCopiado = 'Em analise do Assistente'

            $attrDocCopiado = Get-Attr0 $documentoCopiado
            if ($attrDocCopiado -and $attrDocCopiado.PoderConcedente -eq 'ARTESP') {
                $stateDocumentoCopiado = 'Em analise da Engenharia'
            }

            Set-StateSeguro -Documentos $documentoCopiado -State $stateDocumentoCopiado
            Set-StateSeguro -Documentos $documento -State 'Validado pelo Sistema - Copiado p/ Disciplina' -SilenciarErroManualChange

            AtualizaStateDocumentosEmitidosParaUnidade -documento $documento -pastaDestino $pastaDestino
            continue
        }

        Medir-Tempo "Move anteriores para Superados - $nomeLog" {
            MovimentaDocumentosAnterioresParaPastaSuperado -pastaDestino $pastaDestino -documento $documento
        }

        $jaExiste = Medir-Tempo "Verifica duplicidade - $nomeLog" {
            ObtemDocumentoExistenteNaPastaDestino -pastaDestino $pastaDestino -documento $documento
        }

        if ($jaExiste) {
            Set-StateSeguro `
                -Documentos $documento `
                -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
                -SilenciarErroManualChange

            AtualizaStateDocumentosEmitidosParaUnidade -documento $documento -pastaDestino $pastaDestino
            continue
        }

        $documentoCopiado = Medir-Tempo "Copia documento - $nomeLog" {
            CopiarDocumentoParaDestino -documento $documento -pastaDestino $pastaDestino
        }

        if (-not $documentoCopiado) {
            Write-Warning ("Cópia não concluída para: " + $nomeLog)
            continue
        }

        $stateDocumentoCopiado = 'Em analise do Assistente'

        $attrDocCopiado = Get-Attr0 $documentoCopiado
        if ($attrDocCopiado -and $attrDocCopiado.PoderConcedente -eq 'ARTESP') {
            $stateDocumentoCopiado = 'Em analise da Engenharia'
        }

        Set-StateSeguro `
            -Documentos $documentoCopiado `
            -State $stateDocumentoCopiado

        Set-StateSeguro `
            -Documentos $documento `
            -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
            -SilenciarErroManualChange

        AtualizaStateDocumentosEmitidosParaUnidade -documento $documento -pastaDestino $pastaDestino
    }
}
finally {
    Undo-PWLogin
}
