<#
.SYNOPSIS
    Movimenta documentos validados para a pasta de disciplina correspondente no ProjectWise.

.MELHORIAS APLICADAS
    - Contagem total de documentos encontrados, processados, copiados, ignorados e com erro.
    - Barra de progresso durante a movimentacao.
    - Cache para reduzir chamadas repetidas ao ProjectWise.
    - Logs de erro com mensagem clara.
    - Tratamento mais seguro para arrays retornados pelos cmdlets.
    - Busca de duplicidade mais objetiva quando possivel.

.OBSERVACAO DE SEGURANCA
    Evite manter senha fixa no script. O ideal e usar Get-Credential, cofre de credenciais
    ou execucao com usuario de servico controlado.
#>

#-------------------------------------------------------
# Configuracoes gerais
#-------------------------------------------------------
$ErrorActionPreference = 'Stop'

$DatasourceName = '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01'
$UserName       = 'admin'

# IMPORTANTE: por seguranca, prefira Get-Credential ou um cofre de credenciais.
# Mantido abaixo para preservar o comportamento original do script.
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force

#-------------------------------------------------------
# Caches em memoria para reduzir chamadas repetidas
#-------------------------------------------------------
$script:CacheCadastroDisciplina = @{}
$script:CacheCadastroFaseProjeto = @{}
$script:CachePastas = @{}
$script:CacheCmdlets = @{}

#-------------------------------------------------------
# Contadores globais
#-------------------------------------------------------
$script:Contador = [ordered]@{
    TotalEncontrado              = 0
    Processados                  = 0
    Copiados                     = 0
    JaEstavamNoDestino           = 0
    DuplicadosNoDestino          = 0
    PastaDestinoNaoLocalizada    = 0
    MovidosParaSuperados         = 0
    UnidadeAtualizada            = 0
    FalhaCopia                   = 0
    Erros                        = 0
}

#-------------------------------------------------------
# Funcoes auxiliares
#-------------------------------------------------------
function Get-ValorAtributo {
    param(
        [Parameter(Mandatory = $false)] $Documento,
        [Parameter(Mandatory = $true)]  [string] $Nome
    )

    if ($null -eq $Documento -or $null -eq $Documento.Attributes -or $null -eq $Documento.Attributes[0]) {
        return $null
    }

    return $Documento.Attributes[0].$Nome
}

function Get-NomeDocumentoParaExibicao {
    param($Documento)

    if ($Documento.FileName) { return $Documento.FileName }
    if ($Documento.Name)     { return $Documento.Name }
    return '<sem nome>'
}

function Get-CmdletCached {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not $script:CacheCmdlets.ContainsKey($Name)) {
        $script:CacheCmdlets[$Name] = Get-Command -Name $Name -ErrorAction SilentlyContinue
    }

    return $script:CacheCmdlets[$Name]
}

function Get-PWFoldersCached {
    param([Parameter(Mandatory = $true)][string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $null }

    $key = $FolderPath.TrimEnd('\').ToLowerInvariant()

    if (-not $script:CachePastas.ContainsKey($key)) {
        try {
            $script:CachePastas[$key] = Get-PWFolders -FolderPath $FolderPath -ErrorAction Stop
        }
        catch {
            $script:CachePastas[$key] = $null
        }
    }

    return $script:CachePastas[$key]
}

function ObterDocumentosValidados {
    @(Get-PWDocumentsBySearch -SearchName 'Scripts\Docs GRD Validados' -GetAttributes)
}

function DefineDisciplinaPai {
    param($CadastroDisciplina, $PoderConcedente)

    $relacionamento = Get-ValorAtributo -Documento $CadastroDisciplina -Nome 'Relacionamento'
    if (-not $relacionamento) { return $null }

    if ($PoderConcedente -eq 'ARTESP') {
        $partes = $relacionamento.Split(';')
        if ($partes.Count -ge 2) { return $partes[1] }
        return $null
    }

    return $relacionamento
}

function DefineFaseProjeto {
    param($CadastroDisciplina, $CadastroFaseProjeto, $PoderConcedente)

    $relacionamentoDisciplina = Get-ValorAtributo -Documento $CadastroDisciplina -Nome 'Relacionamento'
    if (-not $relacionamentoDisciplina) { return $null }

    if ($PoderConcedente -eq 'ARTESP') {
        $partes = $relacionamentoDisciplina.Split(';')
        if ($partes.Count -ge 1) { return $partes[0] }
        return $null
    }

    $relacionamentoFase = Get-ValorAtributo -Documento $CadastroFaseProjeto -Nome 'Relacionamento'
    if (-not $relacionamentoFase) { return $null }

    return $relacionamentoFase
}

function ObtemCadastroDisciplina {
    param($DocumentoBase, $PoderConcedente)

    $disciplina = Get-ValorAtributo -Documento $DocumentoBase -Nome 'Disciplina'
    if (-not $disciplina) { return $null }

    $tipoRegistro = switch ($PoderConcedente) {
        'ARTESP' { 'Disciplinas ARTESP' }
        Default  { 'Disciplinas ANTT' }
    }

    $cacheKey = "$tipoRegistro|$disciplina"
    if (-not $script:CacheCadastroDisciplina.ContainsKey($cacheKey)) {
        $script:CacheCadastroDisciplina[$cacheKey] = Get-PWDocumentsBySearch `
            -Environment 'dmsRegistro' `
            -Attributes @{ TipoRegistro = $tipoRegistro; Codigo = $disciplina } `
            -GetAttributes
    }

    return $script:CacheCadastroDisciplina[$cacheKey]
}

function ObtemCadastroFaseProjeto {
    param($DocumentoBase, $PoderConcedente)

    if ($PoderConcedente -eq 'ARTESP') { return $null }

    $faseProjeto = Get-ValorAtributo -Documento $DocumentoBase -Nome 'FaseProjeto'
    if (-not $faseProjeto) { return $null }

    $cacheKey = "Tipos de Projeto ANTT|$faseProjeto"
    if (-not $script:CacheCadastroFaseProjeto.ContainsKey($cacheKey)) {
        $script:CacheCadastroFaseProjeto[$cacheKey] = Get-PWDocumentsBySearch `
            -Environment 'dmsRegistro' `
            -Attributes @{ TipoRegistro = 'Tipos de Projeto ANTT'; Codigo = $faseProjeto } `
            -GetAttributes
    }

    return $script:CacheCadastroFaseProjeto[$cacheKey]
}

function TemAtributosParaRoteamento {
    param($Doc, $PoderConcedente)

    if ($null -eq $Doc -or $null -eq $Doc.Attributes -or $null -eq $Doc.Attributes[0]) { return $false }

    $disciplina = Get-ValorAtributo -Documento $Doc -Nome 'Disciplina'
    $faseProjeto = Get-ValorAtributo -Documento $Doc -Nome 'FaseProjeto'

    if ($PoderConcedente -eq 'ARTESP') {
        return [bool]$disciplina
    }

    return ([bool]$disciplina -and [bool]$faseProjeto)
}

function TentarEncontrarHospedeiro {
    param($Documento)

    $rels = $null

    try {
        $rels = Get-PWDocumentRelationships -InputDocuments $Documento -ReferencedBy -ErrorAction Stop
    }
    catch {
        try {
            $rels = Get-PWDocumentReferences -InputDocuments $Documento -ReferencedBy -ErrorAction Stop
        }
        catch {
            return $null
        }
    }

    if (-not $rels) { return $null }

    $hospedeiros = @($rels | Where-Object { $_.FileName -match '\.(dgn|dwg)$' })
    if ($hospedeiros.Count -gt 0) { return $hospedeiros[0] }

    return @($rels)[0]
}

function ValidaSeModeloFederadoAutoral {
    param($Documento)

    $poderConcedente = Get-ValorAtributo -Documento $Documento -Nome 'PoderConcedente'
    $sequencial      = Get-ValorAtributo -Documento $Documento -Nome 'Sequencial'
    $tipoDocumento   = Get-ValorAtributo -Documento $Documento -Nome 'TipoDocumento'

    if ($poderConcedente -eq 'ANTT' -and ($sequencial -as [int]) -in 800..999) { return $true }
    if ($poderConcedente -eq 'ARTESP' -and $tipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

function CalculaPastaDestino {
    param(
        $Documento,
        [ref] $DocumentoBaseUsado
    )

    if (-not $Documento) { return $null }

    $poderConcedente = Get-ValorAtributo -Documento $Documento -Nome 'PoderConcedente'
    if (-not $poderConcedente) { return $null }

    $docBase = $Documento

    if (-not (TemAtributosParaRoteamento -Doc $Documento -PoderConcedente $poderConcedente)) {
        $hospedeiro = TentarEncontrarHospedeiro -Documento $Documento

        if ($hospedeiro -and (TemAtributosParaRoteamento -Doc $hospedeiro -PoderConcedente $poderConcedente)) {
            $docBase = $hospedeiro
        }
        else {
            return $null
        }
    }

    $DocumentoBaseUsado.Value = $docBase

    $cadastroDisciplina = ObtemCadastroDisciplina -DocumentoBase $docBase -PoderConcedente $poderConcedente
    $disciplina = DefineDisciplinaPai -CadastroDisciplina $cadastroDisciplina -PoderConcedente $poderConcedente
    if (-not $disciplina) { return $null }

    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -DocumentoBase $docBase -PoderConcedente $poderConcedente
    $faseProjeto = DefineFaseProjeto -CadastroDisciplina $cadastroDisciplina -CadastroFaseProjeto $cadastroFaseProjeto -PoderConcedente $poderConcedente
    if (-not $faseProjeto) { return $null }

    $volume = Get-ValorAtributo -Documento $docBase -Nome 'Volume'

    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $docBase
    if (-not $pastaRaizProjeto) { return $null }

    $caminhoEspecialModeloAutoral = ''
    $disciplinaDocumento = Get-ValorAtributo -Documento $Documento -Nome 'Disciplina'
    $tipoDocumento = Get-ValorAtributo -Documento $Documento -Nome 'TipoDocumento'

    if ($disciplinaDocumento -ne 'U4' -and (ValidaSeModeloFederadoAutoral -Documento $Documento)) {
        $caminhoEspecialModeloAutoral = 'Modelo BIM\Modelos Autorais'
    }

    if ($poderConcedente -eq 'ARTESP' -and $tipoDocumento -eq 'MI') {
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
    ) | Where-Object { $_ }

    $pastaDestino = $pastaDestino -join '\'

    $pasta = Get-PWFoldersCached -FolderPath $pastaDestino
    if ($pasta) { return $pastaDestino }

    return $null
}

function ObtemDocumentosAnterioresNaPastaDestino {
    param($PastaDestino, $NumeroDocumento, $SequencialEmissao)

    if (-not $PastaDestino -or -not $NumeroDocumento -or -not $SequencialEmissao) { return @() }

    $documentos = @(Get-PWDocumentsBySearch `
        -FolderPath $PastaDestino `
        -JustThisFolder `
        -Attributes @{ NumeroPoderConcedente = $NumeroDocumento } `
        -GetAttributes `
        -WarningAction SilentlyContinue)

    return @($documentos | Where-Object { $_.Attributes[0].SequencialEmissao -lt $SequencialEmissao })
}

function MovimentaDocumentosAnterioresParaPastaSuperado {
    param($PastaDestino, $Documento)

    $numeroDocumento = Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente'
    $sequencialEmissao = Get-ValorAtributo -Documento $Documento -Nome 'SequencialEmissao'

    if (-not $numeroDocumento -or -not $sequencialEmissao) { return }

    $documentosAnteriores = @(ObtemDocumentosAnterioresNaPastaDestino `
        -PastaDestino $PastaDestino `
        -NumeroDocumento $numeroDocumento `
        -SequencialEmissao $sequencialEmissao)

    if ($documentosAnteriores.Count -eq 0) { return }

    $pastaSuperados = $PastaDestino + '\Superados'
    $null = Get-PWFoldersCached -FolderPath $pastaSuperados

    $documentosMovimentados = Move-PWDocumentsToFolder `
        -InputDocument $documentosAnteriores `
        -TargetFolderPath $pastaSuperados `
        -ErrorAction Stop

    if ($documentosMovimentados) {
        Set-PWDocumentState -InputDocuments $documentosMovimentados -State 'Superado' -Force
        $script:Contador.MovidosParaSuperados += @($documentosMovimentados).Count
    }
}

function CalculaState {
    param($StateAtual)

    switch ($StateAtual) {
        'Emitido pela Engenharia'                 { 'Nova emissao sendo analisada pela Engenharia' }
        'Solicitado reanalise da Engenharia'      { 'Nova emissao sendo analisada pela Engenharia' }
        'Enviado ao Poder Concedente'             { 'Enviado ao Poder Concedente - Nova emissao em analise Eng' }
        'Concluido pela Unidade'                  { 'Concluido - Nova emissao em analise Eng' }
        Default                                   { $null }
    }
}

function AtualizaStateDocumentosEmitidosParaUnidade {
    param($Documento, $PastaDestino)

    $pastaUnidade = $PastaDestino.Replace('1 - Area de Trabalho', '2 - Unidade')
    $numeroDocumento = Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente'
    $sequencialDocumento = Get-ValorAtributo -Documento $Documento -Nome 'SequencialEmissao'

    if (-not $numeroDocumento -or -not $sequencialDocumento) { return }

    $documentosUnidade = @(Get-PWDocumentsBySearch `
        -FolderPath $pastaUnidade `
        -JustThisFolder `
        -Attributes @{ NumeroPoderConcedente = $numeroDocumento } `
        -GetAttributes `
        -WarningAction SilentlyContinue)

    if ($documentosUnidade.Count -eq 0) { return }

    $docUnidadeMaisRecente = $documentosUnidade |
        Sort-Object { $_.Attributes[0].SequencialEmissao } -Descending |
        Select-Object -First 1

    if ($docUnidadeMaisRecente.Attributes[0].SequencialEmissao -ge $sequencialDocumento) { return }

    $stateAtual = $docUnidadeMaisRecente.WorkflowState
    $proximoState = CalculaState -StateAtual $stateAtual

    if ($proximoState) {
        Set-PWDocumentState -InputDocuments $documentosUnidade -State $proximoState -Force
        $script:Contador.UnidadeAtualizada += $documentosUnidade.Count
    }
}

function BuscaDocumentoDuplicadoNoDestino {
    param($Documento, [string]$PastaDestino)

    $numeroDocumento = Get-ValorAtributo -Documento $Documento -Nome 'NumeroPoderConcedente'
    $sequencialEmissao = Get-ValorAtributo -Documento $Documento -Nome 'SequencialEmissao'

    # Busca mais especifica por NumeroPoderConcedente quando o atributo existe.
    if ($numeroDocumento) {
        $params = @{
            FolderPath     = $PastaDestino
            JustThisFolder = $true
            Attributes     = @{ NumeroPoderConcedente = $numeroDocumento }
            GetAttributes  = $true
            WarningAction  = 'SilentlyContinue'
        }

        $encontrados = @(Get-PWDocumentsBySearch @params)

        if ($encontrados.Count -gt 0) {
            $duplicado = $encontrados |
                Where-Object {
                    $_.Name -eq $Documento.Name -or
                    $_.FileName -eq $Documento.FileName -or
                    ($sequencialEmissao -and $_.Attributes[0].SequencialEmissao -eq $sequencialEmissao)
                } |
                Select-Object -First 1

            if ($duplicado) { return $duplicado }
        }
    }

    # Fallback: busca por FileName se o cmdlet aceitar o parametro.
    try {
        if ($Documento.FileName) {
            $porArquivo = Get-PWDocumentsBySearch `
                -FolderPath $PastaDestino `
                -JustThisFolder `
                -FileName $Documento.FileName `
                -GetAttributes `
                -WarningAction SilentlyContinue

            if ($porArquivo) { return @($porArquivo)[0] }
        }
    }
    catch {
        # Alguns ambientes/cmdlets podem nao aceitar FileName. Nesse caso, usa busca ampla.
    }

    $existentes = @(Get-PWDocumentsBySearch `
        -FolderPath $PastaDestino `
        -JustThisFolder `
        -GetAttributes `
        -WarningAction SilentlyContinue)

    return @($existentes | Where-Object { $_.Name -eq $Documento.Name -or $_.FileName -eq $Documento.FileName } | Select-Object -First 1)
}

function Copiar-DocumentoSemArquivo {
    param($Documento, [string]$PastaDestino)

    $duplicado = BuscaDocumentoDuplicadoNoDestino -Documento $Documento -PastaDestino $PastaDestino
    if ($duplicado) { return $duplicado }

    $temp = Join-Path $env:TEMP ('PWEmpty_' + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $temp | Out-Null

    $nomeArquivo = if ($Documento.FileName) { $Documento.FileName } else { ($Documento.Name + '.tmp') }
    $tempFile = Join-Path $temp $nomeArquivo
    New-Item -Path $tempFile -ItemType File | Out-Null

    try {
        $novo = $null
        $newCmd = Get-CmdletCached -Name 'New-PWDocument'

        if ($newCmd -and $newCmd.Parameters.ContainsKey('FilePath') -and $newCmd.Parameters.ContainsKey('FolderPath')) {
            $params = @{ FolderPath = $PastaDestino; FilePath = $tempFile }

            if ($newCmd.Parameters.ContainsKey('Description') -and $Documento.Description) { $params['Description'] = $Documento.Description }
            if ($newCmd.Parameters.ContainsKey('Name') -and $Documento.Name) { $params['Name'] = $Documento.Name }

            $novo = New-PWDocument @params -ErrorAction Stop
        }
        else {
            $importFromFolder = Get-CmdletCached -Name 'Import-PWDocumentsFromFolder'

            if ($importFromFolder) {
                $novo = Import-PWDocumentsFromFolder `
                    -InputFolder $temp `
                    -ProjectWiseFolder $PastaDestino `
                    -JustOneLevel `
                    -ErrorAction Stop `
                    -WarningAction SilentlyContinue |
                    Select-Object -First 1
            }
            else {
                $importPlanilha = Get-CmdletCached -Name 'Import-PWDocuments'

                if ($importPlanilha) {
                    $novo = Import-PWDocuments `
                        -InputFolder $temp `
                        -FolderPath $PastaDestino `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue |
                        Select-Object -First 1
                }
                else {
                    throw 'Nenhum cmdlet disponivel para importar/criar documento com arquivo.'
                }
            }
        }

        if (-not $novo -and $Documento.FileName) {
            $novo = Get-PWDocumentsBySearch `
                -FolderPath $PastaDestino `
                -JustThisFolder `
                -FileName $Documento.FileName `
                -WarningAction SilentlyContinue |
                Select-Object -First 1
        }

        return $novo
    }
    finally {
        Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function CopiarDocumentoParaDestino {
    param($Documento, [string]$PastaDestino)

    $documentoCopiado = $null

    try {
        $documentoCopiado = Copy-PWDocumentsToFolder `
            -InputDocument $Documento `
            -TargetFolderPath $PastaDestino `
            -ErrorAction Stop `
            -WarningAction Stop
    }
    catch {
        Write-Warning ('Copy-PWDocumentsToFolder falhou para {0}: {1}' -f (Get-NomeDocumentoParaExibicao $Documento), $_.Exception.Message)
    }

    if ($documentoCopiado) { return @($documentoCopiado)[0] }

    $temp = Join-Path $env:TEMP ('PWCopy_' + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $temp | Out-Null

    try {
        $caminhoArquivo = if ($Documento.FileName) { Join-Path $temp $Documento.FileName } else { $null }
        $exportOk = $false

        try {
            $null = Export-PWDocuments `
                -InputDocuments $Documento `
                -OutputFolder $temp `
                -ErrorAction SilentlyContinue `
                -WarningAction SilentlyContinue

            if ($caminhoArquivo) { $exportOk = Test-Path $caminhoArquivo }
        }
        catch {
            Write-Warning ('Export-PWDocuments falhou para {0}: {1}' -f (Get-NomeDocumentoParaExibicao $Documento), $_.Exception.Message)
        }

        if ($exportOk) {
            try {
                $importFromFolder = Get-CmdletCached -Name 'Import-PWDocumentsFromFolder'

                if ($importFromFolder) {
                    $documentoCopiado = Import-PWDocumentsFromFolder `
                        -InputFolder $temp `
                        -ProjectWiseFolder $PastaDestino `
                        -JustOneLevel `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue
                }
                else {
                    $importPlanilha = Get-CmdletCached -Name 'Import-PWDocuments'

                    if ($importPlanilha) {
                        $documentoCopiado = Import-PWDocuments `
                            -InputFolder $temp `
                            -FolderPath $PastaDestino `
                            -ErrorAction Stop `
                            -WarningAction SilentlyContinue
                    }
                    else {
                        $documentoCopiado = New-PWDocument `
                            -FolderPath $PastaDestino `
                            -Name $Documento.Name `
                            -FilePath $caminhoArquivo `
                            -ErrorAction Stop
                    }
                }
            }
            catch {
                Write-Warning ('Falha ao importar/criar a partir do arquivo exportado para {0}: {1}' -f (Get-NomeDocumentoParaExibicao $Documento), $_.Exception.Message)
            }

            if ($documentoCopiado -is [System.Array]) {
                $documentoCopiado = $documentoCopiado |
                    Where-Object { $_.FolderPath -eq $PastaDestino -and $_.FileName -eq $Documento.FileName } |
                    Select-Object -First 1
            }

            if (-not $documentoCopiado -and $Documento.FileName) {
                $documentoCopiado = Get-PWDocumentsBySearch `
                    -FolderPath $PastaDestino `
                    -JustThisFolder `
                    -FileName $Documento.FileName `
                    -WarningAction SilentlyContinue |
                    Select-Object -First 1
            }
        }
        else {
            try {
                $documentoCopiado = Copiar-DocumentoSemArquivo -Documento $Documento -PastaDestino $PastaDestino
            }
            catch {
                Write-Warning ('Falha ao criar documento sem arquivo para {0}: {1}' -f (Get-NomeDocumentoParaExibicao $Documento), $_.Exception.Message)
            }
        }

        return $documentoCopiado
    }
    finally {
        Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function ExibirResumoFinal {
    Write-Host ''
    Write-Host '============================================================'
    Write-Host 'Resumo da movimentacao de documentos'
    Write-Host '============================================================'

    foreach ($item in $script:Contador.GetEnumerator()) {
        Write-Host ('{0}: {1}' -f $item.Key, $item.Value)
    }

    Write-Host '============================================================'
}

#-------------------------------------------------------
# Inicio da execucao
#-------------------------------------------------------
try {
    New-PWLogin `
        -DatasourceName $DatasourceName `
        -Password $SecurePassword `
        -UserName $UserName

    $documentos = @(ObterDocumentosValidados)
    $totalDocumentos = $documentos.Count
    $script:Contador.TotalEncontrado = $totalDocumentos

    Write-Host ('Documentos validados encontrados: {0}' -f $totalDocumentos)

    if ($totalDocumentos -eq 0) {
        Write-Host 'Nenhum documento encontrado para processamento.'
        return
    }

    for ($i = 0; $i -lt $totalDocumentos; $i++) {
        $documento = $documentos[$i]
        $script:Contador.Processados++

        $nomeExibicao = Get-NomeDocumentoParaExibicao -Documento $documento
        $percentual = [math]::Round((($i + 1) / $totalDocumentos) * 100, 2)

        Write-Progress `
            -Activity 'Movimentacao de documentos validados' `
            -Status ('Processando {0}/{1} - {2}' -f ($i + 1), $totalDocumentos, $nomeExibicao) `
            -PercentComplete $percentual

        try {
            $docBaseRef = New-Object PSObject
            $pastaDestino = CalculaPastaDestino -Documento $documento -DocumentoBaseUsado ([ref]$docBaseRef)

            if (-not $pastaDestino) {
                $script:Contador.PastaDestinoNaoLocalizada++
                Set-PWDocumentState `
                    -InputDocuments $documento `
                    -State 'Validado pelo Sistema - Pasta destino nao localizada' `
                    -Force
                continue
            }

            $srcFolder = if ($documento.FolderPath) { $documento.FolderPath.TrimEnd('\').ToLowerInvariant() } else { '' }
            $dstFolder = $pastaDestino.TrimEnd('\').ToLowerInvariant()

            if ($srcFolder -eq $dstFolder) {
                $script:Contador.JaEstavamNoDestino++
                Set-PWDocumentState `
                    -InputDocuments $documento `
                    -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
                    -Force

                AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
                continue
            }

            MovimentaDocumentosAnterioresParaPastaSuperado -PastaDestino $pastaDestino -Documento $documento

            $jaExiste = BuscaDocumentoDuplicadoNoDestino -Documento $documento -PastaDestino $pastaDestino
            if ($jaExiste) {
                $script:Contador.DuplicadosNoDestino++
                Set-PWDocumentState `
                    -InputDocuments $documento `
                    -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
                    -Force

                AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
                continue
            }

            $documentoCopiado = CopiarDocumentoParaDestino -Documento $documento -PastaDestino $pastaDestino

            if (-not $documentoCopiado) {
                $script:Contador.FalhaCopia++
                Write-Warning ('Copia nao concluida, com ou sem arquivo, para: {0}' -f $nomeExibicao)
                continue
            }

            Set-PWDocumentState `
                -InputDocuments $documentoCopiado `
                -State 'Em analise do Assistente' `
                -Force

            Set-PWDocumentState `
                -InputDocuments $documento `
                -State 'Validado pelo Sistema - Copiado p/ Disciplina' `
                -Force

            AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino

            $script:Contador.Copiados++
        }
        catch {
            $script:Contador.Erros++
            Write-Warning ('Erro ao processar {0}: {1}' -f $nomeExibicao, $_.Exception.Message)
        }
    }
}
finally {
    Write-Progress -Activity 'Movimentacao de documentos validados' -Completed
    ExibirResumoFinal

    try {
        Undo-PWLogin
    }
    catch {
        Write-Warning ('Falha ao encerrar login no ProjectWise: {0}' -f $_.Exception.Message)
    }
}
