
function Mostrar-Log {
    param(
        [string]$Mensagem,
        [string]$Tipo = "INFO"
    )

    $hora = Get-Date -Format "HH:mm:ss"

    switch ($Tipo) {
        "OK"      { Write-Host "[$hora] [OK] $Mensagem" -ForegroundColor Green }
        "ERRO"    { Write-Host "[$hora] [ERRO] $Mensagem" -ForegroundColor Red }
        "AVISO"   { Write-Host "[$hora] [AVISO] $Mensagem" -ForegroundColor Yellow }
        "ETAPA"   { Write-Host "[$hora] [ETAPA] $Mensagem" -ForegroundColor Cyan }
        "MOV"     { Write-Host "[$hora] [MOVIMENTACAO] $Mensagem" -ForegroundColor Magenta }
        default   { Write-Host "[$hora] [INFO] $Mensagem" -ForegroundColor Gray }
    }
}

function Obter-NomeDocumentoLog {
    param($documento)

    if ($documento.FileName) { return $documento.FileName }
    if ($documento.Name) { return $documento.Name }
    return "Documento sem nome identificado"
}


function Formatar-TamanhoArquivo {
    param($Bytes)

    if ($null -eq $Bytes -or $Bytes -le 0) { return "Tamanho não identificado" }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0:N0} bytes" -f $Bytes)
}

function Obter-TamanhoDocumento {
    param($documento)

    foreach ($prop in @('FileSize', 'Size', 'DocumentSize')) {
        if ($documento.PSObject.Properties.Name -contains $prop -and $documento.$prop) {
            return [double]$documento.$prop
        }
    }

    try {
        if ($documento.Attributes -and $documento.Attributes[0]) {
            foreach ($prop in @('FileSize', 'Size', 'DocumentSize')) {
                if ($documento.Attributes[0].PSObject.Properties.Name -contains $prop -and $documento.Attributes[0].$prop) {
                    return [double]$documento.Attributes[0].$prop
                }
            }
        }
    }
    catch {}

    return $null
}

function Atualizar-ProgressoGeral {
    param(
        [int]$Atual,
        [int]$Total,
        [string]$Status
    )

    if ($Total -le 0) { return }
    $percentual = [math]::Round(($Atual / $Total) * 100, 0)
    Write-Progress -Id 1 -Activity "Movimentação de documentos validados" -Status $Status -PercentComplete $percentual
}

function Atualizar-ProgressoCopia {
    param(
        [string]$NomeArquivo,
        [string]$TamanhoFormatado,
        [int]$Percentual
    )

    Write-Progress -Id 2 -ParentId 1 -Activity "Cópia do arquivo atual" -Status "$NomeArquivo | $TamanhoFormatado" -PercentComplete $Percentual
}

function Obter-PastaDestinoComCache {
    param([string]$pastaDestino)

    if (-not $script:CachePastasDestino) { $script:CachePastasDestino = @{} }

    if ($script:CachePastasDestino.ContainsKey($pastaDestino)) {
        if ($script:CachePastasDestino[$pastaDestino]) { return $pastaDestino }
        return $null
    }

    $pasta = Get-PWFolders -FolderPath $pastaDestino
    $script:CachePastasDestino[$pastaDestino] = [bool]$pasta

    if ($pasta) { return $pastaDestino }
    return $null
}

function Obter-DocumentosDaPastaComCache {
    param([string]$pastaDestino)

    if (-not $script:CacheDocumentosPorPasta) { $script:CacheDocumentosPorPasta = @{} }

    if (-not $script:CacheDocumentosPorPasta.ContainsKey($pastaDestino)) {
        $script:CacheDocumentosPorPasta[$pastaDestino] = @(
            Get-PWDocumentsBySearch -FolderPath $pastaDestino -JustThisFolder -GetAttributes -WarningAction SilentlyContinue
        )
    }

    return $script:CacheDocumentosPorPasta[$pastaDestino]
}

function Atualizar-CacheDocumentosDaPasta {
    param(
        [string]$pastaDestino,
        $documento
    )

    if (-not $documento) { return }
    if (-not $script:CacheDocumentosPorPasta) { $script:CacheDocumentosPorPasta = @{} }

    if ($script:CacheDocumentosPorPasta.ContainsKey($pastaDestino)) {
        $script:CacheDocumentosPorPasta[$pastaDestino] = @($script:CacheDocumentosPorPasta[$pastaDestino]) + @($documento)
    }
}

function ObterDocumentosValidados {
    Get-PWDocumentsBySearch -SearchName 'Scripts\Docs GRD Validados' -GetAttributes
}
function DefineDisciplinaPai {
    param ($cadastroDisciplina, $poderConcedente)
    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) { return $null }
    if ($poderConcedente -eq "ARTESP") {
        $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[1]
    }
    else {
        $cadastroDisciplina.Attributes[0].Relacionamento
    }
}
function DefineFaseProjeto {
    param ($cadastroDisciplina, $cadastroFaseProjeto, $poderConcedente)
    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) { return $null }
    if ($poderConcedente -eq "ARTESP") {
        $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[0]
    }
    else {
        if ($null -eq $cadastroFaseProjeto -or $null -eq $cadastroFaseProjeto.Attributes -or $null -eq $cadastroFaseProjeto.Attributes[0].Relacionamento) { return $null }
        $cadastroFaseProjeto.Attributes[0].Relacionamento
    }
}
function ObtemCadastroDisciplina {
    param ($documentoBase, $poderConcedente)
    if ($null -eq $documentoBase -or $null -eq $documentoBase.Attributes -or $null -eq $documentoBase.Attributes[0].Disciplina) { return $null }
    $disciplina = $documentoBase.Attributes[0].Disciplina
    $tipoRegistro = switch ($poderConcedente) {
        "ARTESP" { 'Disciplinas ARTESP' }
        Default { 'Disciplinas ANTT' }
    }
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = $tipoRegistro; Codigo = $disciplina } -GetAttributes
}
function ObtemCadastroFaseProjeto {
    param ($documentoBase, $poderConcedente)
    if ($poderConcedente -eq "ARTESP") { return $null }
    if ($null -eq $documentoBase -or $null -eq $documentoBase.Attributes -or $null -eq $documentoBase.Attributes[0].FaseProjeto) { return $null }
    $faseProjeto = $documentoBase.Attributes[0].FaseProjeto
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = 'Tipos de Projeto ANTT'; Codigo = $faseProjeto } -GetAttributes
}
function TemAtributosParaRoteamento {
    param ($doc, $poderConcedente)
    if ($null -eq $doc -or $null -eq $doc.Attributes -or $null -eq $doc.Attributes[0]) { return $false }
    $a = $doc.Attributes[0]
    if ($poderConcedente -eq 'ARTESP') {
        return [bool]$a.Disciplina
    }
    else {
        return ([bool]$a.Disciplina -and [bool]$a.FaseProjeto)
    }
}
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
        $_.FileName -match '.(dgn|dwg)$'
    }
    if (-not $hospedeiros) { return $rels[0] }
    $hospedeiros[0]
}
function CalculaPastaDestino {
    param ($documento, [ref]$documentoBaseUsado)

    if (-not $documento) { return $null }

    $poderConcedente = $documento.Attributes[0].PoderConcedente
    if (-not $poderConcedente) { return $null }

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

    $cadastroDisciplina = ObtemCadastroDisciplina -documentoBase $docBase -PoderConcedente $poderConcedente
    $disciplina = DefineDisciplinaPai -CadastroDisciplina $cadastroDisciplina -PoderConcedente $poderConcedente
    if (-not $disciplina) { return $null }

    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -documentoBase $docBase -PoderConcedente $poderConcedente
    $faseProjeto = DefineFaseProjeto -CadastroDisciplina $cadastroDisciplina -CadastroFaseProjeto $cadastroFaseProjeto -PoderConcedente $poderConcedente
    if (-not $faseProjeto) { return $null }

    $volume = $docBase.Attributes[0].Volume

    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $docBase
    if (-not $pastaRaizProjeto) { return $null }

    # Melhoria: Exceção para Modelos Autorais
    $caminhoEspecialModeloAutoral = if ($documento.Attributes[0].Disciplina -ne 'U4' -and (ValidaSeModeloFederadoAutoral $documento)) { 'Modelo BIM\Modelos Autorais' } else { '' }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -eq 'MI') { 
        $caminhoEspecialModeloAutoral = 'Modelo BIM' 
        $disciplina = '' # MI, por ser Modelo BIM fica direto na raiz dessa pasta e ignora a disciplina
    }

    $pastaDestino = @($pastaRaizProjeto.FullPath, '1 - Area de Trabalho', $faseProjeto, $volume, $caminhoEspecialModeloAutoral, $disciplina) | Where-Object { $_ }
    $pastaDestino = $pastaDestino -join '\'

    return (Obter-PastaDestinoComCache -pastaDestino $pastaDestino)
}
function ObtemDocumentosAnterioresNaPastaDestino {
    param ($pastaDestino, $numeroDocumento, $sequencialEmissao)

    $documentos = Obter-DocumentosDaPastaComCache -pastaDestino $pastaDestino
    $documentos | Where-Object {
        $_.Attributes -and
        $_.Attributes[0].NumeroPoderConcedente -eq $numeroDocumento -and
        ([int]$_.Attributes[0].SequencialEmissao) -lt ([int]$sequencialEmissao)
    }
}
function MovimentaDocumentosAnterioresParaPastaSuperado {
    param ($pastaDestino, $documento)
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $sequencialEmissao = $documento.Attributes[0].SequencialEmissao
    if (-not $numeroDocumento -or -not $sequencialEmissao) { return }
    $documentosAnteriores = ObtemDocumentosAnterioresNaPastaDestino -PastaDestino $pastaDestino -NumeroDocumento $numeroDocumento -SequencialEmissao $sequencialEmissao
    if (-not $documentosAnteriores) { return }
    $pastaSuperados = $pastaDestino + "\Superados"
    $documentosMovimentados = Move-PWDocumentsToFolder -InputDocument $documentosAnteriores -TargetFolderPath $pastaSuperados
    Set-PWDocumentState -InputDocuments $documentosMovimentados -State 'Superado' -Force

    if ($script:CacheDocumentosPorPasta -and $script:CacheDocumentosPorPasta.ContainsKey($pastaDestino)) {
        $idsMovidos = @($documentosAnteriores | ForEach-Object { $_.DocumentID })
        $script:CacheDocumentosPorPasta[$pastaDestino] = @(
            $script:CacheDocumentosPorPasta[$pastaDestino] | Where-Object { $_.DocumentID -notin $idsMovidos }
        )
    }
}
function CalculaState {
    param ($stateAtual)
    switch ($stateAtual) {
        "Emitido pela Engenharia" { "Nova emissao sendo analisada pela Engenharia" }
        "Solicitado reanalise da Engenharia" { "Nova emissao sendo analisada pela Engenharia" }
        "Enviado ao Poder Concedente" { "Enviado ao Poder Concedente - Nova emissao em analise Eng" }
        "Concluido pela Unidade" { "Concluido - Nova emissao em analise Eng" }
        Default { $null }
    }
}
function AtualizaStateDocumentosEmitidosParaUnidade {
    param ($documento, $pastaDestino)
    $pastaUnidade = $pastaDestino.Replace('1 - Area de Trabalho', '2 - Unidade')
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $documentosUnidade = Get-PWDocumentsBySearch -FolderPath $pastaUnidade -JustThisFolder -Attributes @{ NumeroPoderConcedente = $numeroDocumento } -GetAttributes
    if (-not $documentosUnidade -or $documentosUnidade[0].Attributes[0].SequencialEmissao -ge $documento.Attributes[0].SequencialEmissao) { return }
    $stateAtual = $documentosUnidade[0].WorkflowState
    $proximoState = CalculaState -StateAtual $stateAtual
    if ($proximoState) {
        Set-PWDocumentState -InputDocuments $documentosUnidade -State $proximoState -Force
    }
}



function Copiar-DocumentoSemArquivo {
    param($documento, [string]$pastaDestino)

    $existentes = Obter-DocumentosDaPastaComCache -pastaDestino $pastaDestino
    if ($existentes) {
        $dup = $existentes | Where-Object { $_.Name -eq $documento.Name -or $_.FileName -eq $documento.FileName } | Select-Object -First 1
        if ($dup) { return $dup }
    }


    $temp = Join-Path $env:TEMP ("PWEmpty_" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $temp | Out-Null
    $tempFile = Join-Path $temp $documento.FileName
    New-Item -Path $tempFile -ItemType File | Out-Null

    try {
        $novo = $null
        $newCmd = Get-Command -Name New-PWDocument -ErrorAction SilentlyContinue
        if ($newCmd -and $newCmd.Parameters.ContainsKey('FilePath') -and $newCmd.Parameters.ContainsKey('FolderPath')) {

            $params = @{ FolderPath = $pastaDestino; FilePath = $tempFile }
            if ($newCmd.Parameters.ContainsKey('Description') -and $documento.Description) { $params['Description'] = $documento.Description }
            if ($newCmd.Parameters.ContainsKey('Name')) { $params['Name'] = $documento.Name }
            $novo = New-PWDocument @params -ErrorAction Stop
        }
        else {
       
            $importFromFolder = Get-Command -Name Import-PWDocumentsFromFolder -ErrorAction SilentlyContinue
            if ($importFromFolder) {
                $novo = Import-PWDocumentsFromFolder -InputFolder $temp -ProjectWiseFolder $pastaDestino -JustOneLevel -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -First 1
            }
            else {
                $importPlanilha = Get-Command -Name Import-PWDocuments -ErrorAction SilentlyContinue
                if ($importPlanilha) {
                    $novo = Import-PWDocuments -InputFolder $temp -FolderPath $pastaDestino -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -First 1
                }
                else {
                    throw "Nenhum cmdlet disponível para importar/criar documento com arquivo."
                }
            }
        }

        if (-not $novo) {
     
            $novo = Get-PWDocumentsBySearch -FolderPath $pastaDestino -JustThisFolder -FileName $documento.FileName -WarningAction SilentlyContinue | Select-Object -First 1
        }

        return $novo
    }
    finally {
        try { Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}
function ValidaSeModeloFederadoAutoral {
    param($documento)

    if ($documento.Attributes[0].PoderConcedente -eq 'ANTT' -and $documento.Attributes[0].Sequencial -as [int] -in 800..999) { return $true }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

#-------------------------------------------------------

#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force

Mostrar-Log "Conectando ao ProjectWise..." "ETAPA"
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' -Password $SecurePassword -UserName 'admin'
Mostrar-Log "Login realizado com sucesso." "OK"

$script:CachePastasDestino = @{}
$script:CacheDocumentosPorPasta = @{}

Mostrar-Log "Analisando a pesquisa salva: Scripts\Docs GRD Validados" "ETAPA"
$documentos = @(ObterDocumentosValidados)
$totalDocumentos = $documentos.Count

Mostrar-Log "Foram encontrados $totalDocumentos arquivo(s) em Docs GRD Validados." "OK"

if ($totalDocumentos -eq 0) {
    Mostrar-Log "Nenhum arquivo para movimentar. Encerrando execução." "AVISO"
    Undo-PWLogin
    return
}

Mostrar-Log "Iniciando processo de movimentação..." "ETAPA"

$contador = 0
$movidosComSucesso = 0
$naoMovidos = 0
$jaExistentesOuJaNaPasta = 0

foreach ($documento in $documentos) {
    $contador++
    $nomeDocumentoLog = Obter-NomeDocumentoLog -documento $documento
    $tamanhoBytes = Obter-TamanhoDocumento -documento $documento
    $tamanhoFormatado = Formatar-TamanhoArquivo -Bytes $tamanhoBytes
    Atualizar-ProgressoGeral -Atual $contador -Total $totalDocumentos -Status ("Processando {0} de {1}: {2}" -f $contador, $totalDocumentos, $nomeDocumentoLog)

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor DarkGray
    Mostrar-Log "Processando arquivo $contador de $totalDocumentos" "ETAPA"
    Mostrar-Log "Arquivo atual: $nomeDocumentoLog" "INFO"
    Mostrar-Log "Tamanho: $tamanhoFormatado" "INFO"
    Mostrar-Log "Origem: $($documento.FolderPath)" "INFO"
    Mostrar-Log "Arquivos movidos até agora: $movidosComSucesso de $totalDocumentos" "MOV"

    $docBaseRef = New-Object PSObject
    Mostrar-Log "Calculando pasta de destino..." "INFO"
    $pastaDestino = CalculaPastaDestino -Documento $documento -documentoBaseUsado ([ref]$docBaseRef)

    if (-not $pastaDestino) {
        Mostrar-Log "Pasta destino não localizada. Alterando estado do documento." "ERRO"
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Pasta destino nao localizada' -Force
        $naoMovidos++
        Mostrar-Log "Movimentação não realizada para: $nomeDocumentoLog" "AVISO"
        continue
    }

    Mostrar-Log "Destino calculado: $pastaDestino" "INFO"

    $srcFolder = ($documento.FolderPath.TrimEnd('\')).ToLowerInvariant()
    $dstFolder = ($pastaDestino.TrimEnd('\')).ToLowerInvariant()
    if ($srcFolder -eq $dstFolder) {
        Mostrar-Log "O documento já está na pasta correta. Apenas atualizando o estado." "AVISO"
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Copiado p/ Disciplina' -Force
        AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
        $jaExistentesOuJaNaPasta++
        Mostrar-Log "Arquivos movidos até agora: $movidosComSucesso de $totalDocumentos" "MOV"
        continue
    }

    Mostrar-Log "Verificando emissões anteriores para enviar à pasta Superados, se existirem..." "INFO"
    MovimentaDocumentosAnterioresParaPastaSuperado -PastaDestino $pastaDestino -Documento $documento

    Mostrar-Log "Verificando se o arquivo já existe no destino..." "INFO"
    $existentes = Obter-DocumentosDaPastaComCache -pastaDestino $pastaDestino
    $jaExiste = $existentes | Where-Object { $_.Name -eq $documento.Name -or $_.FileName -eq $documento.FileName }
    if ($jaExiste) {
        Mostrar-Log "O arquivo já existe no destino. Não será copiado novamente." "AVISO"
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Copiado p/ Disciplina' -Force
        AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
        $jaExistentesOuJaNaPasta++
        Mostrar-Log "Arquivos movidos até agora: $movidosComSucesso de $totalDocumentos" "MOV"
        continue
    }

    Mostrar-Log "Iniciando cópia para o destino..." "MOV"
    Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 10
    Mostrar-Log "$nomeDocumentoLog" "INFO"
    Mostrar-Log "De:   $($documento.FolderPath)" "INFO"
    Mostrar-Log "Para: $pastaDestino" "INFO"

    $documentoCopiado = $null
    try {
        Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 35
        $documentoCopiado = Copy-PWDocumentsToFolder -InputDocument $documento -TargetFolderPath $pastaDestino -ErrorAction Stop -WarningAction Stop
        Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 75
    }
    catch {
        Mostrar-Log "Cópia direta falhou. Tentando método alternativo por exportação/importação." "AVISO"
    }

    if (-not $documentoCopiado) {
        # Exporta; se não houver arquivo físico, cria doc sem arquivo
        $temp = Join-Path $env:TEMP ("PWCopy_" + [guid]::NewGuid())
        New-Item -ItemType Directory -Path $temp | Out-Null
        $caminhoArquivo = Join-Path $temp $documento.FileName
        $exportOk = $false
        try {
            $null = Export-PWDocuments -InputDocuments $documento -OutputFolder $temp -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $exportOk = Test-Path $caminhoArquivo
        }
        catch {}

        if ($exportOk) {
            Mostrar-Log "Arquivo exportado temporariamente. Importando no destino..." "INFO"
            Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 55
            try {
                $importFromFolder = Get-Command -Name Import-PWDocumentsFromFolder -ErrorAction SilentlyContinue
                if ($importFromFolder) {
                    $documentoCopiado = Import-PWDocumentsFromFolder -InputFolder $temp -ProjectWiseFolder $pastaDestino -JustOneLevel -ErrorAction Stop -WarningAction SilentlyContinue
                }
                else {
                    $importPlanilha = Get-Command -Name Import-PWDocuments -ErrorAction SilentlyContinue
                    if ($importPlanilha) {
                        $documentoCopiado = Import-PWDocuments -InputFolder $temp -FolderPath $pastaDestino -ErrorAction Stop -WarningAction SilentlyContinue
                    }
                    else {
                        $documentoCopiado = New-PWDocument -FolderPath $pastaDestino -Name $documento.Name -FilePath $caminhoArquivo -ErrorAction Stop
                    }
                }
            }
            catch {
                Mostrar-Log ("Falha ao importar/criar a partir do arquivo exportado: " + $_.Exception.Message) "ERRO"
            }

            if ($documentoCopiado -is [System.Array]) {
                $documentoCopiado = ($documentoCopiado | Where-Object { $_.FolderPath -eq $pastaDestino -and $_.FileName -eq $documento.FileName } | Select-Object -First 1)
            }
            if (-not $documentoCopiado) {
                $documentoCopiado = Get-PWDocumentsBySearch -FolderPath $pastaDestino -JustThisFolder -FileName $documento.FileName -WarningAction SilentlyContinue
                if ($documentoCopiado -is [System.Array]) { $documentoCopiado = $documentoCopiado | Select-Object -First 1 }
            }
        }
        else {
            Mostrar-Log "Não foi possível exportar arquivo físico. Tentando criar documento sem arquivo." "AVISO"
            try {
                $documentoCopiado = Copiar-DocumentoSemArquivo -documento $documento -pastaDestino $pastaDestino
            }
            catch {
                Mostrar-Log ("Falha ao criar documento sem arquivo: " + $_.Exception.Message) "ERRO"
            }
        }

        try { Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }

    if (-not $documentoCopiado) {
        Mostrar-Log "Cópia não concluída para: $nomeDocumentoLog" "ERRO"
        $naoMovidos++
        Mostrar-Log "Arquivos movidos até agora: $movidosComSucesso de $totalDocumentos" "MOV"
        continue
    }

    Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 90
    Mostrar-Log "Cópia concluída. Atualizando estados do documento copiado e do original..." "OK"
    Set-PWDocumentState -InputDocuments $documentoCopiado -State 'Em analise do Assistente' -Force
    Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Copiado p/ Disciplina' -Force
    AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
    Atualizar-CacheDocumentosDaPasta -pastaDestino $pastaDestino -documento $documentoCopiado
    Atualizar-ProgressoCopia -NomeArquivo $nomeDocumentoLog -TamanhoFormatado $tamanhoFormatado -Percentual 100
    Write-Progress -Id 2 -ParentId 1 -Activity "Cópia do arquivo atual" -Completed

    $movidosComSucesso++
    Mostrar-Log "Movimentação concluída para: $nomeDocumentoLog" "OK"
    Mostrar-Log "Arquivos movidos até agora: $movidosComSucesso de $totalDocumentos" "MOV"
}

Write-Progress -Id 1 -Activity "Movimentação de documentos validados" -Completed
Write-Progress -Id 2 -Activity "Cópia do arquivo atual" -Completed

Write-Host ""
Write-Host "============================================================" -ForegroundColor DarkGray
Mostrar-Log "Processo finalizado." "ETAPA"
Mostrar-Log "Total encontrado: $totalDocumentos" "INFO"
Mostrar-Log "Movidos com sucesso: $movidosComSucesso" "OK"
Mostrar-Log "Já existentes ou já estavam na pasta correta: $jaExistentesOuJaNaPasta" "AVISO"
Mostrar-Log "Não movidos: $naoMovidos" "ERRO"

Undo-PWLogin
Mostrar-Log "Logout realizado no ProjectWise." "OK"
