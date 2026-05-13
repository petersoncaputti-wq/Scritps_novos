function ObterDocumentosValidados {
    $documentos = Get-PWDocumentsBySearch -SearchName 'Scripts\Importacao\Docs Importados' -GetAttributes
    return $documentos
}
function DefineDisciplinaPai {
    param ($cadastroDisciplina, $poderConcedente)

    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) {
        return $null
    }

    if($poderConcedente -eq "ARTESP"){
        $disciplinaPai = $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[1]
    }
    else{
        $disciplinaPai = $cadastroDisciplina.Attributes[0].Relacionamento
    }

    return $disciplinaPai
}
function DefineFaseProjeto {
    param ($cadastroDisciplina, $cadastroFaseProjeto, $poderConcedente)

    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) {
        return $null
    }

    if($poderConcedente -eq "ARTESP"){
        $faseProjeto = $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[0]
    }
    else{
        $faseProjeto = $cadastroFaseProjeto.Attributes[0].Relacionamento
    }
    return $faseProjeto
}
function ObtemCadastroDisciplina {
    param ($documento, $poderConcedente)

    if ($null -eq $documento -or $null -eq $documento.Attributes.Disciplina) {
        return $null
    }

    $disciplina = $documento.Attributes.Disciplina    
    $tipoRegistro = switch ($poderConcedente) {
        "ARTESP" { 'Disciplinas ARTESP' }
        Default { 'Disciplinas ANTT'}
    }
    $cadastroDisciplina = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{TipoRegistro = $tipoRegistro; Codigo = $disciplina } -GetAttributes

    return $cadastroDisciplina
}
function ObtemCadastroFaseProjeto {
    param ($documento, $poderConcedente)

    if ($null -eq $documento -or $null -eq $documento.Attributes.Disciplina) {
        return $null
    }    

    if($poderConcedente -eq "ARTESP"){
        return $null
    }

    $faseProjeto = $documento.Attributes.FaseProjeto        
    $cadastroDisciplina = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{TipoRegistro = 'Tipos de Projeto ANTT'; Codigo = $faseProjeto } -GetAttributes

    return $cadastroDisciplina
}
function CalculaPastaDestino {
    param ($documento)

    if (-not $documento) {
        return $null
    }

    $poderConcedente = $documento.Attributes[0].PoderConcedente

    $cadastroDisciplina = ObtemCadastroDisciplina -Documento $documento -PoderConcedente $poderConcedente

    $disciplina = DefineDisciplinaPai -CadastroDisciplina $cadastroDisciplina -PoderConcedente $poderConcedente
    if (-not $disciplina) {
        return $null
    }
    
    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -Documento $documento -poderConcedente $poderConcedente

    $faseProjeto = DefineFaseProjeto -CadastroDisciplina $cadastroDisciplina -cadastroFaseProjeto $cadastroFaseProjeto -poderConcedente $poderConcedente
    if (-not $faseProjeto) {
        return $null
    }
    
    $volume = $documento.Attributes[0].Volume

    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $documento
    if (-not $pastaRaizProjeto) {
        return $null
    }

    if($volume){
        $pastaDestino = $pastaRaizProjeto.FullPath + '\1 - Area de Trabalho\' + $faseProjeto + '\' + $volume + '\' + $disciplina
    }
    else{
        $pastaDestino = $pastaRaizProjeto.FullPath + '\1 - Area de Trabalho\' + $faseProjeto + '\' + $disciplina
    }

    $pasta = Get-PWFolders -FolderPath $pastaDestino
    if($pasta){
        return $pastaDestino
    }
    return $null
}
function ObtemDocumentosAnterioresNaPastaDestino {
    param ($pastaDestino, $numeroDocumento, $sequencialEmissao)

    $documentos = Get-PWDocumentsBySearch -FolderPath $pastaDestino -JustThisFolder -Attributes @{NumeroPoderConcedente = $numeroDocumento } -GetAttributes

    $documentosAnteriores = $documentos.Where({ $_.Attributes[0].SequencialEmissao -lt $sequencialEmissao })

    return $documentosAnteriores
}
function MovimentaDocumentosAnterioresParaPastaSuperado {
    param ($pastaDestino, $documento)
    
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $sequencialEmissao = $documento.Attributes[0].SequencialEmissao

    $documentosAnteriores = ObtemDocumentosAnterioresNaPastaDestino -PastaDestino $pastaDestino -NumeroDocumento $numeroDocumento -SequencialEmissao $sequencialEmissao
    if (-not $documentosAnteriores) {
        return
    }

    $pastaSuperados = $pastaDestino + "\Superados"

    $documentosMovimentados = Move-PWDocumentsToFolder -InputDocument $documentosAnteriores -TargetFolderPath $pastaSuperados

    Set-PWDocumentState -InputDocuments $documentosMovimentados -State 'Superado' -Force
}
function CalculaState {
    param ($stateAtual)
    
    $stateFinal = switch ($stateAtual) {
        "Emitido pela Engenharia" { "Nova emissao sendo analisada pela Engenharia" }
        "Solicitado reanalise da Engenharia" { "Nova emissao sendo analisada pela Engenharia" }
        "Enviado ao Poder Concedente" { "Enviado ao Poder Concedente - Nova emissao em analise Eng" }
        "Concluido pela Unidade" { "Concluido - Nova emissao em analise Eng" }
        Default { $null }
    }

    return $stateFinal
}
function AtualizaStateDocumentosEmitidosParaUnidade {
    param ($documento, $pastaDestino)
    
    $pastaUnidade = $pastaDestino.Replace('1 - Area de Trabalho', '2 - Unidade')
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $documentosUnidade = Get-PWDocumentsBySearch -FolderPath $pastaUnidade -JustThisFolder -Attributes @{NumeroPoderConcedente = $numeroDocumento } -GetAttributes

    if (-not $documentosUnidade -or $documentosUnidade.Attributes[0].SequencialEmissao -ge $documento.Attributes[0].SequencialEmissao) {
        return
    }

    $stateAtual = $documentosUnidade[0].WorkflowState
    $proximoState = CalculaState -StateAtual $stateAtual

    if ($proximoState) {
        Set-PWDocumentState -InputDocuments $documentosUnidade -State $proximoState -Force
    }
}
function defineEnvironment {
    param ($poderConcedente)
    
    if($poderConcedente -eq "ARTESP") { return "dmsEngenhariaARTESP"}
    if($poderConcedente -eq "ANTT") { return "dmsEngenhariaANTT"}
}
function obterEmissoesDocumento {
    param ($documento)

    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $poderConcedente = $documento.Attributes[0].PoderConcedente
    $environment = defineEnvironment -poderConcedente $poderConcedente

    $documentosArquivados = Get-PWDocumentsBySearch -Environment $environment -Attributes @{NumeroPoderConcedente = $numeroDocumento}

    return $documentosArquivados
}
function verificaSeDocumentoEmitido {
    param ($documentosEmitidos, $documento)

    $revisao = $documento.CustomAttributes.Revisao
    $versao = $documento.CustomAttributes.Versao
    $sufixo = $documento.CustomAttributes.Sufixo
    $extensao = $documento.CustomAttributes.Extensao

    $pesquisa = $documentosEmitidos.Where({$_.CustomAttributes.Revisao -eq $revisao -and $_.CustomAttributes.Versao -eq $versao -and $_.CustomAttributes.Sufixo -eq $sufixo -and $_.CustomAttributes.Extensao -eq $extensao})

    if($pesquisa){ return $true }
    return $false
}
function verficaSeUltimaEmissao {
    param ($documento, $emissoes)
    
    $revisao = $documento.Attributes[0].Revisao
    $versao = $documento.Attributes[0].Versao

    if(-not $emissoes){ return $true }

    $pesquisa = $emissoes.Where({$_.Attributes[0].Revisao -gt $revisao})
    if($pesquisa){ return $false }

    $pesquisa = $emissoes.Where({$_.Attributes[0].Revisao -eq $revisao -and $_.Attributes[0].Versao -gt $versao})
    if($pesquisa){ return $false }

    return $true
}
#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' -Password $SecurePassword -UserName 'admin'

$documentos = ObterDocumentosValidados

foreach ($documento in $documentos) {
    $emissoes = obterEmissoesDocumento -documento $documento
    $emitido = verificaSeDocumentoEmitido -documentosEmitidos $emissoes -documento $documento
    if($emitido) { 
        Set-PWDocumentState -InputDocuments $documento -State 'Documento ja importado'
        continue 
    }

    $pastaDestino = CalculaPastaDestino -Documento $documento    
    if($pastaDestino){
        $ultimaEmissao = verficaSeUltimaEmissao -documento $documento -emissoes $emissoes

        if($ultimaEmissao){
            MovimentaDocumentosAnterioresParaPastaSuperado -PastaDestino $pastaDestino -Documento $documento
            $documentoCopiado = Copy-PWDocumentsToFolder -InputDocument $documento -TargetFolderPath $pastaDestino
            Set-PWDocumentState -InputDocuments $documentoCopiado -State "Em analise do Assistente" -Force            
        }
        else{
            $pastaDestino = "$pastaDestino\Superados"
            $documentoCopiado = Copy-PWDocumentsToFolder -InputDocument $documento -TargetFolderPath $pastaDestino
            Set-PWDocumentState -InputDocuments $documentoCopiado -State "Superado" -Force
        }
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Copiado p/ Disciplina' -Force
        AtualizaStateDocumentosEmitidosParaUnidade -Documento $documento -PastaDestino $pastaDestino
    }
    else{
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - Pasta destino nao localizada' -Force
    }
    
}

Undo-PWLogin