function ObterDocumentosASeremEnviadosParaUnidade {
    $documentos = Get-PWDocumentsBySearch -SearchName 'Scripts\Envio p Unidade' -GetAttributes
    return $documentos
}
function CalcularPastaDestino {
    param ($documento)
    
    $pastaDestino = $documento.FolderPath.Replace('1 - Area de Trabalho', '2 - Unidade')
    return $pastaDestino
}
function CalculaState {
    param ($stateAtual)
    
    $stateFinal = switch ($stateAtual) {
        "Em Envio para Unidade" { "Enviado para Unidade" }
        "Em Envio para Unidade c/ Ressalvas" { "Enviado para Unidade com Ressalvas" }
        "Em Envio para Unidade c/ Pendencias do Especialista" { "Enviado para Unidade c/ Pendencias do Especialista" }
        "Em devolucao para Unidade" { "Enviado para Unidade" }
        Default { stateAtual }
    }

    return $stateFinal
}
function defineEnvironmentUnidade {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') { 
        return 'dmsUnidadeARTESP'
    }
    
    if ($poderConcedente -eq 'ANTT') {
        return 'dmsUnidadeANTT'
    }

    return ''
}
function obtemDocumentosAnterioresNaPastaDestino {
    param ($pastaDestino, $numeroDocumento, $sequencialEmissao)

    $documentos = Get-PWDocumentsBySearch -FolderPath $pastaDestino -JustThisFolder -Attributes @{NumeroPoderConcedente = $numeroDocumento } -GetAttributes

    $documentosAnteriores = $documentos.Where({ $_.Attributes[0].SequencialEmissao -lt $sequencialEmissao })

    return $documentosAnteriores
}
function movimentaDocumentosAnterioresSuperados {
    param ($pastaDestino, $documento)

    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $sequencialEmissao = $documento.Attributes[0].SequencialEmissao

    $documentosAnteriores = obtemDocumentosAnterioresNaPastaDestino -pastaDestino $pastaDestino -numeroDocumento $numeroDocumento -sequencialEmissao $sequencialEmissao
    if (-not $documentosAnteriores) {
        return
    }

    $pastaSuperados = $pastaDestino + "\\Superados\\"
    $documentosMovimentados = Move-PWDocumentsToFolder -InputDocument $documentosAnteriores -TargetFolderPath $pastaSuperados
    Set-PWDocumentState -InputDocuments $documentosMovimentados -State 'Superado' -Force
}
function copiaDocumento {
    param ($documento, $pastaDestino)

    if (-not $documento -or -not $pastaDestino) {
        return
    }

    $sufixo = $documento.Attributes[0].Sufixo
    $stateAtual = $documento.WorkflowState
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente
    $revisao = $documento.Attributes[0].Revisao
    $versao = $documento.Attributes[0].Versao
    $poderConcedente = $documento.Attributes[0].PoderConcedente
    $environmentUnidade = defineEnvironmentUnidade -poderConcedente $poderConcedente

    if ([string]::IsNullOrEmpty($sufixo)) {
        if ($stateAtual -ne "Em devolucao para Unidade") {
            $documentoUnidade = Copy-PWDocument -InputDocument $documento -TargetFolderPath $pastaDestino
        }
        else {
            $documentoUnidade = Get-PWDocumentsBySearch -Environment $environmentUnidade -Attributes @{NumeroPoderConcedente = $numeroDocumento; Revisao = $revisao; Versao = $versao }
        }

        Set-PWDocumentState -InputDocuments $documentoUnidade -State 'Emitido pela Engenharia' -Force
        $proximoState = CalculaState -stateAtual $stateAtual
        Set-PWDocumentState -InputDocuments $documento -State $proximoState -Force
    }
    else {
        $proximoState = CalculaState -stateAtual $stateAtual
        Set-PWDocumentState -InputDocuments $documento -State 'Arquivo interno Engenharia' -Force
    }
}
#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:Ecorodovias-01' -Password $SecurePassword -UserName 'admin'

$documentos = ObterDocumentosASeremEnviadosParaUnidade

foreach ($documento in $documentos) {
    $pastaDestino = CalcularPastaDestino -documento $documento
    movimentaDocumentosAnterioresSuperados -pastaDestino $pastaDestino -documento $documento
    copiaDocumento -documento $documento -pastaDestino $pastaDestino
}

Undo-PWLogin