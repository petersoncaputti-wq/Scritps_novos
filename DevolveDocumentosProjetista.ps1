function obterDocumentosASeremDevolvidosParaProjetista {
    $documentos = Get-PWDocumentsBySearch -SearchName 'Scripts\Docs p Projetista' -GetAttributes

    $documentosAgrupados = $documentos | Group-Object -Property @{Expression = {"GRD-COM-$($_.Attributes.Unidade)-$($_.Attributes.Projeto)"}}

    return $documentosAgrupados
}
#-----------------------------------------------------------------------------
# Funções para criação da pasta da GRD
#-----------------------------------------------------------------------------
function definirEnvironment {
    param ($poderConcedente)

    $environment = switch ($poderConcedente) {
        "ARTESP" { "dmsGRDRetornoARTESP" }
        "ANTT" { "dmsGRDRetornoANTT"}
        Default { "" }
    }
    return $environment
}
function criarPastaGRD {
    param ($prefixoGRD, $documento)
    
    if(-not $documento -or -not $documento.Attributes[0].PoderConcedente){
        return $null
    }

    $pastaGRDSaida = definePastaRaizGRDSaida -documento $documento
    if(-not $pastaGRDSaida){
        return $null
    }

    $numeroProximaGRD = definirNumeroProximaGRD -prefixoGRD $prefixoGRD -documento $documento -pastaGRDSaida $pastaGRDSaida
    $nomeNovaPasta = $pastaGRDSaida.FullPath + "\$numeroProximaGRD"
    $poderConcedente = $documento.Attributes[0].PoderConcedente
    $environment = definirEnvironment -poderConcedente $poderConcedente
    
    $novaPastaGRDSaida = New-PWFolder -FolderPath $nomeNovaPasta -Environment $environment

    return $novaPastaGRDSaida
}
function definePastaRaizGRDSaida {
    param ($documento)

    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $documento
    $nomePastaRaizGRD = $pastaRaizProjeto.FullPath+"\Area de Transferencia\Saida"
    $pastaRaizGRD = Get-PWFolders -FolderPath $nomePastaRaizGRD -JustOne
    return $pastaRaizGRD
}
function definirNumeroProximaGRD {
    param ($prefixoGRD, $documento, $pastaGRDSaida)

    $maiorGRD = obtemMaiorGRDEmitidaNoProjeto -prefixoGRD $prefixoGRD -pastaRaizGRDSaida $pastaGRDSaida
    if(-not $maiorGRD){
        return "$prefixoGRD-0000001"
    }
    $sequencial = $maiorGRD.Split('-')[-1]

    $proximoSequencial = ([int]$sequencial+1).ToString("D7")

    $numeroProximaGRD = "$prefixoGRD-$proximoSequencial"

    return $numeroProximaGRD
}
function obtemMaiorGRDEmitidaNoProjeto {
    param ($prefixoGRD, $pastaRaizGRDSaida)
    
    $query = "select Max(o_projectname) as 'GRD' from dms_proj where o_projectname like '$prefixoGRD%' and o_parentno = "+$pastaRaizGRDSaida.ProjectID

    $dados = Select-PWSQLDataTable -SQLSelectStatement $query
    return $dados.GRD
}
#-----------------------------------------------------------------------------
# Funções para atualização de atributos dos documentos
#-----------------------------------------------------------------------------
function defineProximoState {
    param ($stateAtual)

    $stateFinal = switch ($stateAtual) {
        "Em Envio para Projetista" { "Em revisao pelo Projetista" }
        "Enviado para Unidade com Ressalvas" {"Enviado para Unidade e Projetista com Ressalvas"}
        Default {stateAtual}
    }
    return $stateFinal
}
function atualizaStateDocumentos {
    param ($documentos)
    
    if(-not $documentos){
        return $null
    }

    $documentosAgrupados = $documentos | Group-Object -Property @{Expression = {"$($_.WorkflowState)"}}

    foreach ($grupo in $documentosAgrupados) {
        $proximoState = defineProximoState -stateAtual $grupo.Name
        Set-PWDocumentState -InputDocuments $grupo.Group -State $proximoState -Force
    }

    return $true
}
function preencheAtributosDaGRDSaida {
    param ($documentos, $pastaGRDSaida)

    if(-not $documentos -or -not $pastaGRDSaida ){
        return $null
    }

    $numeroGRDSaida = $pastaGRDSaida.Name
    $dataGRDSaida = [datetime]::UtcNow.AddHours(-3).ToString("yyyy-MM-dd HH:mm:ss")

    Update-PWDocumentAttributes -InputDocuments $documentos -Attributes @{NumeroGRDSaida = $numeroGRDSaida; DataGRDSaida = "$dataGRDSaida"}
    return $true
}
function atualizaAtributosEStateDocumentos {
    param ($documentos, $pastaGRDSaida)

    $resultado = preencheAtributosDaGRDSaida -documentos $documentos -pastaGRDSaida $pastaGRDSaida

    if(-not $resultado){
        return $null
    }

    $resultado = atualizaStateDocumentos -documentos $documentos
    return $resultado
}
#-----------------------------------------------------------------------------
# Funções para cópia dos documentos para GRD
#-----------------------------------------------------------------------------
function copiaDocumentosParaPastaGRD {
    param ($pastaGRDSaida, $documentos
    )

    $enderecoPasta = $pastaGRDSaida.FullPath
    $documentosASeremCopiados = $documentos.Where({$_.Attributes[0].ComentarioParaProjetista -eq '1'})
    Copy-PWDocumentsToFolder -InputDocument $documentosASeremCopiados -TargetFolderPath $enderecoPasta
    
}

#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' -Password $SecurePassword -UserName 'admin'

$documentos = obterDocumentosASeremDevolvidosParaProjetista

foreach ($documento in $documentos) {
    $pastaGRDSaida = criarPastaGRD -prefixoGRD $documento.Name -documento $documento.Group[0]

    atualizaAtributosEStateDocumentos -documentos $documento.Group -pastaGRDSaida $pastaGRDSaida
    copiaDocumentosParaPastaGRD -pastaGRDSaida $pastaGRDSaida -documentos $documento.Group
}


Undo-PWLogin