# Testes isolados; nao importa PWPS_DAB nem acessa ProjectWise.
$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot) 'Funcoes\PreservarSegurancaARTESP.ps1')
$script:chamadas=New-Object 'System.Collections.Generic.List[string]'
$script:resposta='APLICAR LOTE'
$script:falhar=''
function Read-Host {param($Prompt) $script:resposta}
function Get-PWDocumentsByGUIDs {
    param($DocumentGUIDs,$ErrorAction)
    [pscustomobject]@{ProjectID=10;Workflow='Fluxo';WorkflowState='Estado'}
}
function Atualizar-ACLWorkflowARTESP {
    param($GUID,$Destino,[switch]$Aplicar,$PlanoAprovado)
    if($Aplicar){
        if(-not $PlanoAprovado){throw 'Aplicacao sem plano aprovado'}
        $script:chamadas.Add('Aplicar:'+ $GUID)
        if($GUID -eq $script:falhar){throw 'Falha simulada'}
    } else {$script:chamadas.Add('Planejar:'+ $GUID)}
}
$linhas=@(1..3 | ForEach-Object {[pscustomobject]@{Documento="arquivo$_.pdf";DocumentGUID="guid$_";FolderID=10;Workflow='Fluxo';Estado='Estado'}})
$raiz=Join-Path $env:TEMP ('TestLoteARTESP_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $raiz | Out-Null
try {
    $destino=Join-Path $raiz 'Sucesso'
    Executar-LoteARTESP $linhas $destino
    $r=@(Import-Csv (Join-Path $destino 'ResultadoDocumentos.csv'))
    if(@($r | Where-Object Resultado -eq 'Concluido - ACL completa verificada').Count -ne 3){throw 'Lote incompleto'}
    if(($script:chamadas -join ',') -ne 'Planejar:guid1,Planejar:guid2,Planejar:guid3,Aplicar:guid1,Aplicar:guid2,Aplicar:guid3'){throw 'Ordem incorreta'}
    $script:chamadas.Clear();$script:resposta=''
    $destino=Join-Path $raiz 'Cancelado'
    Executar-LoteARTESP $linhas $destino
    if(@($script:chamadas | Where-Object {$_ -like 'Aplicar:*'}).Count){throw 'Escreveu sem confirmacao'}
    $script:chamadas.Clear();$script:resposta='APLICAR LOTE';$script:falhar='guid2'
    $destino=Join-Path $raiz 'Falha'
    $falhou=$false
    try{Executar-LoteARTESP $linhas $destino}catch{if($_.Exception.Message -eq 'Falha simulada'){$falhou=$true}else{throw}}
    if(-not $falhou -or $script:chamadas.Contains('Aplicar:guid3')){throw 'Continuou depois de falha'}
    $r=@(Import-Csv (Join-Path $destino 'ResultadoDocumentos.csv'))
    if($r[0].Resultado -ne 'Concluido - ACL completa verificada' -or $r[1].Resultado -notlike 'Falha*' -or $r[2].Resultado -ne 'Nao executado - lote interrompido'){throw 'Relatorio incorreto'}
    'OK: lote completo, planejamento antes da escrita, confirmacao obrigatoria, interrupcao e relatorio de pendentes.'
}
finally {
    # Somente arquivos gerados neste diretorio temporario de teste conhecido.
    Get-ChildItem -LiteralPath $raiz -Recurse -File | Remove-Item
    Get-ChildItem -LiteralPath $raiz -Directory | Remove-Item
    Remove-Item -LiteralPath $raiz
}
