# Usa apenas conversao local do modulo. Nenhuma sessao ou escrita no ProjectWise.
$ErrorActionPreference='Stop'
Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop
. (Join-Path (Split-Path $PSScriptRoot) 'Funcoes\PreservarSegurancaARTESP.ps1')
$script:leiturasPasta=0;$script:leiturasSeguranca=0;$script:consultasDocumento=0;$script:mask='NA'
function Get-PWFolders {param($FolderID,[switch]$JustOne,$ErrorAction) $script:leiturasPasta++;[pscustomobject]@{ProjectID=$FolderID}}
function Get-PWFolderSecurity {
    param($InputFolder,$ErrorAction)
    $script:leiturasSeguranca++
    foreach($estado in @('A','B')){
        [pscustomobject]@{SecurityType='Document';Workflow='Fluxo';State=$estado;Type='All_Users (*Everyone)';Name='All_Users (*Everyone)';Access_Control_Settings=$script:mask;InheritingFrom='Origem'}
    }
}
function Get-PWDocumentsByGUIDs {param($DocumentGUIDs,$ErrorAction) $script:consultasDocumento++;$script:doc}
function Ler-ACLWorkflowPropria {param($Documento)}
$script:doc=[pscustomobject]@{FileName='DE-teste-015-R00.pdf';ProjectID=10;DocumentID=20;WorkflowId=6;StateId=42;Workflow='Fluxo';WorkflowState='A'}
$ctx=Novo-ContextoACLARTESP
$ctx.GrupoID=11448
$destino=Join-Path $env:TEMP ('CacheARTESP_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $destino | Out-Null
try {
    $a=@(Ler-ACLWorkflowOrigem $script:doc -Contexto $ctx)
    $b=@(Ler-ACLWorkflowOrigem $script:doc -Contexto $ctx)
    $script:doc.WorkflowState='B';$script:doc.StateId=43
    $c=@(Ler-ACLWorkflowOrigem $script:doc -Contexto $ctx)
    if($script:leiturasPasta -ne 1 -or $script:leiturasSeguranca -ne 1){throw 'Cache de pasta nao reutilizado entre estados'}
    $script:doc.WorkflowState='A';$script:doc.StateId=42
    $plano=Join-Path $destino 'Plano'
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino $plano -Contexto $ctx -DocumentoLido $script:doc -Silencioso
    if($script:consultasDocumento -ne 0){throw 'Documento pre-lido foi consultado novamente'}
    $script:mask='----r--R---'
    $bloqueou=$false
    try {
        Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $destino 'Aplicacao') -Contexto $ctx -PlanoAprovado (Join-Path $plano 'PlanoPreservacao.json') -Aplicar -Silencioso
    } catch {if($_.Exception.Message -like 'A ACL completa diverge*'){$bloqueou=$true}else{throw}}
    if(-not $bloqueou -or $script:leiturasSeguranca -ne 2){throw 'Mudanca na origem nao detectada antes da escrita'}
    'OK: uma leitura de pasta para multiplos estados, documento pre-lido reutilizado e origem alterada bloqueada antes da escrita.'
} finally {
    Get-ChildItem -LiteralPath $destino -Recurse -File | Remove-Item
    Get-ChildItem -LiteralPath $destino -Directory | Remove-Item
    Remove-Item -LiteralPath $destino
}
