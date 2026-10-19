$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot) 'Funcoes\PreservarSegurancaARTESP.ps1')
$script:doc=[pscustomobject]@{FileName='DE-SP-903-R00.pdf';ProjectID=10;DocumentID=20;WorkflowId=6;StateId=42;Workflow='Fluxo';WorkflowState='Estado'}
$script:acl=@();$script:origem=@();$script:remocoes=0
function Get-PWDocumentsByGUIDs {param($DocumentGUIDs,$ErrorAction) $script:doc}
function Ler-ACLWorkflowPropria {param($Documento) $script:acl}
function Ler-ACLWorkflowOrigem {param($Documento,$Contexto,[switch]$Atualizar) $script:origem}
function Remover-EntradaWorkflowARTESP {
    param($Documento,$GrupoID)
    if($GrupoID -ne 11448 -or $Documento.WorkflowId -ne 6 -or $Documento.StateId -ne 42){throw 'Remocao fora de escopo'}
    $script:remocoes++
    $script:acl=@($script:acl|Where-Object {-not ($_.MemberType -eq 2 -and $_.MemberId -eq $GrupoID)})
    return $true
}
$artesp=[pscustomobject]@{MemberType=2;MemberId=11448;AccessMask=20}
$outro=[pscustomobject]@{MemberType=3;MemberId=1862;AccessMask=278}
$raiz=Join-Path $env:TEMP ('ReversaoARTESP_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $raiz|Out-Null
$ctx=Novo-ContextoACLARTESP;$ctx.GrupoID=11448
try{
    $script:acl=@($artesp,$outro)
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Plano') -Contexto $ctx -Reverter -Silencioso
    if($script:remocoes){throw 'Planejamento removeu permissao'}
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Aplicacao') -PlanoAprovado (Join-Path $raiz 'Plano\PlanoPreservacao.json') -Contexto $ctx -Reverter -Aplicar -Silencioso
    if($script:remocoes -ne 1 -or $script:acl.Count -ne 1 -or $script:acl[0].AccessMask -ne 278 -or $ctx.UltimoResultado -notlike 'Revertido*'){throw 'Reversao nao preservou outros direitos'}
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Repetir') -Contexto $ctx -Reverter -Aplicar -Silencioso
    if($script:remocoes -ne 1 -or $ctx.UltimoResultado -notlike 'Sem alteracao*'){throw 'Reversao nao idempotente'}
    $script:acl=@($artesp);$script:origem=@($outro)
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Ultima') -Contexto $ctx -Reverter -Aplicar -Silencioso
    if($script:acl.Count -ne 0){throw 'Ultima entrada nao removida'}
    $script:acl=@($artesp);$script:origem=@($artesp,$outro)
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Herdado') -Contexto $ctx -Reverter -Aplicar -Silencioso
    if($script:remocoes -ne 2 -or $ctx.UltimoResultado -notlike 'Nao revertido*'){throw 'Heranca nao sinalizada'}
    $script:doc.FileName='DE-SP-015-R00.pdf';$bloqueou=$false
    try{Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Invalido') -Contexto $ctx -Reverter -Aplicar}catch{if($_.Exception.Message -like 'Reversao bloqueada*'){$bloqueou=$true}else{throw}}
    if(-not $bloqueou){throw 'Reverteu arquivo fora das regras'}
    'OK: planejar sem remover, remover so ARTESP, preservar direitos, repetir sem escrita, ultima entrada, heranca e bloqueio fora das regras.'
}finally{
    Get-ChildItem -LiteralPath $raiz -Recurse -File|Remove-Item
    Get-ChildItem -LiteralPath $raiz -Directory|Remove-Item
    Remove-Item -LiteralPath $raiz
}
