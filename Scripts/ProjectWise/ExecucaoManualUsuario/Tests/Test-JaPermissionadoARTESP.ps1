# Teste do caminho real da funcao, com leituras PW simuladas e sem escrita nativa.
$ErrorActionPreference='Stop'
Import-Module pwps_dab -DisableNameChecking -ErrorAction Stop
. (Join-Path (Split-Path $PSScriptRoot) 'Funcoes\PreservarSegurancaARTESP.ps1')
$script:documentos=0;$script:acls=0
$script:doc=[pscustomobject]@{FileName='DE-teste-015-R00.pdf';ProjectID=10;DocumentID=20;WorkflowId=6;StateId=42;Workflow='Fluxo';WorkflowState='Estado'}
$script:acl=@([pscustomobject]@{MemberType=4;MemberId=0;AccessMask=0},[pscustomobject]@{MemberType=3;MemberId=1862;AccessMask=20},[pscustomobject]@{MemberType=2;MemberId=11448;AccessMask=20})
function Get-PWDocumentsByGUIDs {param($DocumentGUIDs,$ErrorAction) $script:documentos++;$script:doc}
function Ler-ACLWorkflowPropria {param($Documento) $script:acls++;$script:acl}
function Ler-ACLWorkflowOrigem {throw 'Nao deveria consultar heranca de ACL propria completa'}
$ctx=Novo-ContextoACLARTESP;$ctx.GrupoID=11448
$raiz=Join-Path $env:TEMP ('JaPermissionadoARTESP_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $raiz | Out-Null
try {
    $planejamento=Join-Path $raiz 'Plano'
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino $planejamento -DocumentoLido $script:doc -Contexto $ctx -Silencioso
    $script:documentos=0;$script:acls=0
    $aplicacao=Join-Path $raiz 'Aplicacao'
    Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino $aplicacao -PlanoAprovado (Join-Path $planejamento 'PlanoPreservacao.json') -Contexto $ctx -Aplicar -Silencioso
    if($script:documentos -ne 2 -or $script:acls -ne 2){throw 'Caminho rapido manteve consultas redundantes'}
    if($ctx.UltimoResultado -notlike 'Sem alteracao*'){throw 'Resultado incorreto'}
    if($ctx.Tempos.ContainsKey('EscreverPermissao') -or $ctx.Tempos.ContainsKey('ValidarACLAntes')){throw 'Entrou no caminho de escrita'}
    if(-not (Test-Path (Join-Path $aplicacao 'ResultadoSemAlteracao.json')) -or (Test-Path (Join-Path $aplicacao 'DocumentoAntes.xml'))){throw 'Arquivos redundantes ou resultado ausente'}
    $script:acl[1].AccessMask=4
    $bloqueou=$false
    try {Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino (Join-Path $raiz 'Mudou') -PlanoAprovado (Join-Path $planejamento 'PlanoPreservacao.json') -Contexto $ctx -Aplicar -Silencioso}
    catch {if($_.Exception.Message -like 'A ACL completa diverge*'){$bloqueou=$true}else{throw}}
    if(-not $bloqueou){throw 'Mudanca em outro participante foi ignorada'}
    'OK: ja permissionado sem escrita, duas leituras de identidade/ACL, sem arquivos duplicados e bloqueio de mudanca nos demais participantes.'
} finally {
    Get-ChildItem -LiteralPath $raiz -Recurse -File | Remove-Item
    Get-ChildItem -LiteralPath $raiz -Directory | Remove-Item
    Remove-Item -LiteralPath $raiz
}
