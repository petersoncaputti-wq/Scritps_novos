$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot) 'Funcoes\PreservarSegurancaARTESP.ps1')
$casos=@(
    @('PL-SP-015-R00.pdf',$true),@('pl-SP-015-R00.dwg',$true),
    @('DE-SP-899-R00b.pdf',$false),@('DE-SP-900-R00b.pdf',$true),
    @('DE-SP-903-R00b.DWG',$true),@('DE-SP-1000-R00.pdf',$true),
    @('DE-SP-000900-R00.pdf',$true),@('DE-SP-000899-R00.pdf',$false),
    @('DE-SP-015-R900.pdf',$false),@('DE-SP-900-015-R01a.pdf',$false),
    @('DE-SP-015-R01a.pdf',$false),@('DE-SP-000-R00.pdf',$false)
)
foreach($caso in $casos){
    $bloqueado=-not [string]::IsNullOrWhiteSpace((Obter-BloqueioArquivoARTESP $caso[0]))
    if($bloqueado -ne $caso[1]){throw "Regra incorreta: $($caso[0])"}
}
function Get-PWDocumentsByGUIDs {param($DocumentGUIDs,$ErrorAction) [pscustomobject]@{FileName='DE-SP-903-R00.pdf'}}
$negou=$false
try {Atualizar-ACLWorkflowARTESP -GUID 'teste' -Destino 'nao-criar' -Aplicar}
catch {if($_.Exception.Message -like 'Arquivo bloqueado:*'){$negou=$true}else{throw}}
if(-not $negou){throw 'Bloqueio na aplicacao ausente'}
'OK: prefixo PL, limite 899/900, revisao distinta do numero e bloqueio antes da API de escrita.'
