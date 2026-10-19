$ErrorActionPreference='Stop'
$p=Join-Path (Split-Path $PSScriptRoot) '18 - Preparar Acesso ARTESP.ps1'
$t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($p,[ref]$t,[ref]$e)
if($e.Count){throw($e|Out-String)}
foreach($f in $ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Texto','Id-Pasta','Nome-Pasta','Escolher','Ler-DocumentosDisciplinas')},$true)){
    . ([scriptblock]::Create($f.Extent.Text))
}
$script:entrada='1,2'
function Read-Host {param($Prompt) $script:entrada}
$pastas=@([pscustomobject]@{ProjectID=10;Name='Geometria'},[pscustomobject]@{ProjectID=20;Name='Pavimento'})
if(@(Escolher $pastas 'Teste virgula' -Varios).Count -ne 2){throw 'Falha virgula'}
$script:entrada='1;2;2'
if(@(Escolher $pastas 'Teste ponto e virgula' -Varios).Count -ne 2){throw 'Falha ponto e virgula/duplicacao'}
$script:consultas=New-Object 'System.Collections.Generic.List[int]'
function Get-PWDocumentsBySearch {
    param($FolderID,[switch]$JustThisFolder,[switch]$PopulatePath,$ErrorAction)
    if(-not $JustThisFolder){throw 'Consulta recursiva inesperada'}
    $script:consultas.Add($FolderID)
    [pscustomobject]@{Name='mesmo.pdf';DocumentGUID="guid-$FolderID";ProjectID=$FolderID}
}
$r=@(Ler-DocumentosDisciplinas @($pastas[0],$pastas[1],$pastas[0]))
if($script:consultas.Count -ne 2 -or $r.Count -ne 2 -or $r[0].FolderID -ne 10 -or $r[1].FolderID -ne 20 -or $r[1].Pasta -ne 'Pavimento'){
    throw 'Escopo das pastas incorreto'
}
'OK: selecao multipla com virgula/ponto e virgula, pastas sem duplicacao e documentos vinculados a pasta correta.'
