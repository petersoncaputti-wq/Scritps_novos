#requires -Version 5.1
$ErrorActionPreference='Stop'
$base=Split-Path $PSScriptRoot -Parent
Add-Type -AssemblyName System.Web,System.IO.Compression.FileSystem,System.IO.Compression
Add-Type -Path (Join-Path $base 'Transfer.cs') -ReferencedAssemblies System.Runtime.Serialization,System.Xml
Add-Type -Path (Join-Path $PSScriptRoot 'Fixture.cs')
. (Join-Path $base 'Core.ps1')
. (Join-Path $base 'Projects.ps1')
$script:RunLog=$null
$temp=Join-Path $env:TEMP ('SharePointPS-test-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($temp)
$server=New-Object SPFixture
function Assert($Condition,[string]$Text){if(!$Condition){throw "FALHOU: $Text"};Write-Host "OK: $Text"}
function TransferTest([string]$Route,[string]$Target){
    $t=New-Object SharePointPS.Transfer
    $t.Start([string[]]@($server.BaseUrl+$Route),$Target,36,'identity-v1',(New-Object Net.CookieContainer))
    Assert ($t.Task.Wait(10000)) "transferencia encerra $Route"
    $t.Task.Result
}
function SeedPartial([string]$Target){
    [IO.File]::WriteAllBytes($Target+'.spdownload.part',$server.Data[0..9])
    SaveJson ($Target+'.spdownload.part.json') @{identity='identity-v1';size=36;etag='"v1"'}
}
try{
    $target=Join-Path $temp 'ok.bin';$r=TransferTest '/ok' $target
    Assert ($r.Success -and [IO.File]::ReadAllText($target) -ceq [Text.Encoding]::UTF8.GetString($server.Data)) 'streaming e conteudo completo'
    $deep=Join-Path (Join-Path (Join-Path $temp ('d'*80)) ('e'*80)) ('f'*80)
    [SharePointPS.LocalPaths]::EnsureDirectory($deep)
    $target=Join-Path $deep 'documento.xlsx';$r=TransferTest '/ok' $target
    Assert ($target.Length -gt 300 -and $r.Success -and [IO.File]::ReadAllText([SharePointPS.LocalPaths]::Native($target)) -ceq [Text.Encoding]::UTF8.GetString($server.Data)) 'caminho maior que 300: documento e metadados gravados'
    $target=Join-Path $deep 'retomada.xlsx'
    [IO.File]::WriteAllBytes([SharePointPS.LocalPaths]::Native($target+'.spdownload.part'),$server.Data[0..9])
    SaveJson ($target+'.spdownload.part.json') @{identity='identity-v1';size=36;etag='"v1"'}
    $r=TransferTest '/range' $target
    Assert ($r.Success -and $server.LastRange['/range'] -eq 'bytes=10-') 'retomada e renomeacao em caminho longo'
    $target=Join-Path $temp 'range.bin';SeedPartial $target;$r=TransferTest '/range' $target
    Assert ($r.Success -and $server.LastRange['/range'] -eq 'bytes=10-' -and $server.LastIfRange['/range'] -eq '"v1"') 'retomada Range e If-Range'
    $target=Join-Path $temp 'ignore.bin';SeedPartial $target;$r=TransferTest '/ignore' $target
    Assert ($r.Success -and (Get-Item -LiteralPath $target).Length -eq 36) 'servidor ignora Range: reinicia sem duplicar bytes'
    $target=Join-Path $temp 'bad.bin';SeedPartial $target;$r=TransferTest '/bad-range' $target
    Assert (!$r.Success -and (Get-Item -LiteralPath ($target+'.spdownload.part')).Length -eq 10) 'Content-Range invalido preserva parcial'
    $target=Join-Path $temp 'broken.bin';$r=TransferTest '/broken' $target
    Assert (!$r.Success -and (Test-Path -LiteralPath ($target+'.spdownload.part'))) 'transferencia truncada preserva parcial'
    $r=TransferTest '/broken' $target;Assert $r.Success 'nova tentativa apos transferencia truncada'
    $mismatch=New-Object SharePointPS.Transfer
    $mismatch.Start([string[]]@($server.BaseUrl+'/ok'),(Join-Path $temp 'mismatch.bin'),10,'old',(New-Object Net.CookieContainer));[void]$mismatch.Task.Wait(10000)
    Assert ($mismatch.Task.Result.SizeMismatch -and $mismatch.Task.Result.Expected -eq 10 -and $mismatch.Task.Result.Received -eq 36) 'divergencia registra bytes esperados e recebidos e nao publica documento'
    $target=Join-Path $temp 'fallback.bin';$t=New-Object SharePointPS.Transfer
    $t.Start([string[]]@(($server.BaseUrl+'/missing'),($server.BaseUrl+'/ok')),$target,36,'id',(New-Object Net.CookieContainer));[void]$t.Task.Wait(10000)
    Assert ($t.Task.Result.Success -and $t.Task.Result.Route -eq '1') 'fallback apenas apos 404'
    $target=Join-Path $temp 'auth.bin';$r=TransferTest '/auth' $target;Assert ($r.Authentication -and !(Test-Path -LiteralPath $target)) '401 identificado sem salvar documento'
    $before=$server.Counts['/ok'];$t=New-Object SharePointPS.Transfer
    $t.Start([string[]]@(($server.BaseUrl+'/auth'),($server.BaseUrl+'/ok')),(Join-Path $temp 'no-fallback.bin'),36,'id',(New-Object Net.CookieContainer));[void]$t.Task.Wait(10000)
    Assert (!$t.Task.Result.Success -and $server.Counts['/ok'] -eq $before) 'falha de login nao aciona fallback'
    $t=New-Object SharePointPS.Transfer;$cancelTarget=Join-Path $temp 'cancel.bin'
    $t.Start([string[]]@($server.BaseUrl+'/slow'),$cancelTarget,36,'id',(New-Object Net.CookieContainer));$t.Cancel=$true;[void]$t.Task.Wait(10000)
    Assert (!$t.Task.Result.Success -and !(Test-Path -LiteralPath $cancelTarget)) 'cancelamento nao publica arquivo incompleto'
    $target=Join-Path $temp 'throttle.bin';$r=TransferTest '/throttle' $target;Assert ($r.Status -eq 429 -and $r.RetryAfter -eq 0) '429 e Retry-After'
    $target=Join-Path $temp 'html.bin';$r=TransferTest '/html' $target;Assert ($r.Authentication -and !(Test-Path -LiteralPath $target)) 'HTML inesperado rejeitado'
    $target=Join-Path $temp 'orphan.bin';[IO.File]::WriteAllText($target+'.spdownload.part','keep');$r=TransferTest '/ok' $target
    Assert (!$r.Success -and [IO.File]::ReadAllText($target+'.spdownload.part') -eq 'keep') 'parcial sem identificacao preservado'
    $parsed=ParseUrl 'https://demo.sharepoint.com/:f:/r/sites/Test/Docs/Pasta?e=abc'
    Assert ($parsed.Path -eq '/sites/Test/Docs/Pasta') 'link /:f:/r normalizado'
    Assert ((ParseUrl 'https://demo.sharepoint.com/sites/Test/Docs/Forms/AllItems.aspx?id=%2Fsites%2FTest%2FDocs%2FA%23B').Path -eq '/sites/Test/Docs/A#B') 'AllItems com id e caracteres especiais'
    Assert ((LocalName 'CON.pdf') -notmatch '^CON\.') 'nomes reservados adaptados'
    $rejected=$false;try{SafePath (Join-Path $temp '..\outside') $temp}catch{$rejected=$true};Assert $rejected 'escape do destino rejeitado'
    $clean=[SharePointPS.Transfer]::Clean('token="secret value" https://example.com/?sig=secret user@example.com Bearer secret')
    Assert ($clean -notmatch 'secret|user@example|https://') 'log remove credenciais, URL e email'
    $script:Web='https://demo.sharepoint.com/sites/Test'
    Assert ((RemoteEndpoint File "/sites/Test/Docs/O'Brien #1.pdf") -like '*O%27%27Brien%20%231.pdf*') 'encoding REST com apostrofo e cerquilha'
    # Fixtures XLSX sem Excel, formulas ou Python.
    $xlsx=Join-Path $temp 'projects.xlsx';$archive=[IO.Compression.ZipFile]::Open($xlsx,[IO.Compression.ZipArchiveMode]::Create)
    $parts=@{
        'xl/workbook.xml'='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Lista" sheetId="1" r:id="r1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels'='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/_rels/sheet1.xml.rels'='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h1" TargetMode="External" Target="https://demo.sharepoint.com/:f:/r/sites/Test/Docs/A"/></Relationships>'
        'xl/worksheets/sheet1.xml'='<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="2"><c r="C2" t="inlineStr"><is><t>SH1</t></is></c><c r="E2"><v>2</v></c><c r="F2" t="inlineStr"><is><t>Exibir</t></is></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>SH2</t></is></c><c r="E3"><v>1</v></c><c r="F3" t="inlineStr"><is><t>https://demo.sharepoint.com/sites/Test/Docs/B</t></is></c></row><row r="4"><c r="C4" t="inlineStr"><is><t>SH3</t></is></c><c r="E4"><v>1</v></c><c r="F4" t="inlineStr"><is><t>L17</t></is></c></row></sheetData><hyperlinks><hyperlink ref="F2" r:id="h1"/></hyperlinks></worksheet>'
    }
    try{foreach($name in $parts.Keys){$entry=$archive.CreateEntry($name);$writer=[IO.StreamWriter]::new($entry.Open());try{$writer.Write($parts[$name])}finally{$writer.Dispose()}}}finally{$archive.Dispose()}
    $queue=LoadProjects $xlsx
    Assert ($queue.Projects.Count -eq 2 -and $queue.Projects[0].Row -eq 3 -and $queue.Issues.Count -eq 1) 'XLSX: prioridade, hyperlink real e pendencia'
    $ProjectRows=@(2);$queue=LoadProjects $xlsx;Assert ($queue.Projects.Count -eq 1 -and $queue.Projects[0].Row -eq 2) 'filtro de linhas XLSX';$ProjectRows=$null
    # Coordenador real com HTTP local: concorrencia, skip, conflito e nova tentativa.
    function CookieJar {return ,(New-Object Net.CookieContainer)}
    function Routes([string]$Remote){$server.BaseUrl+$Remote}
    $script:OriginUri=[uri]'https://demo.sharepoint.com';$script:RemoteRoot='/';$Workers=3
    $root=Join-Path $temp 'batch';$state=Join-Path $temp 'state';[void][IO.Directory]::CreateDirectory($root);[void][IO.Directory]::CreateDirectory($state)
    $inventory=Join-Path $temp 'inventory.jsonl';$rows=@()
    foreach($remote in @('/slow','/ok','/retry-throttle')){$rows+=(@{Remote=$remote;Local=(Join-Path $root ($remote.TrimStart('/')+'.bin'));Size=36;Version='v1'}|ConvertTo-Json -Compress)}
    $longLocal=Join-Path (Join-Path (Join-Path (Join-Path $root ('g'*80)) ('h'*80)) ('i'*80)) 'coordenador.xlsx'
    $rows+=(@{Remote='/long';Local=$longLocal;Size=36;Version='v1'}|ConvertTo-Json -Compress)
    [IO.File]::WriteAllLines($inventory,[string[]]$rows)
    $totals=[pscustomobject]@{Count=4;Bytes=144}
    $r=RunTransfers $inventory $state $root $totals;Assert ($r.Downloaded -eq 4 -and $r.Failed -eq 0 -and $server.Counts['/retry-throttle'] -eq 2) 'coordenador paralelo, caminho longo, limitacao 429 e nova tentativa'
    $r=RunTransfers $inventory $state $root $totals;Assert ($r.Skipped -eq 4 -and $r.Downloaded -eq 0) 'reinicio reaproveita arquivos existentes inclusive caminho longo'
    [IO.File]::WriteAllText((Join-Path $root 'ok.bin'),'conflict')
    $r=RunTransfers $inventory $state $root $totals;Assert ($r.Failed -eq 1 -and $r.Skipped -eq 3) 'conflito preservado e lote continua'
    function BrowserJson {param($Endpoint) [pscustomobject]@{Length='36';ETag='v2'}}
    $changedRoot=Join-Path $temp 'changed';$changedState=Join-Path $temp 'changed-state'
    [void][IO.Directory]::CreateDirectory($changedRoot);[void][IO.Directory]::CreateDirectory($changedState)
    $changedInventory=Join-Path $temp 'changed-inventory.jsonl'
    [IO.File]::WriteAllText($changedInventory,(@{Remote='/changed';Local=(Join-Path $changedRoot 'changed.bin');Size=10;Version='v1'}|ConvertTo-Json -Compress))
    $changedTotals=[pscustomobject]@{Count=1;Bytes=10}
    $changedResult=RunTransfers $changedInventory $changedState $changedRoot $changedTotals
    Assert ($changedResult.Downloaded -eq 1 -and $changedResult.Failed -eq 0 -and $changedTotals.Bytes -eq 36) 'origem alterada: revalida metadados e baixa novamente com tamanho correto'
    $manifest=Get-Content -LiteralPath (Join-Path $temp 'manifest.json') -Raw|ConvertFrom-Json
    Assert (!$manifest.Complete) 'alteracao na origem invalida inventario antigo para proxima execucao'
    function BrowserJson {param($Endpoint) [pscustomobject]@{Name='Docs';ServerRelativeUrl='/Docs'}}
    $script:ListingCalls=0
    function Collection {param($Endpoint)
        $script:ListingCalls++
        if($Endpoint -like '*/Folders?*' -and $Endpoint -match "decodedUrl='%2FDocs'"){
            [pscustomobject]@{Name='Sub';ServerRelativeUrl='/Docs/Sub'}
            [pscustomobject]@{Name='Empty';ServerRelativeUrl='/Docs/Empty'}
        }
        if($Endpoint -like '*/Files?*' -and $Endpoint -match "decodedUrl='%2FDocs%2FSub'"){
            [pscustomobject]@{Name='a.txt';ServerRelativeUrl='/Docs/Sub/a.txt';Length=36;ETag='v1'}
        }
    }
    $script:Web='https://demo.sharepoint.com';$ListOnly=$true;$Recursive=$true;$OnlyCurrentFolder=$false;$RefreshInventory=$false
    $parsed=ParseUrl 'https://demo.sharepoint.com/Docs';$dest=Join-Path $temp 'inventory-cache'
    $r=RunBatch $parsed $dest;$first=$script:ListingCalls
    Assert ($first -eq 6 -and (Test-Path -LiteralPath (Join-Path $dest 'Docs\Empty'))) 'inventario recursivo inclui pasta vazia'
    $r=RunBatch $parsed $dest;Assert ($script:ListingCalls -eq $first) 'inventario completo reutilizado sem reler listagem'
    $RefreshInventory=$true;$r=RunBatch $parsed $dest;Assert ($script:ListingCalls -eq 2*$first) 'RefreshInventory rele origem e substitui inventario atomicamente'
    $RefreshInventory=$false;$OnlyCurrentFolder=$true;$r=RunBatch $parsed $dest
    Assert ($script:ListingCalls -eq 2*$first+1) 'mudanca do escopo invalida cache recursivo'
    $lock=[IO.File]::Open((Join-Path $dest '.sharepoint_download.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $rejected=$false;try{RunBatch $parsed $dest}catch{$rejected=$true}finally{$lock.Dispose()};Assert $rejected 'bloqueio de destino impede lotes simultaneos'
    $script:OpenFolderCalls=0
    function OpenFolder {param($Parsed) $script:OpenFolderCalls++}
    function RunBatch {param($Parsed,$Destination) if($Parsed.Path -eq '/fail'){throw 'Falha simulada'};[pscustomobject]@{Failed=0}}
    $queueTarget=Join-Path $temp 'queue'
    $queue=[pscustomobject]@{Issues=@();Projects=@(
        [pscustomobject]@{Row=2;Priority=1;Folder='first';Parsed=[pscustomobject]@{Path='/fail'}},
        [pscustomobject]@{Row=3;Priority=2;Folder='second';Parsed=[pscustomobject]@{Path='/ok'}}
    )}
    $code=RunQueue $queue $queueTarget
    $report=Get-ChildItem -LiteralPath $queueTarget -Filter '*.json'|Select-Object -First 1|Get-Content -Raw|ConvertFrom-Json
    Assert ($code -eq 1 -and $report.projetos[0].estado -eq 'falha' -and $report.projetos[1].estado -eq 'concluido') 'fila continua apos falha e persiste resumo por projeto'
    Assert ($script:OpenFolderCalls -eq 2) 'fila navega para cada projeto'
    $script:OpenFolderCalls=0
    $code=RunQueue $queue (Join-Path $temp 'queue-first-opened') -FirstFolderOpened
    Assert ($script:OpenFolderCalls -eq 1) 'primeira pasta autenticada nao repete navegacao; proximo projeto abre normalmente'
    Write-Host "TODOS OS TESTES PASSARAM. Fixtures preservadas em: $temp"
}finally{$server.Dispose()}
