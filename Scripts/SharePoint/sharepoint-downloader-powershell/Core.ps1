function HashText([string]$Text) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}
function LocalName([string]$Name) {
    $clean=([regex]::Replace($Name,'[<>:"/\\|?*\x00-\x1f]','_')).TrimEnd(' ','.')
    if (!$clean -or $clean -in @('.','..')) { $clean='_' }
    if ($clean -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $clean='_'+$clean }
    if ($clean -cne $Name -or $clean.Length -gt 120) {
        $ext=[IO.Path]::GetExtension($clean); if($ext.Length -gt 20){$ext=$ext.Substring(0,20)}
        $clean=$clean.Substring(0,[Math]::Min(80,$clean.Length))+'_'+(HashText $Name).Substring(0,12)+$ext
    }
    $clean
}
function SafePath([string]$Path, [string]$Root) {
    $full=[IO.Path]::GetFullPath($Path); $base=[IO.Path]::GetFullPath($Root).TrimEnd('\')
    if ($full -ine $base -and !$full.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Caminho local fora do destino.' }
    $node=$full
    while ($node) {
        if ([SharePointPS.LocalPaths]::Exists($node)) {
            if ([IO.File]::GetAttributes([SharePointPS.LocalPaths]::Native($node)) -band [IO.FileAttributes]::ReparsePoint) { throw 'Destino contem junction/link; use uma pasta local sem redirecionamento.' }
        }
        $parent=[IO.Path]::GetDirectoryName($node)
        if($parent -eq $node){break}; $node=$parent
    }
    $full
}
function SaveJson([string]$Path, $Value) {
    $Path=[SharePointPS.LocalPaths]::Native($Path)
    $temp=$Path+'.tmp'
    [IO.File]::WriteAllText($temp,($Value | ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
    if([IO.File]::Exists($Path)){[IO.File]::Replace($temp,$Path,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temp,$Path)}
}
function Log([string]$Event, $Data=@{}) {
    if(!$script:RunLog){return}
    $safe=@{}
    foreach($key in $Data.Keys) {
        $value=$Data[$key]
        if($value -is [string]){ $value=[SharePointPS.Transfer]::Clean($value) }
        $safe[$key]=$value
    }
    $line=(@{hora=[DateTimeOffset]::Now.ToString('o'); evento=$Event; dados=$safe} | ConvertTo-Json -Compress -Depth 10)
    [IO.File]::AppendAllText($script:RunLog,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    $history=Join-Path $PSScriptRoot 'logs\sharepoint_downloader.log'
    if((Test-Path -LiteralPath $history) -and (Get-Item -LiteralPath $history).Length -ge 5MB) {
        for($n=3;$n -ge 1;$n--){$from=if($n -eq 1){$history}else{"$history.$($n-1)"}; if(Test-Path -LiteralPath $from){Move-Item -LiteralPath $from -Destination "$history.$n" -Force}}
    }
    [IO.File]::AppendAllText($history,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
}
function ParseUrl([string]$Text) {
    $u=[uri]$Text.Trim()
    if(!$u.IsAbsoluteUri -or $u.Scheme -ne 'https' -or $u.Host -notlike '*.sharepoint.com' -or $u.UserInfo -or !$u.IsDefaultPort){throw 'Use uma URL HTTPS de SharePoint Online.'}
    $q=[Web.HttpUtility]::ParseQueryString($u.Query)
    $path=$q['id']; if(!$path){$path=$q['RootFolder']}; if(!$path){$path=[uri]::UnescapeDataString($u.AbsolutePath)}
    if($path.StartsWith('/:f:/r/')){$path=$path.Substring(6)}
    if(!$path.StartsWith('/') -or $path.StartsWith('//') -or $path.StartsWith('/:') -or $path -match '/Forms/'){throw 'Use URL de pasta ou AllItems.aspx com id/RootFolder; link opaco nao suportado.'}
    $path=$path.TrimEnd('/'); if(!$path){throw 'Informe uma pasta, nao apenas o dominio.'}
    $encoded=($path.Split('/') | ForEach-Object{[uri]::EscapeDataString($_)}) -join '/'
    [pscustomobject]@{Host=$u.Host; Path=$path; Url='https://'+$u.Host+$encoded}
}
function Driver([string]$Method,[string]$Route,$Body=$null) {
    $options=@{Uri="$script:DriverBase/$Route";Method=$Method;TimeoutSec=150;UseBasicParsing=$true}
    if($null -ne $Body){$options.ContentType='application/json; charset=utf-8';$options.Body=[Text.Encoding]::UTF8.GetBytes(($Body|ConvertTo-Json -Depth 20 -Compress))}
    try {(Invoke-RestMethod @options).value} catch {throw "Falha de comunicacao com Edge WebDriver ($Method $($Route.Split('/')[-1])). Verifique versoes do Edge/driver e se a janela esta aberta."}
}
function SelectTab {
    foreach($handle in @(Driver GET "session/$script:SessionId/window/handles")){
        $null=Driver POST "session/$script:SessionId/window" @{handle=$handle}
        $current=[uri](Driver GET "session/$script:SessionId/url")
        if($current.Host -eq $script:OriginUri.Host){return}
    }
    throw 'Nenhuma aba aberta no SharePoint solicitado.'
}
function Renew {
    Log 'LOGIN_RENOVACAO'
    $null=Driver POST "session/$script:SessionId/url" @{url=$script:LoginUrl}
    [void](Read-Host 'Conclua login/MFA no Edge e confira acesso a pasta; ENTER para continuar')
    SelectTab
}
function BrowserJson([string]$Endpoint) {
    $u=[uri]$Endpoint
    if($u.Scheme -ne 'https' -or $u.Host -ne $script:OriginUri.Host -or $u.UserInfo -or !$u.IsDefaultPort){throw 'Endpoint fora da origem solicitada.'}
    $js=@'
var done=arguments[arguments.length-1];
fetch(arguments[0],{credentials:'same-origin',headers:{Accept:'application/json;odata=nometadata'}})
.then(async r=>{if(!r.ok)return done({status:r.status,retry:r.headers.get('Retry-After')});
if(!(r.headers.get('content-type')||'').includes('json'))return done({status:401});
done({status:r.status,data:await r.json()});}).catch(()=>done({status:0}));
'@
    for($attempt=1;$attempt -le 4;$attempt++){
        $r=Driver POST "session/$script:SessionId/execute/async" @{script=$js;args=@($Endpoint)}
        if($null -ne $r.data){return $r.data}
        if($attempt -eq 4){break}
        if($r.status -in @(401,403)){Renew;continue}
        if($r.status -in @(0,429,503,500,502,504)){WaitSeconds (RetryDelay $r.retry ($attempt*3));continue}
        throw "Leitura SharePoint falhou: HTTP $($r.status)."
    }
    throw "Leitura SharePoint falhou: HTTP $($r.status)."
}
function RetryDelay($Value,[double]$Default=3) {
    $n=0.0; $dt=[DateTimeOffset]::MinValue
    if([double]::TryParse([string]$Value,[ref]$n)){return [Math]::Max(0,$n)}
    if([DateTimeOffset]::TryParse([string]$Value,[ref]$dt)){return [Math]::Max(0,($dt-[DateTimeOffset]::UtcNow).TotalSeconds)}
    $Default
}
function WaitSeconds([double]$Seconds){$until=[DateTime]::UtcNow.AddSeconds($Seconds); while([DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 200}}
function Collection([string]$Endpoint){
    $seen=@{}
    while($Endpoint){
        if($seen.ContainsKey($Endpoint)){throw 'Paginacao circular.'};$seen[$Endpoint]=$true
        $data=BrowserJson $Endpoint
        foreach($item in $data.value){$item}
        $next=$data.'@odata.nextLink';if(!$next){$next=$data.'odata.nextLink'}
        if($next){$Endpoint=[uri]::new([uri]$Endpoint,[string]$next).AbsoluteUri}else{$Endpoint=$null}
    }
}
function RemoteEndpoint([string]$Kind,[string]$Path){
    $literal=[uri]::EscapeDataString($Path.Replace("'","''")).Replace("'",'%27')
    "$script:Web/_api/web/Get${Kind}ByServerRelativePath(decodedUrl='$literal')"
}
function Routes([string]$Remote){
    (RemoteEndpoint File $Remote)+'/$value'
    'https://'+$script:OriginUri.Host+(($Remote.Split('/')|ForEach-Object{[uri]::EscapeDataString($_)}) -join '/')
    $script:Web+'/_layouts/15/download.aspx?SourceUrl='+[uri]::EscapeDataString($Remote)
}
function CookieJar {
    $jar=New-Object Net.CookieContainer
    foreach($c in @(Driver GET "session/$script:SessionId/cookie")){
        $domain=$c.domain.TrimStart('.')
        if($script:OriginUri.Host -eq $domain -or $script:OriginUri.Host.EndsWith('.'+$domain)){
            $cookie=New-Object Net.Cookie($c.name,$c.value,$c.path,$c.domain);$cookie.Secure=$c.secure;$jar.Add($cookie)
        }
    }
    return ,$jar
}
function OpenFolder($Parsed) {
    $script:OriginUri=[uri]$Parsed.Url; $script:LoginUrl=$Parsed.Url
    Write-Host "Abrindo pasta no Edge: $($Parsed.Path)"
    Log 'NAVEGACAO_INICIO' @{host=$Parsed.Host;pasta=$Parsed.Path}
    $null=Driver POST "session/$script:SessionId/url" @{url=$Parsed.Url}
    Log 'NAVEGACAO_FIM' @{host=$Parsed.Host;pasta=$Parsed.Path}
    $current=[uri](Driver GET "session/$script:SessionId/url")
    if(!$script:AuthenticatedHost -or $current.Host -ne $Parsed.Host -or $script:AuthenticatedHost -ne $Parsed.Host){
        [void](Read-Host 'Faca login/MFA no Edge. Quando a pasta aparecer, pressione ENTER aqui'); SelectTab
    }
    $script:AuthenticatedHost=$Parsed.Host
    $script:Web=$null; $segments=$Parsed.Path.Trim('/').Split('/')
    for($n=$segments.Count;$n -ge 0;$n--){
        $candidate='https://'+$Parsed.Host
        if($n -gt 0){$candidate+='/'+(($segments[0..($n-1)]|ForEach-Object{[uri]::EscapeDataString($_)}) -join '/')}
        try{$info=BrowserJson ($candidate+'/_api/web?$select=Url,Title');$script:Web=$info.Url.TrimEnd('/');break}
        catch{if($_.Exception.Message -notmatch 'HTTP 404'){throw}}
    }
    if(!$script:Web -or ([uri]$script:Web).Host -ne $Parsed.Host){throw 'Site SharePoint nao localizado.'}
    Write-Host "Site: $script:Web";Log 'SITE_LOCALIZADO' @{host=$Parsed.Host;pasta=$Parsed.Path}
}
function SelectDestination {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog=New-Object Windows.Forms.FolderBrowserDialog
        try{$dialog.Description='Selecione o destino local dos downloads SharePoint';if($dialog.ShowDialog() -eq 'OK'){return $dialog.SelectedPath}}finally{$dialog.Dispose()}
    }catch{Write-Host 'Seletor visual indisponivel.'}
    $text=Read-Host 'Pasta local de destino (ENTER para cancelar)'
    if(!$text){throw 'Selecao de destino cancelada.'};$text
}
function BytesText([double]$Value){
    $units=@('B','KiB','MiB','GiB','TiB');$n=0
    while($Value -ge 1024 -and $n -lt 4){$Value/=1024;$n++}
    '{0:N1} {1}' -f $Value,$units[$n]
}
function BuildInventory([string]$Remote,[string]$Root,[string]$Inventory,[string]$Manifest,[string]$Scope){
    $temp=$Inventory+'.tmp';$writer=[IO.StreamWriter]::new($temp,$false,[Text.UTF8Encoding]::new($false))
    $stack=New-Object Collections.Stack;$stack.Push(@($Remote,$Root));$visited=@{};$targets=@{};$count=0;$bytes=0L
    try{
        while($stack.Count){
            $entry=$stack.Pop();$path=[string]$entry[0];$local=SafePath $entry[1] $Root
            if($visited.ContainsKey($path)){continue};$visited[$path]=$true
            [SharePointPS.LocalPaths]::EnsureDirectory($local)
            Write-Host "Inventariando: $path";Log 'PASTA_LIDA' @{remoto=$path;local=$local}
            $endpoint=RemoteEndpoint Folder $path
            foreach($file in @(Collection ($endpoint+'/Files?$select=Name,Length,ServerRelativeUrl,ETag,TimeLastModified&$top=500'))){
                if(!$file.ServerRelativeUrl.StartsWith($path+'/',[StringComparison]::OrdinalIgnoreCase)){throw 'Arquivo fora da pasta.'}
                $target=SafePath (Join-Path $local (LocalName $file.Name)) $Root
                if($targets.ContainsKey($target)){throw 'Colisao entre nomes locais de arquivos/pastas.'};$targets[$target]=$true
                $version=[string]$file.ETag;if(!$version){$version=[string]$file.TimeLastModified}
                $row=@{Remote=$file.ServerRelativeUrl;Local=$target;Size=[long]$file.Length;Version=$version}
                $writer.WriteLine(($row|ConvertTo-Json -Compress));$count++;$bytes+=[long]$file.Length
                Log 'ARQUIVO_INVENTARIADO' $row
            }
            if($Scope -eq 'recursive'){
                foreach($folder in @(Collection ($endpoint+'/Folders?$select=Name,ServerRelativeUrl&$top=500'))){
                    if(!$folder.ServerRelativeUrl.StartsWith($path+'/',[StringComparison]::OrdinalIgnoreCase)){throw 'Subpasta fora da origem.'}
                    $child=SafePath (Join-Path $local (LocalName $folder.Name)) $Root
                    if($targets.ContainsKey($child)){throw 'Colisao entre nomes locais de arquivos/pastas.'};$targets[$child]=$true
                    $stack.Push(@($folder.ServerRelativeUrl,$child))
                }
            }
            Write-Host ("Inventario: {0} pastas | {1} arquivos | {2}" -f $visited.Count,$count,(BytesText $bytes))
        }
    }finally{$writer.Dispose()}
    if([IO.File]::Exists($Inventory)){[IO.File]::Replace($temp,$Inventory,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temp,$Inventory)}
    $manifestData=@{Complete=$true;Remote=$Remote;Root=$Root;Scope=$Scope;Count=$count;Bytes=$bytes;Folders=$visited.Count}
    SaveJson $Manifest $manifestData
    return [pscustomobject]$manifestData
}
function RunTransfers([string]$Inventory,[string]$State,[string]$Root,$Totals){
    $reader=[IO.StreamReader]::new($Inventory);$jobs=New-Object Collections.ArrayList;$retry=New-Object Collections.ArrayList
    $downloaded=0;$skipped=0;$failed=0;$doneBytes=0L;$limit=$Workers;$cooldown=[DateTime]::MinValue;$eof=$false
    $watch=[Diagnostics.Stopwatch]::StartNew();$lastLog=0.0
    try{
        while(!$eof -or $jobs.Count -or $retry.Count){
            while($jobs.Count -lt $limit -and [DateTime]::UtcNow -ge $cooldown){
                $ready=-1
                for($i=0;$i -lt $retry.Count;$i++){if($retry[$i].At -le [DateTime]::UtcNow){$ready=$i;break}}
                if($ready -ge 0){$job=$retry[$ready];$retry.RemoveAt($ready)}
                elseif(!$eof){
                    $line=$reader.ReadLine();if($null -eq $line){$eof=$true;continue}
                    $row=$line|ConvertFrom-Json;$target=SafePath $row.Local $Root
                    $null=SafePath ($target+'.spdownload.part') $Root
                    $null=SafePath ($target+'.spdownload.part.json') $Root
                    $record=Join-Path $State ((HashText $row.Remote)+'.json')
                    $known=$null;if(Test-Path -LiteralPath $record){$known=Get-Content -LiteralPath $record -Raw|ConvertFrom-Json}
                    if([SharePointPS.LocalPaths]::Exists($target)){
                        $native=[SharePointPS.LocalPaths]::Native($target)
                        $existing=[IO.FileInfo]::new($native)
                        if(![IO.Directory]::Exists($native) -and $existing.Length -eq $row.Size -and (!$known -or $known.Version -ceq $row.Version)){
                            $skipped++;$doneBytes+=$row.Size;SaveJson $record @{Version=$row.Version;Size=$row.Size;Status='existing'}
                            Log 'ARQUIVO_EXISTENTE' @{remoto=$row.Remote;local=$target;bytes=$row.Size}
                        }else{$failed++;Log 'ARQUIVO_CONFLITO' @{remoto=$row.Remote;local=$target};Write-Warning "Conflito preservado: $target"}
                        continue
                    }
                    [SharePointPS.LocalPaths]::EnsureDirectory([IO.Path]::GetDirectoryName($target))
                    $job=[pscustomobject]@{Row=$row;Record=$record;Attempt=1;At=[DateTime]::UtcNow;Transfer=$null}
                }else{break}
                $transfer=New-Object SharePointPS.Transfer
                $transfer.Start([string[]]@(Routes $job.Row.Remote),$job.Row.Local,[long]$job.Row.Size,($job.Row.Remote+'|'+$job.Row.Version),(CookieJar))
                $job.Transfer=$transfer;[void]$jobs.Add($job)
                Log 'DOWNLOAD_INICIO' @{remoto=$job.Row.Remote;local=$job.Row.Local;tentativa=$job.Attempt}
            }
            $activeBytes=0L
            foreach($job in @($jobs.ToArray())){
                $message=$null
                while($job.Transfer.Events.TryDequeue([ref]$message)){Log 'TRANSFERENCIA_EVENTO' @{remoto=$job.Row.Remote;mensagem=$message}}
                if(!$job.Transfer.Task.IsCompleted){$activeBytes+=$job.Transfer.Bytes;continue}
                $result=$job.Transfer.Task.Result;[void]$jobs.Remove($job)
                if($result.Success){
                    $downloaded++;$doneBytes+=$job.Row.Size
                    SaveJson $job.Record @{Version=$job.Row.Version;Size=$job.Row.Size;Status='done'}
                    Log 'DOWNLOAD_CONCLUIDO' @{remoto=$job.Row.Remote;local=$job.Row.Local;rota=$result.Route;bytes=$job.Row.Size}
                }else{
                    Log 'DOWNLOAD_ERRO' @{remoto=$job.Row.Remote;local=$job.Row.Local;http=$result.Status;mensagem=$result.Message;correlacao=$result.Correlation;tentativa=$job.Attempt;esperado=$result.Expected;recebido=$result.Received;content_length=$result.ResponseLength;encoding=$result.ContentEncoding;etag=$result.ResponseETag}
                    if($result.SizeMismatch -and $job.Attempt -lt 3){
                        try{
                            $fresh=BrowserJson ((RemoteEndpoint File $job.Row.Remote)+'?$select=Length,ETag,TimeLastModified')
                            $version=[string]$fresh.ETag;if(!$version){$version=[string]$fresh.TimeLastModified}
                            $newSize=0L
                            if($null -ne $fresh.Length -and [long]::TryParse([string]$fresh.Length,[ref]$newSize) -and $newSize -ge 0 -and
                                ($newSize -ne $job.Row.Size -or $version -cne $job.Row.Version)){
                                # O payload divergente nunca e aceito. Revalidar a origem e baixar novamente.
                                $manifest=Join-Path ([IO.Path]::GetDirectoryName($Inventory)) 'manifest.json'
                                SaveJson $manifest @{Complete=$false}
                                Log 'ORIGEM_ATUALIZADA' @{remoto=$job.Row.Remote;tamanho_anterior=$job.Row.Size;tamanho_atual=$newSize;versao=$version}
                                $Totals.Bytes+=$newSize-$job.Row.Size
                                $job.Row.Size=$newSize;$job.Row.Version=$version
                            }
                        }catch{Log 'REVALIDACAO_ORIGEM_FALHOU' @{remoto=$job.Row.Remote;tipo=$_.Exception.GetType().Name}}
                    }
                    if($job.Attempt -ge 3){
                        $failed++;SaveJson $job.Record @{Version=$job.Row.Version;Size=$job.Row.Size;Status='failed'}
                        $failure=@{Host=$script:OriginUri.Host;Remote=$job.Row.Remote;Folder=$script:RemoteRoot}
                        SaveJson (Join-Path $PSScriptRoot 'logs\ultima_falha.json') $failure
                        Log 'DOWNLOAD_FALHA_FINAL' @{remoto=$job.Row.Remote;local=$job.Row.Local;mensagem=$result.Message}
                        Write-Warning "Falha: $($job.Row.Remote) - $($result.Message)"
                    }else{
                        $delay=[Math]::Max($job.Attempt*3,$result.RetryAfter)
                        if($result.Status -in @(429,503)){$limit=1;$cooldown=[DateTime]::UtcNow.AddSeconds($delay);Log 'LOTE_LIMITADO' @{http=$result.Status;espera=$delay;workers=1}}
                        elseif($result.Authentication){Renew}
                        $job.Attempt++;$job.At=[DateTime]::UtcNow.AddSeconds($delay);[void]$retry.Add($job)
                    }
                }
            }
            $completed=$downloaded+$skipped
            $percent=if($Totals.Bytes -gt 0){[Math]::Min(100,100*($doneBytes+$activeBytes)/$Totals.Bytes)}else{if($Totals.Count -eq 0){100}else{100*$completed/$Totals.Count}}
            $speed=($doneBytes+$activeBytes)/[Math]::Max(1,$watch.Elapsed.TotalSeconds)
            $eta=if($speed -gt 0){[TimeSpan]::FromSeconds([Math]::Max(0,($Totals.Bytes-$doneBytes-$activeBytes)/$speed)).ToString('d\.hh\:mm\:ss')}else{'--'}
            $text="$completed/$($Totals.Count) | $(BytesText ($doneBytes+$activeBytes))/$(BytesText $Totals.Bytes) | $(BytesText $speed)/s | ETA: $eta | erros: $failed | ativos: $($jobs.Count)"
            Write-Progress -Activity 'Download SharePoint' -Status $text -PercentComplete $percent
            if($watch.Elapsed.TotalSeconds-$lastLog -ge 30){Log 'LOTE_PROGRESSO' @{concluidos=$completed;bytes=$doneBytes;ativos=$jobs.Count;falhas=$failed};$lastLog=$watch.Elapsed.TotalSeconds}
            Start-Sleep -Milliseconds 100
        }
    }finally{
        $reader.Dispose()
        foreach($job in $jobs){$job.Transfer.Cancel=$true}
        foreach($job in $jobs){try{[void]$job.Transfer.Task.Wait(65000)}catch{}}
        Write-Progress -Activity 'Download SharePoint' -Completed
    }
    Log 'LOTE_RESUMO' @{baixados=$downloaded;existentes=$skipped;falhas=$failed;duracao=$watch.Elapsed.TotalSeconds;destino=$Root}
    Write-Host "Finalizado: $downloaded baixados | $skipped existentes | $failed falhas/conflitos. Destino: $Root"
    [pscustomobject]@{Downloaded=$downloaded;Skipped=$skipped;Failed=$failed}
}
function RunBatch($Parsed,[string]$Destination){
    $script:RemoteRoot=$Parsed.Path
    $folder=BrowserJson ((RemoteEndpoint Folder $Parsed.Path)+'?$select=Name,ServerRelativeUrl')
    $base=[IO.Path]::GetFullPath($Destination);$rootName=LocalName $folder.Name
    $root=if([IO.Path]::GetFileName($base.TrimEnd('\')) -ieq $rootName){$base}else{Join-Path $base $rootName}
    $root=SafePath $root $base;[SharePointPS.LocalPaths]::EnsureDirectory($root)
    $lockPath=Join-Path $base '.sharepoint_download.lock'
    try{$lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}catch{throw 'Ja existe um lote usando esse destino ou nao ha permissao de escrita.'}
    try{
        $key=(HashText ($Parsed.Url+'|'+$root)).Substring(0,16);$control=SafePath (Join-Path $base ('.sharepoint_'+$key)) $base
        [void][IO.Directory]::CreateDirectory($control);$state=Join-Path $control 'files';[void][IO.Directory]::CreateDirectory($state)
        $inventory=Join-Path $control 'inventory.jsonl';$manifest=Join-Path $control 'manifest.json'
        $scope=if($Recursive -and !$OnlyCurrentFolder){'recursive'}else{'direct'}
        $totals=$null
        if(!$RefreshInventory -and (Test-Path -LiteralPath $manifest) -and (Test-Path -LiteralPath $inventory)){
            $cached=Get-Content -LiteralPath $manifest -Raw|ConvertFrom-Json
            if($cached.Complete -and $cached.Scope -eq $scope -and $cached.Root -ieq $root -and $cached.Remote -ceq $Parsed.Path){$totals=$cached;Write-Host 'Reutilizando inventario salvo. Use -RefreshInventory para reler a origem.'}
        }
        if(!$totals){
            SaveJson $manifest @{Complete=$false}
            $totals=BuildInventory $Parsed.Path $root $inventory $manifest $scope
        }
        Write-Host "Inventario pronto: $($totals.Count) arquivos | $(BytesText $totals.Bytes) | $($totals.Folders) pastas."
        if(!$totals.Count){Write-Warning 'Nenhum arquivo encontrado no escopo consultado.'}
        if($ListOnly){foreach($line in [IO.File]::ReadLines($inventory)){$r=$line|ConvertFrom-Json;Write-Host "$($r.Size) bytes | $($r.Remote)"};return [pscustomobject]@{Failed=0;Downloaded=0;Skipped=0}}
        RunTransfers $inventory $state $root $totals
    }finally{$lock.Dispose()}
}
