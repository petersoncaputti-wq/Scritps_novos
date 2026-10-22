#requires -Version 5.1
$ErrorActionPreference='Stop'
$base=Split-Path $PSScriptRoot -Parent
Add-Type -Path (Join-Path $PSScriptRoot 'Fixture.cs')
. (Join-Path $base 'Core.ps1')
# Somente no teste: permitir HTTP/porta do fixture loopback nesta copia em memoria.
# O arquivo de producao continua exigindo HTTPS e porta padrao.
$body=${function:BrowserJson}.ToString().Replace('$u.Scheme -ne ''https''','$u.Scheme -ne ''http''').Replace('!$u.IsDefaultPort','$u.Port -ne $script:OriginUri.Port')
Set-Item Function:BrowserJson ([scriptblock]::Create($body))
$fixture=New-Object SPFixture;$proc=$null;$script:SessionId=$null
try{
    $listener=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,0)
    $listener.Start();$port=$listener.LocalEndpoint.Port;$listener.Stop()
    $script:DriverBase="http://127.0.0.1:$port"
    $proc=Start-Process -FilePath (Join-Path $base 'msedgedriver.exe') -ArgumentList "--port=$port","--allowed-ips=127.0.0.1","--log-level=OFF" -WindowStyle Hidden -PassThru
    for($i=0;$i -lt 40;$i++){try{$null=Driver GET 'status';break}catch{Start-Sleep -Milliseconds 250}}
    $session=Driver POST 'session' @{capabilities=@{alwaysMatch=@{browserName='MicrosoftEdge';pageLoadStrategy='eager';'ms:edgeOptions'=@{args=@('--headless=new','--no-first-run')}}}}
    if($session.capabilities.pageLoadStrategy -ne 'eager'){throw 'Estrategia de navegacao nao aplicada.'}
    $script:SessionId=$session.sessionId
    $null=Driver POST "session/$script:SessionId/url" @{url=$fixture.BaseUrl+'/page'}
    $script:OriginUri=[uri]$fixture.BaseUrl
    $data=BrowserJson ($fixture.BaseUrl+'/json')
    if(@($data.value).Count -ne 1 -or $data.value[0].Name -ne 'fixture.txt'){throw 'JSON recebido pelo Edge divergente.'}
    $items=@(Collection ($fixture.BaseUrl+'/json'))
    if($items.Count -ne 1){throw 'Leitura de colecao divergente.'}
    $jar=CookieJar
    # Adaptacao em memoria para verificar login automatico no fixture HTTP local.
    $waitBody=${function:WaitSharePoint}.ToString().Replace('$site=''https://''+$Parsed.Host','$site=$fixture.BaseUrl')
    Set-Item Function:WaitSharePoint ([scriptblock]::Create($waitBody))
    $script:RunLog=$null
    WaitSharePoint ([pscustomobject]@{Host='127.0.0.1';Path='/'})
    Write-Host 'OK: Edge WebDriver real, sessao headless, fetch JSON no navegador, colecao e cookies em memoria.'
}finally{
    if($script:SessionId){try{$null=Driver DELETE "session/$script:SessionId"}catch{}}
    if($proc -and !$proc.HasExited){Stop-Process -Id $proc.Id}
    $fixture.Dispose()
}
