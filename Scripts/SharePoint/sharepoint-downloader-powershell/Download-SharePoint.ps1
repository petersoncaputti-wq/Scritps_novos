#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Url,
    [string]$Destination,
    [string]$DriverPath,
    [switch]$Recursive=$true,
    [switch]$OnlyCurrentFolder,
    [switch]$ListOnly,
    [ValidateRange(1,8)][int]$Workers=3,
    [switch]$RefreshInventory,
    [string]$ProjectList,
    [string]$Sheet,
    [int[]]$Priority,
    [int[]]$ProjectRows,
    [switch]$PlanOnly,
    [switch]$BrowserTest,
    [switch]$DiagnoseDownload,
    [string]$DiagnoseFile
)
$ErrorActionPreference='Stop'
$script:ExitCode=0;$script:SessionId=$null;$script:AuthenticatedHost=$null;$driverProcess=$null;$script:RunLog=$null
try{
    [AppContext]::SetSwitch('Switch.System.IO.UseLegacyPathHandling',$false)
    [AppContext]::SetSwitch('Switch.System.IO.BlockLongPaths',$false)
    Add-Type -AssemblyName System.Web
    if(!('SharePointPS.Transfer' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'Transfer.cs') -ReferencedAssemblies System.Runtime.Serialization,System.Xml}
    . (Join-Path $PSScriptRoot 'Core.ps1')
    . (Join-Path $PSScriptRoot 'Projects.ps1')
    [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
    [Net.ServicePointManager]::DefaultConnectionLimit=16
    $logDir=Join-Path $PSScriptRoot 'logs';[void][IO.Directory]::CreateDirectory($logDir)
    $script:RunLog=Join-Path $logDir ('execucao_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)+'.log')
    Write-Host "Log desta execucao: $script:RunLog";Log 'APLICACAO_INICIO' @{workers=$Workers}
    if($PlanOnly -and !$ProjectList){throw '-PlanOnly requer -ProjectList.'}
    if($ProjectList -and ($BrowserTest -or $DiagnoseDownload -or $DiagnoseFile)){throw 'Fila de projetos nao pode ser combinada com teste ou diagnostico.'}
    if($BrowserTest -and ($DiagnoseDownload -or $DiagnoseFile)){throw 'Escolha teste ou diagnostico.'}
    if($ProjectList){
        $queue=LoadProjects $ProjectList
        foreach($issue in $queue.Issues){Write-Warning "Linha $($issue.Linha): $($issue.Motivo)"}
        foreach($project in $queue.Projects){Write-Host "Prioridade $($project.Priority) | linha $($project.Row) | $($project.Folder)"}
        Write-Host "Fila: $($queue.Projects.Count) projetos | $($queue.Issues.Count) pendencias."
        if($PlanOnly){if($queue.Issues.Count){$script:ExitCode=1};exit $script:ExitCode}
        if(!$queue.Projects.Count){throw 'Nenhum projeto valido selecionado.'}
    }else{
        if($DiagnoseDownload -and !$DiagnoseFile){
            $last=Join-Path $logDir 'ultima_falha.json'
            if(!(Test-Path -LiteralPath $last)){throw 'Nenhuma falha salva. Informe -Url e -DiagnoseFile.'}
            $failure=Get-Content -LiteralPath $last -Raw|ConvertFrom-Json;$DiagnoseFile=$failure.Remote
            if(!$Url){$Url='https://'+$failure.Host+(($failure.Folder.Split('/')|ForEach-Object{[uri]::EscapeDataString($_)}) -join '/')}
        }
        if(!$Url){$Url=Read-Host 'URL da pasta SharePoint'}
        $parsed=ParseUrl $Url
    }
    if(!$Destination -and !($DiagnoseDownload -or $DiagnoseFile)){$Destination=SelectDestination}
    if($Destination){$Destination=[IO.Path]::GetFullPath($Destination)}
    if(!$DriverPath){$DriverPath=Join-Path $PSScriptRoot 'msedgedriver.exe'}
    if(!(Test-Path -LiteralPath $DriverPath -PathType Leaf)){throw 'Coloque msedgedriver.exe nesta pasta ou informe -DriverPath. Veja README.md.'}
    $listener=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,0)
    $listener.Start();$port=$listener.LocalEndpoint.Port;$listener.Stop()
    $script:DriverBase="http://127.0.0.1:$port"
    $driverProcess=Start-Process -FilePath $DriverPath -ArgumentList "--port=$port","--allowed-ips=127.0.0.1","--log-level=OFF" -WindowStyle Hidden -PassThru
    $ready=$false
    for($i=0;$i -lt 40;$i++){try{$null=Driver GET 'status';$ready=$true;break}catch{Start-Sleep -Milliseconds 250}}
    if(!$ready){throw 'Edge WebDriver nao iniciou. Confira bloqueios corporativos e versao do driver.'}
    $initialUrl=if($ProjectList){$queue.Projects[0].Parsed.Url}else{$parsed.Url}
    if(!$initialUrl){throw 'A primeira pasta nao tem URL valida; navegador nao sera iniciado.'}
    Write-Host 'Abrindo Edge para autenticar no SharePoint...'
    Log 'EDGE_SESSAO_INICIO' @{etapa='criacao';versao_fluxo=2}
    $session=Driver POST 'session' @{capabilities=@{alwaysMatch=@{browserName='MicrosoftEdge';pageLoadStrategy='eager';'ms:edgeOptions'=@{args=@('--no-first-run')}}}}
    $script:SessionId=$session.sessionId
    if(!$script:SessionId){throw 'Edge abriu sem estabelecer sessao de controle.'}
    Log 'EDGE_SESSAO_PRONTA' @{etapa='configuracao'}
    Write-Host 'Edge conectado. Configurando navegacao...'
    $null=Driver POST "session/$script:SessionId/timeouts" @{script=120000;pageLoad=120000}
    Log 'EDGE_CONFIGURADO'
    if($ProjectList){
        # Navegar antes de gravar o resumo da fila evita deixar about:blank
        # enquanto um destino lento/rede esta preparando o controle.
        OpenFolder $queue.Projects[0].Parsed
        $script:ExitCode=RunQueue $queue $Destination -FirstFolderOpened
    }
    else{
        OpenFolder $parsed
        $script:RemoteRoot=$parsed.Path
        if($DiagnoseDownload -or $DiagnoseFile){Diagnose $DiagnoseFile}
        elseif($BrowserTest){BrowserTest $parsed $Destination}
        else{$result=RunBatch $parsed $Destination;if($result.Failed){$script:ExitCode=1}}
    }
}catch [System.Management.Automation.PipelineStoppedException]{
    $script:ExitCode=130
    if(Get-Command Log -ErrorAction SilentlyContinue){Log 'APLICACAO_CANCELADA'}
    throw
}catch{
    $script:ExitCode=1
    $reason=if('SharePointPS.Transfer' -as [type]){[SharePointPS.Transfer]::Clean($_.Exception.Message)}else{'Falha ao carregar componentes .NET: '+$_.Exception.GetType().Name}
    Write-Host "ERRO: $reason" -ForegroundColor Red
    Write-Host "Local: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)"
    if(Get-Command Log -ErrorAction SilentlyContinue){Log 'APLICACAO_ERRO' @{motivo=$reason;linha=$_.InvocationInfo.ScriptLineNumber}}
}finally{
    if($script:SessionId){try{$null=Driver DELETE "session/$script:SessionId"}catch{}}
    if($driverProcess -and !$driverProcess.HasExited){Stop-Process -Id $driverProcess.Id -ErrorAction SilentlyContinue}
    if(Get-Command Log -ErrorAction SilentlyContinue){Log 'APLICACAO_FIM' @{codigo=$script:ExitCode}}
    $global:LASTEXITCODE=$script:ExitCode
}
exit $script:ExitCode
