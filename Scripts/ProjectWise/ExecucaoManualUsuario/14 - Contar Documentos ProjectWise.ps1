<#
.SYNOPSIS
    Conta os documentos de uma pasta do ProjectWise e de todas as subpastas.
.DESCRIPTION
    Na janela de login, selecione o datasource e a autenticacao Bentley IMS.
    Informe o caminho interno da pasta (por exemplo: Engenharia\Projetos)
    ou cole o caminho pw:\servidor:datasource\Documents\Engenharia\Projetos.
    A consulta utiliza o datasource selecionado na janela de login.
    Conta os documentos retornados para o usuario conectado, sem baixar arquivos.
    Por padrao considera documentos atuais; -IncluirVersoes inclui o historico.
    Documentos com nomes iguais em pastas distintas sao contados separadamente.
    Requer ProjectWise Explorer, CONNECTION Client e PWPS_DAB instalados.
.EXAMPLE
    powershell.exe -NoProfile -MTA -File ".\14 - Contar Documentos ProjectWise.ps1"
.EXAMPLE
    .\14 - Contar Documentos ProjectWise.ps1 -Caminho 'Engenharia\Projetos' -IncluirVersoes
#>
[CmdletBinding()]
param(
    [string]$DatasourceName = '',
    [string]$Caminho,
    [switch]$IncluirVersoes
)

$ErrorActionPreference = 'Stop'

# A importacao do PWPS_DAB 24 requer Windows PowerShell em MTA.
if ($PSVersionTable.PSEdition -ne 'Desktop' -or
    [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'MTA') {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Salve o script e execute o arquivo completo para abrir a sessao MTA.'
    }

    # Literais PowerShell e comando Unicode preservam espacos, acentos e aspas
    # sem interpretar o caminho ou os parametros como comandos.
    $arquivoLiteral = "'" + $PSCommandPath.Replace("'", "''") + "'"
    $datasourceLiteral = "'" + $DatasourceName.Replace("'", "''") + "'"
    $caminhoLiteral = "'" + ([string]$Caminho).Replace("'", "''") + "'"
    $comando = "& $arquivoLiteral -DatasourceName $datasourceLiteral -Caminho $caminhoLiteral"
    if ($IncluirVersoes) { $comando += ' -IncluirVersoes' }
    $comandoCodificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($comando))
    $executavel = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        $executavel = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    }
    # Janela interativa necessaria para login e Read-Host; NoExit mantem o total visivel.
    $processo = Start-Process -FilePath $executavel -WindowStyle Normal -PassThru `
        -ArgumentList @('-NoProfile', '-MTA', '-NoExit', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $comandoCodificado)
    Write-Host "Windows PowerShell MTA aberto (PID $($processo.Id)). Continue na nova janela." -ForegroundColor Cyan
    return
}

$loginRealizado = $false
try {
    Import-Module PWPS_DAB -DisableNameChecking -ErrorAction Stop

    $parametrosLogin = @{
        UseGui = $true
        DoNotCreateWorkingDirectory = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($DatasourceName)) {
        $parametrosLogin.DatasourceName = $DatasourceName
    }
    Write-Host 'Selecione o datasource e a autenticacao Bentley IMS na janela de login.' -ForegroundColor Cyan
    $loginOk = New-PWLogin @parametrosLogin
    if (-not $loginOk) {
        throw 'Login cancelado ou nao concluido.'
    }
    $loginRealizado = $true

    if ([string]::IsNullOrWhiteSpace($Caminho)) {
        $Caminho = Read-Host 'Caminho interno da pasta ou caminho completo copiado do ProjectWise'
    }
    if ([string]::IsNullOrWhiteSpace($Caminho)) {
        throw 'Informe um caminho de pasta. A contagem nao foi iniciada.'
    }
    $Caminho = $Caminho.Trim().Trim('"').Replace('/', '\').Trim('\')
    if ($Caminho -match '^pw:') {
        if ($Caminho -notmatch '^pw:\\+(?<Origem>[^\\]+)\\Documents\\(?<Pasta>.+)$') {
            throw 'Caminho ProjectWise invalido. Copie o caminho de uma pasta abaixo de Documents.'
        }
        $origemCaminho = $Matches.Origem
        $Caminho = $Matches.Pasta.Trim('\')
        Write-Host "Origem do caminho copiado: $origemCaminho" -ForegroundColor Cyan
        Write-Host "Caminho interno: $Caminho (consulta no datasource selecionado no login)." -ForegroundColor Cyan
    }
    if ([string]::IsNullOrWhiteSpace($Caminho) -or $Caminho -match ':') {
        throw 'Informe uma pasta interna (ex.: Engenharia\Projetos) ou um caminho pw:\servidor:datasource\Documents\Pasta.'
    }

    # Valida a pasta antes de pesquisar, evitando uma busca sem escopo.
    $pastas = @(Get-PWFolders -FolderPath $Caminho -JustOne -ErrorAction Stop |
        Where-Object { $null -ne $_ })
    if ($pastas.Count -ne 1) {
        throw "Pasta nao localizada ou inacessivel: $Caminho"
    }
    $folderId = 0
    if (-not [int]::TryParse([string]$pastas[0].ProjectID, [ref]$folderId) -or $folderId -le 0) {
        throw "A pasta '$Caminho' nao retornou um ProjectID valido."
    }

    $parametrosBusca = @{
        FolderID = $folderId
        JustThisFolder = $false
        GetVersionsToo = [bool]$IncluirVersoes
        ErrorAction = 'Stop'
        WarningAction = 'Stop'
    }
    Write-Host "Contando documentos em '$Caminho' e todas as subpastas..." -ForegroundColor Cyan
    $inicio = Get-Date
    # Measure-Object evita manter outra lista completa de documentos no script.
    $medida = Get-PWDocumentsBySearch @parametrosBusca |
        Where-Object { $null -ne $_ } | Measure-Object

    Write-Host ''
    Write-Host ("Total de documentos: {0:N0}" -f $medida.Count) -ForegroundColor Green
    Write-Host 'Escopo: pasta informada e todas as subpastas acessiveis ao usuario.'
    $criterio = 'Documentos atuais (sem versoes anteriores)'
    if ($IncluirVersoes) { $criterio = 'Documentos atuais e versoes anteriores' }
    Write-Host "Criterio: $criterio"
    [pscustomobject]@{
        Caminho = $Caminho
        FolderID = $folderId
        IncluiSubpastas = $true
        IncluiVersoes = [bool]$IncluirVersoes
        TotalDocumentos = [long]$medida.Count
        DataContagem = Get-Date
        Duracao = (Get-Date) - $inicio
    }
}
finally {
    if ($loginRealizado) {
        try { Undo-PWLogin -ErrorAction Stop | Out-Null }
        catch { Write-Warning "Nao foi possivel encerrar a sessao: $($_.Exception.Message)" }
    }
}
