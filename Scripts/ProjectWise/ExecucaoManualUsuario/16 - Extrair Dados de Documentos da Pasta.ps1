<#
.SYNOPSIS
    Exporta dados dos documentos de uma pasta ProjectWise e todas as subpastas.
.DESCRIPTION
    Selecione concessoes e, para cada uma, a concessao completa ou projetos.
    A concessao completa inclui todas as pastas, mesmo fora da pasta Projetos.
    Opcionalmente, -Caminho permite informar uma pasta diretamente.
    Abre uma janela para escolher o destino; -PastaSaida dispensa essa janela.
    Gera Excel com nome do documento, pasta, versao, estado e extensao.
    Por padrao consulta documentos atuais; -IncluirVersoes inclui o historico.
    Preserva nomes repetidos e documentos logicos, sem baixar arquivos.
    Requer ProjectWise Explorer, PWPS_DAB e ImportExcel instalados.
.EXAMPLE
    powershell.exe -NoProfile -MTA -File ".\16 - Extrair Dados de Documentos da Pasta.ps1"
.EXAMPLE
    & '.\16 - Extrair Dados de Documentos da Pasta.ps1' -Caminho 'Engenharia\Projetos' -PastaSaida 'C:\Relatorios' -IncluirVersoes
#>
[CmdletBinding()]
param(
    [string]$DatasourceName = '',
    [string]$Caminho,
    [string]$PastaSaida,
    [switch]$IncluirVersoes
)

$ErrorActionPreference = 'Stop'

# PWPS_DAB 24 exige Windows PowerShell em MTA. Executa de forma sincrona,
# mantendo o console disponivel para login e entrada dos caminhos.
if ($PSVersionTable.PSEdition -ne 'Desktop' -or
    [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'MTA') {
    if (-not $PSCommandPath) { throw 'Execute o script salvo em arquivo.' }
    $comando = "& '" + $PSCommandPath.Replace("'", "''") + "'"
    foreach ($nomeParametro in @('DatasourceName', 'Caminho', 'PastaSaida')) {
        $valor = [string](Get-Variable -Name $nomeParametro -ValueOnly)
        $comando += " -$nomeParametro '" + $valor.Replace("'", "''") + "'"
    }
    if ($IncluirVersoes) { $comando += ' -IncluirVersoes' }
    $codificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($comando))
    $executavel = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        $executavel = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    }
    & $executavel -NoProfile -MTA -EncodedCommand $codificado
    if ($LASTEXITCODE -ne 0) { throw 'Falha na extracao em Windows PowerShell MTA.' }
    return
}

function Selecionar-PastaSaida {
    # O seletor Windows precisa de STA; a sessao do PW permanece em MTA.
    $codigoSeletor = @'
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
$dialogo.Description = 'Selecione a pasta onde o Excel de documentos sera salvo'
$dialogo.ShowNewFolderButton = $true
$dialogo.SelectedPath = [Environment]::GetFolderPath('DesktopDirectory')
try {
    if ($dialogo.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        [Console]::Out.Write($dialogo.SelectedPath)
    }
}
finally { $dialogo.Dispose() }
'@
    $codificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($codigoSeletor))
    $executavel = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    Write-Host 'Abrindo a janela para selecionar o destino do Excel...' -ForegroundColor Cyan
    $resultado = & $executavel -NoProfile -STA -EncodedCommand $codificado
    if ($LASTEXITCODE -ne 0) { throw 'Nao foi possivel abrir a janela de selecao do destino.' }
    $pasta = ($resultado | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($pasta)) { return $null }
    if (-not (Test-Path -LiteralPath $pasta -PathType Container)) {
        throw "A pasta selecionada nao esta acessivel: $pasta"
    }
    return (Resolve-Path -LiteralPath $pasta).Path
}

function Obter-Texto {
    param([object]$Objeto, [string[]]$Nomes)
    foreach ($nome in $Nomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $texto = $propriedade.Value.ToString().Trim()
            if ($texto.Length -gt 0) { return $texto }
        }
    }
    return ''
}

function Obter-IdPasta {
    param([object]$Pasta)
    # Usa o primeiro identificador numerico valido, mesmo se outro campo
    # existir com valor vazio ou zero. Normaliza a chave usada no mapa.
    foreach ($nome in @('ProjectID', 'FolderID', 'ID')) {
        $valor = Obter-Texto -Objeto $Pasta -Nomes @($nome)
        $id = 0
        if ([int]::TryParse($valor, [ref]$id) -and $id -gt 0) {
            return $id.ToString()
        }
    }
    return ''
}

function Obter-NomePasta {
    param([object]$Pasta)
    return Obter-Texto -Objeto $Pasta -Nomes @('Name', 'FolderName', 'ProjectName', 'ObjectName')
}

function Obter-RotuloPasta {
    param([object]$Pasta)

    $descricao = Obter-Texto -Objeto $Pasta -Nomes @('Description', 'Descricao', 'ProjectDescription', 'FolderDescription')
    if (-not [string]::IsNullOrWhiteSpace($descricao)) { return $descricao }
    return Obter-NomePasta -Pasta $Pasta
}

function Obter-FilhosRaiz {
    $cmd = Get-Command Get-PWFoldersImmediateChildren -ErrorAction Stop
    if ($cmd.Parameters.ContainsKey('Root')) {
        return @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop | Where-Object { $null -ne $_ })
    }
    return @(Get-PWFoldersImmediateChildren -ErrorAction Stop | Where-Object { $null -ne $_ })
}

function Obter-FilhosPasta {
    param([Parameter(Mandatory)][string]$FolderId)
    # Uma pasta final da arvore gera um aviso normal no PWPS_DAB.
    # Captura os avisos para aceitar somente esse caso e manter outros bloqueios.
    $avisos = @()
    $filhos = @(
        Get-PWFoldersImmediateChildren `
            -FolderID $FolderId `
            -WarningAction SilentlyContinue `
            -WarningVariable avisos `
            -ErrorAction Stop |
            Where-Object { $null -ne $_ }
    )
    foreach ($aviso in $avisos) {
        if ([string]$aviso -notmatch '^\s*No immediate child folders found\.?\s*$') {
            throw "Aviso ao listar subpastas da pasta ID ${FolderId}: $aviso"
        }
    }
    foreach ($filho in $filhos) {
        # Algumas consultas PWPS_DAB retornam $false junto ao aviso de
        # ausencia de subpastas. Esse indicador nao representa uma pasta.
        if ($filho -is [bool] -and -not $filho -and
            @($avisos | Where-Object { [string]$_ -match '^\s*No immediate child folders found\.?\s*$' }).Count -gt 0) {
            continue
        }
        if (-not (Obter-IdPasta $filho)) {
            $campos = @($filho.PSObject.Properties | ForEach-Object {
                if ($_.Name -in @('ProjectID', 'FolderID', 'ID', 'Name', 'FolderName', 'ProjectName')) {
                    '{0}={1}' -f $_.Name, $_.Value
                }
            }) -join '; '
            $nomes = $filho.PSObject.Properties.Name -join ', '
            throw "Retorno sem ID de pasta valido abaixo do ID ${FolderId}. Tipo: $($filho.GetType().FullName). Campos: $campos. Propriedades: $nomes"
        }
        $filho
    }
}

function Localizar-Pasta {
    param(
        [array]$Pastas,
        [string[]]$Nomes,
        [string]$Descricao
    )

    foreach ($nomeEsperado in $Nomes) {
        $resultado = @($Pastas | Where-Object {
            (Obter-NomePasta $_) -ieq $nomeEsperado -or (Obter-RotuloPasta $_) -ieq $nomeEsperado
        })
        if ($resultado.Count -gt 0) { return $resultado[0] }
    }

    $disponiveis = @($Pastas | ForEach-Object { Obter-RotuloPasta $_ }) -join ', '
    throw "$Descricao nao encontrada. Pastas disponiveis: $disponiveis"
}

function Mostrar-Itens {
    param([array]$Itens, [string]$Titulo)

    Write-Host ''
    Write-Host $Titulo -ForegroundColor Cyan
    Write-Host ('-' * 70)
    for ($i = 0; $i -lt $Itens.Count; $i++) {
        Write-Host ("{0:000}. {1}" -f ($i + 1), (Obter-RotuloPasta $Itens[$i]))
    }
}

function Selecionar-UmItem {
    param([array]$Itens, [string]$Titulo)

    Mostrar-Itens -Itens $Itens -Titulo $Titulo
    while ($true) {
        $entrada = (Read-Host 'Digite o numero desejado').Trim()
        $numero = 0
        if ([int]::TryParse($entrada, [ref]$numero) -and $numero -ge 1 -and $numero -le $Itens.Count) {
            return $Itens[$numero - 1]
        }
        Write-Host "Informe um numero entre 1 e $($Itens.Count)." -ForegroundColor Yellow
    }
}

function Converter-SelecaoMultipla {
    param([string]$Entrada, [int]$Quantidade)

    if ($Entrada -match '^\s*[Tt]\s*$') { return @(0..($Quantidade - 1)) }
    if ([string]::IsNullOrWhiteSpace($Entrada)) { throw 'Selecao vazia.' }

    $indices = New-Object System.Collections.Generic.HashSet[int]
    foreach ($parteOriginal in ($Entrada -split ',')) {
        $parte = $parteOriginal.Trim()
        if ($parte -match '^(\d+)\s*-\s*(\d+)$') {
            $inicio = [int]$Matches[1]
            $fim = [int]$Matches[2]
            if ($inicio -gt $fim) { throw "Intervalo invalido: $parte" }
            foreach ($numero in $inicio..$fim) {
                if ($numero -lt 1 -or $numero -gt $Quantidade) { throw "Numero fora da lista: $numero" }
                $null = $indices.Add($numero - 1)
            }
        }
        elseif ($parte -match '^\d+$') {
            $numero = [int]$parte
            if ($numero -lt 1 -or $numero -gt $Quantidade) { throw "Numero fora da lista: $numero" }
            $null = $indices.Add($numero - 1)
        }
        else {
            throw "Item invalido: '$parte'"
        }
    }
    return @($indices | Sort-Object)
}

function Selecionar-VariosItens {
    param([array]$Itens, [string]$Titulo)

    Mostrar-Itens -Itens $Itens -Titulo $Titulo
    Write-Host ''
    Write-Host 'Use virgulas, intervalos ou T para todos. Exemplos: 1,3,5 | 1-4,8 | T'

    while ($true) {
        try {
            $indices = @(Converter-SelecaoMultipla -Entrada (Read-Host 'Selecao') -Quantidade $Itens.Count)
            if ($indices.Count -eq 0) { throw 'Nenhum projeto selecionado.' }
            return @($indices | ForEach-Object { $Itens[$_] })
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Yellow
        }
    }
}

$loginRealizado = $false
try {
    Import-Module PWPS_DAB -DisableNameChecking -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
    $parametrosLogin = @{
        UseGui = $true
        DoNotCreateWorkingDirectory = $true
        ErrorAction = 'Stop'
    }
    if ($DatasourceName) { $parametrosLogin.DatasourceName = $DatasourceName }
    Write-Host 'Selecione o datasource e a autenticacao na janela de login.' -ForegroundColor Cyan
    if (-not (New-PWLogin @parametrosLogin)) { throw 'Login cancelado ou nao concluido.' }
    $loginRealizado = $true

    $escopos = New-Object 'System.Collections.Generic.List[object]'
    if ([string]::IsNullOrWhiteSpace($Caminho)) {
        $engenharia = Localizar-Pasta -Pastas (Obter-FilhosRaiz) -Nomes @('Engenharia', 'Engineering') -Descricao 'Pasta Engenharia'
        $concessoes = @(Obter-FilhosPasta (Obter-IdPasta $engenharia) | Sort-Object { Obter-RotuloPasta $_ })
        if ($concessoes.Count -eq 0) { throw 'Nenhuma concessao encontrada.' }
        $selecionadas = @(Selecionar-VariosItens -Itens $concessoes -Titulo 'Selecione as concessoes')
        foreach ($concessao in $selecionadas) {
            $caminhoConcessao = (Obter-NomePasta $engenharia) + '\' + (Obter-NomePasta $concessao)
            Write-Host ''
            Write-Host (Obter-RotuloPasta $concessao) -ForegroundColor Cyan
            Write-Host '1 - Extrair a concessao completa (todas as pastas e subpastas)'
            Write-Host '2 - Selecionar projetos'
            Write-Host '3 - Extrair todos os projetos (somente dentro de Projetos)'
            Write-Host '0 - Ignorar esta concessao'
            do { $opcao = (Read-Host 'Escolha uma opcao (1/2/3/0)').Trim() } while ($opcao -notin @('1', '2', '3', '0'))
            if ($opcao -eq '0') { continue }
            if ($opcao -eq '1') {
                $escopos.Add([pscustomobject]@{ Pasta = $concessao; Caminho = $caminhoConcessao })
                continue
            }
            $pastaProjetos = Localizar-Pasta -Pastas (Obter-FilhosPasta (Obter-IdPasta $concessao)) -Nomes @('Projetos', 'Projeto', 'Projects', 'Project') -Descricao "Pasta Projetos de $(Obter-RotuloPasta $concessao)"
            $caminhoProjetos = $caminhoConcessao + '\' + (Obter-NomePasta $pastaProjetos)
            if ($opcao -eq '3') {
                # Inclui documentos diretamente na pasta Projetos, alem dos projetos.
                $escopos.Add([pscustomobject]@{ Pasta = $pastaProjetos; Caminho = $caminhoProjetos })
                continue
            }
            $projetos = @(Obter-FilhosPasta (Obter-IdPasta $pastaProjetos) | Sort-Object { Obter-RotuloPasta $_ })
            if ($projetos.Count -eq 0) { throw "Nenhum projeto encontrado em $caminhoProjetos." }
            foreach ($projeto in @(Selecionar-VariosItens -Itens $projetos -Titulo 'Selecione os projetos')) {
                $escopos.Add([pscustomobject]@{ Pasta = $projeto; Caminho = $caminhoProjetos + '\' + (Obter-NomePasta $projeto) })
            }
        }
    }
    else {
    $Caminho = $Caminho.Trim().Trim('"').Replace('/', '\').Trim('\')
    if ($Caminho -match '^pw:') {
        if ($Caminho -notmatch '^pw:\\+(?<Origem>[^\\]+)\\Documents\\(?<Pasta>.+)$') {
            throw 'Copie o caminho de uma pasta abaixo de Documents.'
        }
        Write-Host "Origem do caminho: $($Matches.Origem). A consulta usa o datasource selecionado no login."
        $Caminho = $Matches.Pasta.Trim('\')
    }
    if ([string]::IsNullOrWhiteSpace($Caminho) -or $Caminho -match ':') {
        throw 'Informe uma pasta interna valida do ProjectWise.'
    }
    $pastas = @(Get-PWFolders -FolderPath $Caminho -JustOne -ErrorAction Stop -WarningAction Stop |
        Where-Object { $null -ne $_ })
    if ($pastas.Count -ne 1) { throw "Pasta nao localizada ou inacessivel: $Caminho" }
    $folderId = 0
    if (-not [int]::TryParse((Obter-Texto $pastas[0] @('ProjectID', 'FolderID')), [ref]$folderId) -or $folderId -le 0) {
        throw 'A pasta nao retornou um ID valido.'
    }
        $escopos.Add([pscustomobject]@{ Pasta = $pastas[0]; Caminho = $Caminho })
    }
    if ($escopos.Count -eq 0) {
        Write-Host 'Nenhuma concessao ou projeto selecionado. Extracao encerrada.' -ForegroundColor Yellow
        return
    }
    Write-Host 'Pastas selecionadas para extracao:' -ForegroundColor Cyan
    foreach ($escopo in $escopos) { Write-Host (' - ' + $escopo.Caminho) }
    if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
        $PastaSaida = Selecionar-PastaSaida
        if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
            Write-Host 'Selecao do destino cancelada. Extracao encerrada.' -ForegroundColor Yellow
            return
        }
    }
    if (-not $PastaSaida -or -not (Test-Path -LiteralPath $PastaSaida -PathType Container)) {
        throw 'Informe uma pasta de saida existente e acessivel.'
    }

    $todasLinhas = New-Object 'System.Collections.Generic.List[object]'
    $indiceEscopo = 0
    foreach ($escopo in $escopos) {
    $indiceEscopo++
    Write-Progress -Id 1 -Activity 'Extracao de documentos ProjectWise' `
        -Status "Escopo $indiceEscopo de $($escopos.Count): $($escopo.Caminho)" `
        -PercentComplete ([int](100 * ($indiceEscopo - 1) / $escopos.Count))
    $Caminho = $escopo.Caminho
    $folderId = 0
    if (-not [int]::TryParse((Obter-IdPasta $escopo.Pasta), [ref]$folderId) -or $folderId -le 0) {
        throw "ID invalido para a pasta $Caminho."
    }
    Write-Host "Mapeando '$Caminho' e todas as subpastas..." -ForegroundColor Cyan
    $mapa = @{}
    $mapa[[string]$folderId] = $Caminho
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([string]$folderId)
    $pastasMapeadas = 0
    while ($fila.Count -gt 0) {
        $idAtual = [string]$fila.Dequeue()
        # O total da arvore ainda e desconhecido nesta etapa.
        Write-Progress -Id 2 -ParentId 1 -Activity 'Mapeando pastas' `
            -Status "$pastasMapeadas pastas mapeadas; $($fila.Count + 1) pendentes" `
            -CurrentOperation $mapa[$idAtual] -PercentComplete -1
        foreach ($filho in @(Obter-FilhosPasta -FolderId $idAtual)) {
            if ($null -eq $filho) { continue }
            $idFilho = Obter-IdPasta $filho
            $numeroId = 0
            if (-not [int]::TryParse($idFilho, [ref]$numeroId) -or $numeroId -le 0) {
                throw 'Uma subpasta retornou ID invalido. Extracao interrompida.'
            }
            if ($mapa.ContainsKey($idFilho)) { continue }
            $nomeFilho = Obter-Texto $filho @('Name', 'FolderName', 'ProjectName')
            if (-not $nomeFilho) { throw "Nome ausente na pasta ID $idFilho." }
            $mapa[$idFilho] = $mapa[$idAtual].TrimEnd('\') + '\' + $nomeFilho
            $fila.Enqueue($idFilho)
        }
        $pastasMapeadas++
    }

    $parametrosBusca = @{
        FolderID = $folderId
        JustThisFolder = $false
        GetVersionsToo = [bool]$IncluirVersoes
        ErrorAction = 'Stop'
        WarningAction = 'Stop'
    }
    Write-Host "Extraindo documentos de '$Caminho' e todas as subpastas..." -ForegroundColor Cyan
    $documentosProcessados = 0
    $relogioProgresso = [Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Id 2 -ParentId 1 -Activity 'Consultando documentos no ProjectWise' `
        -Status "$pastasMapeadas pastas mapeadas. Aguardando documentos..." -PercentComplete -1
    $linhas = @(Get-PWDocumentsBySearch @parametrosBusca | ForEach-Object {
        if ($null -ne $_) {
            $documentosProcessados++
            # Evita atualizar a tela a cada documento em consultas grandes.
            if ($documentosProcessados -eq 1 -or $relogioProgresso.ElapsedMilliseconds -ge 250) {
                Write-Progress -Id 2 -ParentId 1 -Activity 'Processando documentos' `
                    -Status "$documentosProcessados documentos processados neste escopo; $($todasLinhas.Count) nos anteriores" `
                    -CurrentOperation $Caminho -PercentComplete -1
                $relogioProgresso.Restart()
            }
            # Em documentos, ID identifica o documento, nao sua pasta.
            $idPasta = Obter-IdPasta ([pscustomobject]@{
                ProjectID = Obter-Texto $_ @('ProjectID')
                FolderID = Obter-Texto $_ @('FolderID')
            })
            if (-not $mapa.ContainsKey($idPasta)) { throw "Pasta do documento ausente no escopo: ID $idPasta." }
            $nomeFisico = Obter-Texto $_ @('FileName', 'Filename')
            $nomeDocumento = Obter-Texto $_ @('Name', 'DocumentName')
            if (-not $nomeDocumento) { $nomeDocumento = $nomeFisico }
            [pscustomobject][ordered]@{
                'Nome do documento' = $nomeDocumento
                'Pasta' = $mapa[$idPasta]
                'Versao' = Obter-Texto $_ @('Version', 'DocumentVersion', 'VersionName')
                'Estado' = Obter-Texto $_ @('WorkflowState', 'StateName', 'State')
                'Extensao' = [IO.Path]::GetExtension($nomeFisico).ToLowerInvariant()
            }
        }
    })
    $relogioProgresso.Stop()
    foreach ($linha in $linhas) { $todasLinhas.Add($linha) }
    Write-Progress -Id 2 -ParentId 1 -Activity 'Processando documentos' -Completed
    Write-Progress -Id 1 -Activity 'Extracao de documentos ProjectWise' `
        -Status "$indiceEscopo de $($escopos.Count) escopos consultados; $($todasLinhas.Count) documentos" `
        -PercentComplete ([int](100 * $indiceEscopo / $escopos.Count))
    }
    if ($todasLinhas.Count -eq 0) {
        Write-Host 'Nenhum documento encontrado. Nenhum Excel foi gerado.' -ForegroundColor Yellow
        return
    }
    if ($todasLinhas.Count -gt 1048575) { throw 'Quantidade excede o limite de linhas de uma aba Excel.' }
    $arquivoSaida = Join-Path $PastaSaida ('Extracao_Documentos_{0}_{1}.xlsx' -f
        (Get-Date -Format 'yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    Write-Progress -Id 2 -ParentId 1 -Activity 'Gerando arquivo Excel' `
        -Status "Salvando $($todasLinhas.Count) documentos..." -CurrentOperation $arquivoSaida -PercentComplete -1
    $todasLinhas | Export-Excel -Path $arquivoSaida -WorksheetName 'Documentos' `
        -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
        -NoNumberConversion @('Nome do documento', 'Pasta', 'Versao', 'Estado', 'Extensao')
    Write-Host "Extracao concluida. Documentos exportados: $($todasLinhas.Count)" -ForegroundColor Green
    Write-Host "Arquivo: $arquivoSaida"
}
finally {
    Write-Progress -Id 2 -Activity 'Extracao de documentos' -Completed
    Write-Progress -Id 1 -Activity 'Extracao de documentos ProjectWise' -Completed
    if ($loginRealizado) {
        try { Undo-PWLogin -ErrorAction Stop | Out-Null }
        catch { Write-Warning "Nao foi possivel encerrar a sessao: $($_.Exception.Message)" }
    }
}
