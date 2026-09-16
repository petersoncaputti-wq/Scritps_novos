<#
.SYNOPSIS
    Exporta o conteudo das pastas Area de Transferencia dos projetos selecionados.
.DESCRIPTION
    Autentica no ProjectWise, solicita o destino local, permite selecionar uma
    concessao e todos ou alguns projetos. Em cada projeto, procura recursivamente
    a pasta Area de Transferencia e exporta todo o seu conteudo.
#>

[CmdletBinding()]
param(
    [string]$DatasourceName = '',
    [string]$Destino,
    [switch]$SomenteListar,
    [switch]$ModoConsole
)

if (-not $ModoConsole) {
    if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        throw 'Execute o arquivo .ps1 completo e salvo, em vez de colar ou executar apenas uma selecao do codigo.'
    }
    $arquivoInterface = Join-Path $PSScriptRoot 'Interface_Exportacao_ProjectWise.ps1'
    if (-not (Test-Path -LiteralPath $arquivoInterface -PathType Leaf)) {
        throw "Arquivo da interface nao encontrado: '$arquivoInterface'. Copie Interface_Exportacao_ProjectWise.ps1 para a mesma pasta deste script. Para executar sem interface, use o parametro -ModoConsole."
    }
    & $arquivoInterface -DatasourceName $DatasourceName -Destino $Destino -SomenteListar:$SomenteListar
    return
}

# O PWPS_DAB 24.0.2 exige uma runspace MTA durante a importacao. O terminal
# especial da extensao PowerShell do VS Code normalmente usa STA. Nesse caso,
# reinicia este arquivo em uma janela Windows PowerShell MTA independente.
$estadoApartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
if ($estadoApartment -ne [System.Threading.ApartmentState]::MTA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'O PWPS_DAB exige MTA. Execute o script salvo usando powershell.exe -MTA -File.'
    }

    Write-Host "Sessao atual em $estadoApartment. Abrindo uma janela PowerShell MTA..." -ForegroundColor Yellow

    function Proteger-ArgumentoProcesso {
        param([string]$Valor)
        return ('"{0}"' -f ($Valor -replace '"', '\"'))
    }

    $partesArgumento = @(
        '-NoProfile',
        '-MTA',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Proteger-ArgumentoProcesso $PSCommandPath),
        '-ModoConsole'
    )
    if (-not [string]::IsNullOrWhiteSpace($DatasourceName)) {
        $partesArgumento += @('-DatasourceName', (Proteger-ArgumentoProcesso $DatasourceName))
    }
    if (-not [string]::IsNullOrWhiteSpace($Destino)) {
        $partesArgumento += @('-Destino', (Proteger-ArgumentoProcesso $Destino))
    }
    if ($SomenteListar) {
        $partesArgumento += '-SomenteListar'
    }

    $executavelPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $processo = Start-Process `
        -FilePath $executavelPowerShell `
        -ArgumentList ($partesArgumento -join ' ') `
        -PassThru

    Write-Host "Processo MTA iniciado (PID $($processo.Id)). Continue na nova janela." -ForegroundColor Green
    return
}

$ErrorActionPreference = 'Stop'
$NomesEngenharia = @('Engenharia', 'Engineering')
$NomesProjetos = @('Projetos', 'Projeto', 'Projects', 'Project')
$script:LoginRealizado = $false

function Write-Etapa {
    param([string]$Mensagem)
    Write-Host ("{0} | {1}" -f (Get-Date -Format 'HH:mm:ss'), $Mensagem) -ForegroundColor Cyan
}

function Remover-Acentos {
    param([AllowNull()][string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return '' }
    $normalizado = $Texto.Normalize([Text.NormalizationForm]::FormD)
    $resultado = foreach ($caractere in $normalizado.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($caractere) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            $caractere
        }
    }
    return (-join $resultado).Normalize([Text.NormalizationForm]::FormC)
}

function Obter-Texto {
    param([object]$Objeto, [string[]]$Nomes)
    if ($null -eq $Objeto) { return '' }
    foreach ($nome in $Nomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $texto = $propriedade.Value.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($texto)) { return $texto }
        }
    }
    return ''
}

function Obter-IdPasta {
    param([object]$Pasta)
    return Obter-Texto $Pasta @('ProjectID', 'ProjectId', 'FolderID', 'FolderId', 'ID', 'Id')
}

function Obter-NomePasta {
    param([object]$Pasta)
    return Obter-Texto $Pasta @('Name', 'FolderName', 'ProjectName', 'ObjectName')
}

function Obter-RotuloPasta {
    param([object]$Pasta)
    $descricao = Obter-Texto $Pasta @('Description', 'Descricao', 'ProjectDescription', 'FolderDescription')
    if (-not [string]::IsNullOrWhiteSpace($descricao)) { return $descricao }
    return Obter-NomePasta $Pasta
}

function Obter-CaminhoPW {
    param([object]$Pasta)
    $caminho = Obter-Texto $Pasta @('FullPath', 'FolderPath', 'Path')
    if ($caminho) { return $caminho }
    try {
        $null = $Pasta.GetFullPath()
        return Obter-Texto $Pasta @('FullPath', 'FolderPath', 'Path')
    }
    catch { return '' }
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
    return @(Get-PWFoldersImmediateChildren -FolderID $FolderId -WarningAction SilentlyContinue -ErrorAction Stop |
        Where-Object { $null -ne $_ })
}

function Localizar-Pasta {
    param([array]$Pastas, [string[]]$Nomes, [string]$Descricao)
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
    Mostrar-Itens $Itens $Titulo
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
    $indices = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($original in ($Entrada -split ',')) {
        $parte = $original.Trim()
        if ($parte -match '^(\d+)\s*-\s*(\d+)$') {
            $inicio = [int]$Matches[1]; $fim = [int]$Matches[2]
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
        else { throw "Item invalido: '$parte'" }
    }
    return @($indices | Sort-Object)
}

function Selecionar-Projetos {
    param([array]$Projetos, [string]$NomeConcessao)
    Write-Host ''
    Write-Host ("{0} - {1} projetos encontrados" -f $NomeConcessao, $Projetos.Count) -ForegroundColor Cyan
    Write-Host '1 - Analisar todos os projetos'
    Write-Host '2 - Escolher projetos'
    do { $opcao = (Read-Host 'Escolha uma opcao (1/2)').Trim() } while ($opcao -notin @('1', '2'))
    if ($opcao -eq '1') { return @($Projetos) }

    Mostrar-Itens $Projetos "Projetos de $NomeConcessao"
    Write-Host 'Use virgulas, intervalos ou T para todos. Exemplos: 1,3,5 | 1-4,8 | T'
    while ($true) {
        try {
            $indices = @(Converter-SelecaoMultipla (Read-Host 'Selecao') $Projetos.Count)
            if ($indices.Count -eq 0) { throw 'Nenhum projeto selecionado.' }
            return @($indices | ForEach-Object { $Projetos[$_] })
        }
        catch { Write-Host $_.Exception.Message -ForegroundColor Yellow }
    }
}

function Selecionar-PastaDestino {
    $padrao = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory).Replace("'", "''")
    $codigo = @"
Add-Type -AssemblyName System.Windows.Forms
`$d = New-Object System.Windows.Forms.FolderBrowserDialog
`$d.Description = 'Selecione a pasta onde os arquivos serao exportados'
`$d.ShowNewFolderButton = `$true
`$d.SelectedPath = '$padrao'
if (`$d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Out.Write(`$d.SelectedPath) }
`$d.Dispose()
"@
    $codificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($codigo))
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    Write-Etapa 'Abrindo a janela para selecionar a pasta de destino...'
    $caminho = ((& $powershell -NoProfile -STA -EncodedCommand $codificado) | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($caminho)) { return $null }
    return [IO.Path]::GetFullPath($caminho)
}

function Testar-NomeAreaTransferencia {
    param([string]$Nome)
    $normalizado = (Remover-Acentos $Nome).Trim().ToLowerInvariant()
    # Aceita a grafia correta "transferencia", a variante "tranferencia" e
    # espacos antes/depois do nome (removidos pelo Trim acima).
    return $normalizado -match '^(\d+\s*-\s*)?area de trans?ferencia$'
}

function Localizar-AreasTransferenciaNoProjeto {
    param([object]$Projeto)
    $resultado = New-Object 'System.Collections.Generic.List[object]'
    $fila = New-Object System.Collections.Queue
    $idProjeto = Obter-IdPasta $Projeto
    if (-not $idProjeto) { return @() }
    $fila.Enqueue($idProjeto)
    while ($fila.Count -gt 0) {
        $idPai = [string]$fila.Dequeue()
        foreach ($filho in (Obter-FilhosPasta $idPai)) {
            $idFilho = Obter-IdPasta $filho
            if (-not $idFilho) { continue }
            $nomeFilho = Obter-NomePasta $filho
            # Ignora estas ramificacoes somente no primeiro nivel do projeto.
            # Aceita 0, 1, 2, 00, 01, 02 e prefixos como 0 - Nome ou 02_Nome.
            # Nao ignora 10, 20, 010 ou 020.
            if ($idPai -eq $idProjeto -and $nomeFilho.Trim() -match '^0?[012](?:$|[\s._\-\u2013\u2014])') {
                continue
            }
            if (Testar-NomeAreaTransferencia $nomeFilho) {
                $resultado.Add($filho)
                continue
            }
            $fila.Enqueue($idFilho)
        }
    }
    # No Windows PowerShell 5.1, @($listaGenerica) pode gerar
    # "Argument types do not match". ToArray() evita o binder problematico.
    return $resultado.ToArray()
}

function Converter-NomePastaLocal {
    param([string]$Nome)
    $resultado = $Nome
    foreach ($caractere in [IO.Path]::GetInvalidFileNameChars()) {
        $resultado = $resultado.Replace([string]$caractere, '_')
    }
    $resultado = $resultado.Trim().TrimEnd('.')
    if (-not $resultado) { return 'Sem_nome' }
    return $resultado
}

try {
    Write-Etapa 'Validando dependencias...'
    # O inicializador do PWPS_DAB 24.0.2 consulta .Length em uma variavel de
    # ambiente que pode ser nula. Por isso, o modulo deve ser importado antes
    # de ativar o modo estrito deste script.
    Set-StrictMode -Off
    Import-Module PWPS_DAB -ErrorAction Stop
    Set-StrictMode -Version 3.0
    $parametrosLogin = @{
        UseGui = $true
        DoNotCreateWorkingDirectory = $true
        ErrorAction = 'Stop'
    }
    if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
        Write-Etapa 'Abrindo login do ProjectWise. Selecione o datasource e Bentley IMS...'
    }
    else {
        Write-Etapa "Abrindo login do ProjectWise no datasource $DatasourceName..."
        $parametrosLogin.DatasourceName = $DatasourceName
    }
    $loginOk = New-PWLogin @parametrosLogin
    if (-not $loginOk) { throw 'O login do ProjectWise foi cancelado ou nao foi concluido.' }
    $script:LoginRealizado = $true
    Write-Etapa 'Login concluido.'

    if ([string]::IsNullOrWhiteSpace($Destino)) {
        $Destino = Selecionar-PastaDestino
        if (-not $Destino) {
            Write-Host 'Selecao do destino cancelada. Nenhum processamento foi iniciado.' -ForegroundColor Yellow
            return
        }
    }
    else { $Destino = [IO.Path]::GetFullPath($Destino) }
    if (-not (Test-Path -LiteralPath $Destino -PathType Container)) {
        New-Item -ItemType Directory -Path $Destino -Force | Out-Null
    }
    Write-Host "Pasta de destino: $Destino" -ForegroundColor Green

    Write-Etapa 'Localizando concessoes...'
    $engenharia = Localizar-Pasta (Obter-FilhosRaiz) $NomesEngenharia 'Pasta Engenharia'
    $concessoes = @(Obter-FilhosPasta (Obter-IdPasta $engenharia) | Sort-Object { (Obter-RotuloPasta $_).ToLowerInvariant() })
    if ($concessoes.Count -eq 0) { throw 'Nenhuma concessao encontrada.' }
    $concessao = Selecionar-UmItem $concessoes 'Concessoes encontradas'
    $nomeConcessao = Obter-RotuloPasta $concessao

    Write-Etapa "Localizando projetos de '$nomeConcessao'..."
    $pastaProjetos = Localizar-Pasta (Obter-FilhosPasta (Obter-IdPasta $concessao)) $NomesProjetos "Pasta Projetos de $nomeConcessao"
    $todosProjetos = @(Obter-FilhosPasta (Obter-IdPasta $pastaProjetos) | Sort-Object { (Obter-RotuloPasta $_).ToLowerInvariant() })
    if ($todosProjetos.Count -eq 0) { throw "Nenhum projeto encontrado em '$nomeConcessao'." }
    $projetos = @(Selecionar-Projetos $todosProjetos $nomeConcessao)

    Write-Etapa "Procurando Areas de Transferencia em $($projetos.Count) projeto(s)..."
    $areas = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $projetos.Count; $i++) {
        $projeto = $projetos[$i]
        $nomeProjeto = Obter-RotuloPasta $projeto
        Write-Progress -Activity 'Localizando Areas de Transferencia' -Status "$($i + 1)/$($projetos.Count) - $nomeProjeto" -PercentComplete ((($i + 1) / $projetos.Count) * 100)
        try {
            $encontradas = @(Localizar-AreasTransferenciaNoProjeto $projeto)
            foreach ($area in $encontradas) {
                $areas.Add([pscustomobject]@{
                    Concessao = $nomeConcessao
                    Projeto = $nomeProjeto
                    ProjetoID = Obter-IdPasta $projeto
                    AreaID = Obter-IdPasta $area
                    CaminhoPW = Obter-CaminhoPW $area
                })
            }
            if ($encontradas.Count -eq 0) { Write-Host "Area de Transferencia nao encontrada em '$nomeProjeto'." -ForegroundColor Yellow }
        }
        catch { Write-Host "Falha ao pesquisar '$nomeProjeto': $($_.Exception.Message)" -ForegroundColor Yellow }
    }
    Write-Progress -Activity 'Localizando Areas de Transferencia' -Completed

    $dataHora = Get-Date -Format 'yyyyMMdd_HHmmss'
    $relatorio = Join-Path $Destino "Relatorio_Areas_Transferencia_$dataHora.csv"
    $areas.ToArray() | Export-Csv -LiteralPath $relatorio -NoTypeInformation -Encoding UTF8
    Write-Host "Projetos analisados: $($projetos.Count)" -ForegroundColor Green
    Write-Host "Areas encontradas: $($areas.Count)" -ForegroundColor Green
    Write-Host "Relatorio: $relatorio"
    if ($areas.Count -eq 0) { throw 'Nenhuma Area de Transferencia foi encontrada nos projetos selecionados.' }
    if ($SomenteListar) {
        Write-Host 'Modo SomenteListar: nenhum arquivo foi exportado.' -ForegroundColor Yellow
        return
    }
    while ($true) {
        $confirmacaoExportacao = (Read-Host 'Deseja exportar os arquivos? (S/N)').Trim()
        $respostaNormalizada = (Remover-Acentos $confirmacaoExportacao).ToLowerInvariant()
        if ($respostaNormalizada -in @('s', 'sim')) { break }
        if ($respostaNormalizada -in @('n', 'nao')) {
            $confirmacaoExportacao = 'N'
            break
        }
        Write-Host 'Resposta invalida. Digite S para Sim ou N para Nao.' -ForegroundColor Yellow
    }
    if ($confirmacaoExportacao -eq 'N') {
        Write-Host 'Operacao cancelada. O relatorio foi mantido.' -ForegroundColor Yellow
        return
    }

    $log = New-Object 'System.Collections.Generic.List[object]'
    $concessaoLocal = Converter-NomePastaLocal $nomeConcessao
    $destinoConcessao = Join-Path $Destino $concessaoLocal
    if (-not (Test-Path -LiteralPath $destinoConcessao -PathType Container)) {
        New-Item -ItemType Directory -Path $destinoConcessao -Force | Out-Null
    }

    $tempoExportacao = [Diagnostics.Stopwatch]::StartNew()
    $totalRetornado = 0
    $totalFalhas = 0
    $caminhoLog = Join-Path $Destino "Log_Exportacao_$dataHora.csv"
    Write-Etapa "Iniciando exportacao de $($areas.Count) area(s)."
    Write-Host 'O percentual acompanha areas processadas, nao arquivos ou bytes copiados.'
    Write-Host 'Mensagens como Returned ... matches sao emitidas pelo modulo; nao comprovam a copia dos arquivos.'
    Write-Host "Log atualizado ao terminar cada area: $caminhoLog"

    for ($i = 0; $i -lt $areas.Count; $i++) {
        $area = $areas[$i]
        $inicio = Get-Date
        $tempoArea = [Diagnostics.Stopwatch]::StartNew()
        Write-Host ''
        Write-Etapa "Area $($i + 1)/$($areas.Count) | $($area.Projeto) | ID $($area.AreaID)"
        Write-Host "Origem: $($area.CaminhoPW)"
        Write-Host "Destino base: $destinoConcessao"
        Write-Progress -Id 1 -Activity 'Exportando Areas de Transferencia' -Status "$i/$($areas.Count) processadas | Atual: $($area.Projeto)" -PercentComplete (($i / $areas.Count) * 100)
        try {
            if ([string]::IsNullOrWhiteSpace($area.CaminhoPW)) {
                throw "Caminho ProjectWise vazio para a Area de Transferencia (ID $($area.AreaID))."
            }

            # Caminho mais rapido: uma unica chamada nativa por Area de
            # Transferencia. Sem -JustOneFolder, inclui toda a arvore abaixo.
            Write-Etapa 'Exportacao em andamento. Aguardando o comando do ProjectWise retornar...'
            Write-Host 'Durante esta chamada, este script nao mede o avanco por arquivo; uma pausa nas mensagens nao indica, por si so, travamento.'
            Write-Progress `
                -Id 2 `
                -ParentId 1 `
                -Activity "Exportando documentos de $($area.Projeto)" `
                -Status 'ProjectWise processando a arvore da pasta...' `
                -PercentComplete -1

            $resultadoExportacao = @(Export-PWDocuments `
                -ProjectWiseFolder $area.CaminhoPW `
                -OutputFolder $destinoConcessao `
                -UseUnderscores `
                -ErrorAction Stop)

            $documentosRetornados = $resultadoExportacao.Count
            Write-Progress -Id 2 -ParentId 1 -Activity "Exportando documentos de $($area.Projeto)" -Completed
            Write-Etapa "Comando finalizado sem erro terminante. Objetos retornados: $documentosRetornados (sem verificacao individual da copia)."
            $status = 'Concluido'
            $erro = ''
        }
        catch {
            Write-Progress -Id 2 -ParentId 1 -Activity "Exportando documentos de $($area.Projeto)" -Completed
            $documentosRetornados = 0
            $status = 'Erro'
            $erro = $_.Exception.Message
            $totalFalhas++
            Write-Host "Falha ao exportar '$($area.Projeto)': $erro" -ForegroundColor Red
        }
        $tempoArea.Stop()
        $totalRetornado += $documentosRetornados
        $processadas = $i + 1
        $pendentes = $areas.Count - $processadas
        Write-Host ("Status: {0} | Duracao da area: {1:hh\:mm\:ss} | Tempo total: {2:hh\:mm\:ss}" -f $status, $tempoArea.Elapsed, $tempoExportacao.Elapsed)
        Write-Host "Areas processadas: $processadas/$($areas.Count) | Pendentes: $pendentes | Falhas: $totalFalhas"
        Write-Progress -Id 1 -Activity 'Exportando Areas de Transferencia' -Status "$processadas/$($areas.Count) processadas | $pendentes pendentes | $totalFalhas falhas" -PercentComplete (($processadas / $areas.Count) * 100)
        $log.Add([pscustomobject]@{
            Concessao = $area.Concessao; Projeto = $area.Projeto; CaminhoPW = $area.CaminhoPW
            DestinoLocal = $destinoConcessao; Status = $status; DocumentosRetornados = $documentosRetornados
            Inicio = $inicio; Fim = Get-Date; Erro = $erro
            AreaID = $area.AreaID; DuracaoSegundos = [Math]::Round($tempoArea.Elapsed.TotalSeconds, 2)
        })
        $log.ToArray() | Export-Csv -LiteralPath $caminhoLog -NoTypeInformation -Encoding UTF8
    }
    $tempoExportacao.Stop()
    Write-Progress -Id 1 -Activity 'Exportando Areas de Transferencia' -Completed
    $falhas = @($log | Where-Object Status -eq 'Erro')
    $corResumo = if ($falhas.Count -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host "Exportacao finalizada. Areas sem erro terminante: $($log.Count - $falhas.Count). Areas com falha: $($falhas.Count)." -ForegroundColor $corResumo
    Write-Host ("Tempo total: {0:hh\:mm\:ss} | Objetos retornados nas areas sem erro: {1}" -f $tempoExportacao.Elapsed, $totalRetornado)
    Write-Host "Detalhes por area: $caminhoLog"
}
finally {
    Write-Progress -Id 2 -Activity 'Exportando documentos' -Completed
    Write-Progress -Id 1 -Activity 'Exportando Areas de Transferencia' -Completed
    if ($script:LoginRealizado) {
        try { Undo-PWLogin -ErrorAction Stop | Out-Null }
        catch { Write-Host "Nao foi possivel encerrar a sessao: $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}
