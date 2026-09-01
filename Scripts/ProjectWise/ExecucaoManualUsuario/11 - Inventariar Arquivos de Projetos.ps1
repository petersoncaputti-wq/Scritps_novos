<#
.SYNOPSIS
    Gera um inventario dos documentos de um ou mais projetos do ProjectWise.

.DESCRIPTION
    Permite selecionar uma concessao e um ou mais projetos. Todos os documentos
    encontrados na arvore dos projetos sao mantidos no relatorio, inclusive
    nomes repetidos. O Excel sinaliza repeticoes dentro do mesmo projeto e no
    conjunto de projetos selecionados.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\11 - Inventariar Arquivos de Projetos.ps1"

.EXAMPLE
    .\11 - Inventariar Arquivos de Projetos.ps1 -IncluirVersoes

.EXAMPLE
    .\11 - Inventariar Arquivos de Projetos.ps1 -ArquivoExistente "C:\Relatorios\Inventario_Anterior.xlsx"

.EXAMPLE
    .\11 - Inventariar Arquivos de Projetos.ps1 -Reconciliar
#>

[CmdletBinding()]
param(
    # Vazio por padrao para que a janela do ProjectWise permita selecionar o
    # datasource. Informe este parametro somente para preselecionar um datasource.
    [string]$DatasourceName = '',
    [string]$PastaSaida,
    [string]$ArquivoExistente,
    [switch]$IncluirVersoes,
    [switch]$Reconciliar
)

# O PWPS_DAB 24 altera ThreadOptions durante a importacao e exige uma runspace
# MTA. Terminais do VS Code e chamadas por dot sourcing podem estar em STA.
# Nesse caso, reinicia este mesmo arquivo em uma janela MTA independente. Isso
# evita encerrar o PowerShell Editor Services quando o arquivo e executado pelo
# terminal especial da extensao PowerShell do VS Code.
$estadoApartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
if ($estadoApartment -ne [System.Threading.ApartmentState]::MTA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'O PWPS_DAB exige MTA. Execute o script salvo em arquivo usando powershell.exe -MTA -File.'
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
        '-File', (Proteger-ArgumentoProcesso $PSCommandPath)
    )
    if (-not [string]::IsNullOrWhiteSpace($DatasourceName)) {
        $partesArgumento += @('-DatasourceName', (Proteger-ArgumentoProcesso $DatasourceName))
    }
    if (-not [string]::IsNullOrWhiteSpace($PastaSaida)) {
        $partesArgumento += @('-PastaSaida', (Proteger-ArgumentoProcesso $PastaSaida))
    }
    if (-not [string]::IsNullOrWhiteSpace($ArquivoExistente)) {
        $partesArgumento += @('-ArquivoExistente', (Proteger-ArgumentoProcesso $ArquivoExistente))
    }
    if ($IncluirVersoes) {
        $partesArgumento += '-IncluirVersoes'
    }
    if ($Reconciliar) {
        $partesArgumento += '-Reconciliar'
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

function Obter-Valor {
    param([object]$Objeto, [string[]]$Nomes)

    if ($null -eq $Objeto) { return $null }
    foreach ($nome in $Nomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            return $propriedade.Value
        }
    }
    return $null
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
    return Obter-Texto -Objeto $Pasta -Nomes @('ProjectID', 'ProjectId', 'FolderID', 'FolderId', 'ID', 'Id')
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
    return @(
        Get-PWFoldersImmediateChildren `
            -FolderID $FolderId `
            -WarningAction SilentlyContinue `
            -ErrorAction Stop |
            Where-Object { $null -ne $_ }
    )
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

function Selecionar-ModoOperacao {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ' INVENTARIO DE ARQUIVOS PROJECTWISE' -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host '1 - Realizar nova extracao'
    Write-Host '2 - Consolidar uma extracao existente'
    Write-Host '0 - Encerrar'

    while ($true) {
        $opcao = (Read-Host 'Escolha uma opcao').Trim()
        switch ($opcao) {
            '1' { return 'Extracao' }
            '2' { return 'Consolidacao' }
            '0' { return 'Encerrar' }
            default { Write-Host 'Informe 1, 2 ou 0.' -ForegroundColor Yellow }
        }
    }
}

function Selecionar-ArquivoInventario {
    Write-Host ''
    Write-Host 'Selecao da extracao existente' -ForegroundColor Cyan
    Write-Host '1 - Informar o caminho completo'
    Write-Host '2 - Listar inventarios da Area de Trabalho'

    while ($true) {
        $opcao = (Read-Host 'Escolha uma opcao (1/2)').Trim()
        if ($opcao -eq '1') {
            $caminho = (Read-Host 'Caminho completo do arquivo .xlsx').Trim().Trim('"')
            if (Test-Path -LiteralPath $caminho -PathType Leaf) { return (Resolve-Path -LiteralPath $caminho).Path }
            Write-Host 'Arquivo nao encontrado. Tente novamente.' -ForegroundColor Yellow
            continue
        }

        if ($opcao -eq '2') {
            $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
            $arquivos = @(
                Get-ChildItem -LiteralPath $desktop -Filter 'Inventario_*.xlsx' -File -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending
            )
            if ($arquivos.Count -eq 0) {
                Write-Host "Nenhum inventario encontrado em: $desktop" -ForegroundColor Yellow
                continue
            }

            Write-Host ''
            Write-Host 'Inventarios encontrados:' -ForegroundColor Cyan
            for ($i = 0; $i -lt $arquivos.Count; $i++) {
                Write-Host ("{0:000}. {1} | {2}" -f ($i + 1), $arquivos[$i].Name, $arquivos[$i].LastWriteTime.ToString('dd/MM/yyyy HH:mm'))
            }

            while ($true) {
                $entrada = (Read-Host 'Digite o numero do arquivo').Trim()
                $numero = 0
                if ([int]::TryParse($entrada, [ref]$numero) -and $numero -ge 1 -and $numero -le $arquivos.Count) {
                    return $arquivos[$numero - 1].FullName
                }
                Write-Host "Informe um numero entre 1 e $($arquivos.Count)." -ForegroundColor Yellow
            }
        }

        Write-Host 'Informe 1 ou 2.' -ForegroundColor Yellow
    }
}

function Selecionar-PastaSaida {
    $pastaPadrao = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
    $pastaPadraoEscapada = $pastaPadrao.Replace("'", "''")
    $codigoSeletor = @"
Add-Type -AssemblyName System.Windows.Forms
`$dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
`$dialogo.Description = 'Selecione a pasta onde o inventario sera salvo'
`$dialogo.ShowNewFolderButton = `$true
`$dialogo.SelectedPath = '$pastaPadraoEscapada'
if (`$dialogo.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    [Console]::Out.Write(`$dialogo.SelectedPath)
}
"@
    $comandoCodificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($codigoSeletor))
    $executavelPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    Write-Etapa 'Abrindo a janela para selecionar a pasta de saida...'
    $resultado = & $executavelPowerShell -NoProfile -STA -EncodedCommand $comandoCodificado
    $caminho = ($resultado | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($caminho)) { return $null }
    if (-not (Test-Path -LiteralPath $caminho -PathType Container)) {
        throw "A pasta selecionada nao esta acessivel: $caminho"
    }
    return (Resolve-Path -LiteralPath $caminho).Path
}

function Converter-Int64 {
    param([object]$Valor)

    if ($null -eq $Valor) { return [int64]0 }
    $numero = [int64]0
    if ([int64]::TryParse($Valor.ToString(), [ref]$numero)) { return $numero }
    return [int64]0
}

function Obter-CaminhoDocumento {
    param(
        [object]$Documento,
        [string]$NomeArquivo,
        [hashtable]$MapaPastas
    )

    $folderId = Obter-Texto $Documento @('ProjectID', 'ProjectId', 'FolderID', 'FolderId')
    if (-not [string]::IsNullOrWhiteSpace($folderId) -and $MapaPastas.ContainsKey($folderId)) {
        return $MapaPastas[$folderId].TrimEnd('\') + '\' + $NomeArquivo
    }

    $caminho = Obter-Texto $Documento @('FullPath', 'DocumentPath', 'FolderPath', 'Path')
    if ([string]::IsNullOrWhiteSpace($caminho)) { return '' }
    if (-not [string]::IsNullOrWhiteSpace($NomeArquivo) -and $caminho.EndsWith($NomeArquivo, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $caminho
    }
    return ($caminho.TrimEnd('\') + '\' + $NomeArquivo)
}

function Obter-DocumentosProjeto {
    param([string]$FolderId, [switch]$ComVersoes)

    $parametros = @{
        FolderID    = [int]$FolderId
        ErrorAction = 'Stop'
    }
    if ($ComVersoes) { $parametros.GetVersionsToo = $true }
    return @(Get-PWDocumentsBySearch @parametros | Where-Object { $null -ne $_ })
}

function Obter-MapaPastasProjeto {
    param([object]$Projeto, [string]$CaminhoProjeto)

    $mapa = @{}
    $fila = New-Object System.Collections.Queue
    $idProjeto = Obter-IdPasta $Projeto
    if ([string]::IsNullOrWhiteSpace($idProjeto)) { throw 'FolderID do projeto nao encontrado.' }

    $mapa[$idProjeto] = $CaminhoProjeto
    $fila.Enqueue([PSCustomObject]@{ FolderId = $idProjeto; Caminho = $CaminhoProjeto })

    while ($fila.Count -gt 0) {
        $pastaAtual = $fila.Dequeue()
        foreach ($filho in @(Obter-FilhosPasta -FolderId $pastaAtual.FolderId)) {
            $idFilho = Obter-IdPasta $filho
            if ([string]::IsNullOrWhiteSpace($idFilho) -or $mapa.ContainsKey($idFilho)) { continue }

            # Caminho logico do ProjectWise: preserva '/' sem usar System.IO.
            $caminhoFilho = $pastaAtual.Caminho.TrimEnd('\') + '\' + (Obter-NomePasta $filho)
            $mapa[$idFilho] = $caminhoFilho
            $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Caminho = $caminhoFilho })
        }
    }
    return $mapa
}

function Nova-LinhaArquivo {
    param(
        [object]$Documento,
        [string]$Concessao,
        [string]$Projeto,
        [hashtable]$MapaPastas
    )

    $nomeArquivoFisico = Obter-Texto $Documento @('FileName', 'Filename')
    $nomeDocumento = Obter-Texto $Documento @('Name', 'DocumentName')
    $nomeExibicao = $nomeArquivoFisico
    if ([string]::IsNullOrWhiteSpace($nomeExibicao)) { $nomeExibicao = $nomeDocumento }
    $tamanhoBytes = Converter-Int64 (Obter-Valor $Documento @('FileSize', 'Filesize', 'Size', 'FileLength'))

    $folderId = Obter-Texto $Documento @('ProjectID', 'ProjectId', 'FolderID', 'FolderId')
    $documentId = Obter-Texto $Documento @('DocumentID', 'DocumentId', 'ID', 'Id')

    return [PSCustomObject][ordered]@{
        Concessao                 = $Concessao
        Projeto                   = $Projeto
        FolderID                  = $folderId
        DocumentID                = $documentId
        'Chave do documento'      = "$folderId-$documentId"
        'Nome do documento'       = $nomeDocumento
        'Nome do arquivo'         = $nomeExibicao
        'Nome fisico do arquivo'  = $nomeArquivoFisico
        'Possui arquivo fisico?'  = if ([string]::IsNullOrWhiteSpace($nomeArquivoFisico)) { 'Nao' } else { 'Sim' }
        'Documento logico?'       = if ([string]::IsNullOrWhiteSpace($nomeArquivoFisico)) { 'Sim' } else { 'Nao' }
        Caminho                   = Obter-CaminhoDocumento -Documento $Documento -NomeArquivo $nomeExibicao -MapaPastas $MapaPastas
        'Tamanho (bytes)'         = $tamanhoBytes
        'Tamanho (MB)'            = [math]::Round(($tamanhoBytes / 1MB), 2)
        State                     = Obter-Texto $Documento @('WorkflowState', 'StateName', 'State')
        Versao                    = Obter-Texto $Documento @('Version', 'DocumentVersion', 'VersionName')
        'Duplicado no projeto?'   = 'Nao'
        'Ocorrencias no projeto'  = 1
        'Duplicado na selecao?'   = 'Nao'
        'Ocorrencias na selecao'  = 1
    }
}

function Obter-ChaveNome {
    param([string]$Nome)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return '' }
    return $Nome.Trim().ToUpperInvariant()
}

function Marcar-Duplicidades {
    param([System.Collections.Generic.List[object]]$Linhas)

    $porProjeto = @{}
    $naSelecao = @{}

    foreach ($linha in $Linhas) {
        $nome = Obter-ChaveNome $linha.'Nome do arquivo'
        if ($nome -eq '') { continue }
        $chaveProjeto = ('{0}|{1}' -f $linha.Projeto.ToUpperInvariant(), $nome)
        if (-not $porProjeto.ContainsKey($chaveProjeto)) { $porProjeto[$chaveProjeto] = 0 }
        if (-not $naSelecao.ContainsKey($nome)) { $naSelecao[$nome] = 0 }
        $porProjeto[$chaveProjeto]++
        $naSelecao[$nome]++
    }

    foreach ($linha in $Linhas) {
        $nome = Obter-ChaveNome $linha.'Nome do arquivo'
        if ($nome -eq '') { continue }
        $chaveProjeto = ('{0}|{1}' -f $linha.Projeto.ToUpperInvariant(), $nome)
        $linha.'Ocorrencias no projeto' = $porProjeto[$chaveProjeto]
        $linha.'Ocorrencias na selecao' = $naSelecao[$nome]
        if ($porProjeto[$chaveProjeto] -gt 1) { $linha.'Duplicado no projeto?' = 'Sim' }
        if ($naSelecao[$nome] -gt 1) { $linha.'Duplicado na selecao?' = 'Sim' }
    }
}

function Exportar-Relatorio {
    param(
        [string]$Arquivo,
        [System.Collections.Generic.List[object]]$Arquivos,
        [System.Collections.Generic.List[object]]$Resumos,
        [System.Collections.Generic.List[object]]$Erros,
        [System.Collections.Generic.List[object]]$Reconciliacao,
        [System.Collections.Generic.List[object]]$Divergencias
    )

    $parametrosExcel = @{
        Path          = $Arquivo
        AutoSize      = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        AutoFilter    = $true
    }

    if ($Arquivos.Count -gt 0) {
        $Arquivos | Export-Excel @parametrosExcel -WorksheetName 'Arquivos' -ClearSheet
    }
    else {
        [PSCustomObject]@{ Aviso = 'Nenhum documento encontrado nos projetos selecionados.' } |
            Export-Excel @parametrosExcel -WorksheetName 'Arquivos' -ClearSheet
    }

    $Resumos | Export-Excel @parametrosExcel -WorksheetName 'Resumo por projeto'

    if ($Reconciliacao.Count -gt 0) {
        $Reconciliacao | Export-Excel @parametrosExcel -WorksheetName 'Reconciliacao'
    }

    if ($Divergencias.Count -gt 0) {
        $Divergencias | Export-Excel @parametrosExcel -WorksheetName 'Divergencias'
    }
    else {
        [PSCustomObject]@{ Status = 'Nenhuma divergencia encontrada no escopo reconciliado.' } |
            Export-Excel @parametrosExcel -WorksheetName 'Divergencias'
    }

    if ($Erros.Count -gt 0) {
        $Erros | Export-Excel @parametrosExcel -WorksheetName 'Erros'
    }
    else {
        [PSCustomObject]@{ Status = 'Nenhum erro registrado.' } |
            Export-Excel @parametrosExcel -WorksheetName 'Erros'
    }
}

function Consolidar-ArquivoExistente {
    param(
        [string]$Arquivo,
        [object]$PastaProjetos,
        [string]$NomeConcessao,
        [string]$DiretorioSaida,
        [switch]$ComVersoes
    )

    if (-not (Test-Path -LiteralPath $Arquivo -PathType Leaf)) {
        throw "Arquivo existente nao encontrado: $Arquivo"
    }
    if ([System.IO.Path]::GetExtension($Arquivo) -ne '.xlsx') {
        throw 'O arquivo existente deve estar no formato .xlsx.'
    }

    Write-Etapa 'Carregando a extracao existente...'
    $dadosExistentes = @(Import-Excel -Path $Arquivo -WorksheetName 'Arquivos' -ErrorAction Stop)
    if ($dadosExistentes.Count -eq 0) { throw 'A aba Arquivos nao possui registros.' }
    $cabecalhos = @($dadosExistentes[0].PSObject.Properties.Name)
    foreach ($obrigatorio in @('FolderID', 'DocumentID', 'Chave do documento', 'Projeto', 'Caminho', 'Tamanho (bytes)')) {
        if ($cabecalhos -notcontains $obrigatorio) {
            throw "A aba Arquivos nao possui a coluna obrigatoria '$obrigatorio'."
        }
    }

    $resumoAnterior = @()
    try { $resumoAnterior = @(Import-Excel -Path $Arquivo -WorksheetName 'Resumo por projeto' -ErrorAction Stop) } catch {}
    $subpastasPorProjeto = @{}
    foreach ($resumo in $resumoAnterior) {
        if (-not [string]::IsNullOrWhiteSpace($resumo.Projeto)) {
            $subpastasPorProjeto[$resumo.Projeto.ToString()] = [int]$resumo.'Total de subpastas'
        }
    }

    $dadosPorChave = @{}
    [decimal]$bytesAnteriores = 0
    foreach ($linhaExistente in $dadosExistentes) {
        $chave = $linhaExistente.'Chave do documento'.ToString()
        if (-not [string]::IsNullOrWhiteSpace($chave)) {
            $dadosPorChave[$chave] = [PSCustomObject]@{
                Projeto = $linhaExistente.Projeto.ToString()
                Caminho = $linhaExistente.Caminho.ToString()
                Linha   = $linhaExistente
            }
        }
        $bytesAnteriores += Converter-Int64 $linhaExistente.'Tamanho (bytes)'
    }
    $quantidadeAnterior = $dadosExistentes.Count
    $dadosExistentes = $null

    Write-Etapa 'Consultando uma unica vez a pasta Projetos completa...'
    $documentosAtuais = @(Obter-DocumentosProjeto -FolderId (Obter-IdPasta $PastaProjetos) -ComVersoes:$ComVersoes)
    $linhas = New-Object 'System.Collections.Generic.List[object]'
    $divergencias = New-Object 'System.Collections.Generic.List[object]'
    $erros = New-Object 'System.Collections.Generic.List[object]'
    $reconciliacao = New-Object 'System.Collections.Generic.List[object]'
    $mapaVazio = @{}
    $chavesAtuais = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    [decimal]$bytesAtuais = 0

    foreach ($documento in $documentosAtuais) {
        $folderId = Obter-Texto $documento @('ProjectID', 'ProjectId', 'FolderID', 'FolderId')
        $documentId = Obter-Texto $documento @('DocumentID', 'DocumentId', 'ID', 'Id')
        $chave = "$folderId-$documentId"
        $null = $chavesAtuais.Add($chave)
        $bytesAtuais += Converter-Int64 (Obter-Valor $documento @('FileSize', 'Filesize', 'Size', 'FileLength'))

        if ($dadosPorChave.ContainsKey($chave)) {
            $referencia = $dadosPorChave[$chave]
            $novaLinha = Nova-LinhaArquivo -Documento $documento -Concessao $NomeConcessao -Projeto $referencia.Projeto -MapaPastas $mapaVazio
            $novaLinha.Caminho = $referencia.Caminho
            $linhas.Add($novaLinha)
        }
        else {
            $novaLinha = Nova-LinhaArquivo -Documento $documento -Concessao $NomeConcessao -Projeto '[Fora da extracao anterior]' -MapaPastas $mapaVazio
            if ([string]::IsNullOrWhiteSpace($novaLinha.Caminho)) {
                $novaLinha.Caminho = "[FolderID $folderId]\$($novaLinha.'Nome do arquivo')"
            }
            $linhas.Add($novaLinha)
            $novaLinha | Add-Member -NotePropertyName 'Tipo de divergencia' -NotePropertyValue 'Existe no ProjectWise e nao estava na extracao anterior'
            $divergencias.Add($novaLinha)
        }
    }

    foreach ($chaveAnterior in $dadosPorChave.Keys) {
        if (-not $chavesAtuais.Contains($chaveAnterior)) {
            $linhaAusente = $dadosPorChave[$chaveAnterior].Linha
            $linhaAusente | Add-Member -NotePropertyName 'Tipo de divergencia' -NotePropertyValue 'Estava na extracao anterior e nao foi localizado na consulta atual' -Force
            $divergencias.Add($linhaAusente)
        }
    }
    $documentosAtuais = $null

    Write-Etapa 'Recalculando resumos e duplicidades...'
    Marcar-Duplicidades -Linhas $linhas
    $resumos = New-Object 'System.Collections.Generic.List[object]'
    foreach ($grupoProjeto in @($linhas | Group-Object Projeto | Sort-Object Name)) {
        $itens = @($grupoProjeto.Group)
        [decimal]$bytesProjeto = ($itens | Measure-Object -Property 'Tamanho (bytes)' -Sum).Sum
        $comArquivo = @($itens | Where-Object { $_.'Possui arquivo fisico?' -eq 'Sim' }).Count
        $semState = @($itens | Where-Object { [string]::IsNullOrWhiteSpace($_.State) }).Count
        $subpastas = 0
        if ($subpastasPorProjeto.ContainsKey($grupoProjeto.Name)) { $subpastas = $subpastasPorProjeto[$grupoProjeto.Name] }
        $resumos.Add([PSCustomObject][ordered]@{
            Concessao             = $NomeConcessao
            Projeto               = $grupoProjeto.Name
            'Total de arquivos'   = $itens.Count
            'Com arquivo fisico'  = $comArquivo
            'Documentos logicos'  = $itens.Count - $comArquivo
            'Sem State'           = $semState
            'Total de subpastas'  = $subpastas
            'Tamanho total (MB)'  = [math]::Round(($bytesProjeto / 1MB), 2)
            Status                = 'Consolidado'
            'Duracao (segundos)'  = ''
        })
    }

    $reconciliacao.Add([PSCustomObject][ordered]@{
        Metrica = 'Documentos'
        'Extracao anterior' = $quantidadeAnterior
        'Consulta atual da pasta Projetos' = $linhas.Count
        Diferenca = $linhas.Count - $quantidadeAnterior
        Observacao = 'Novos registros sao detalhados na aba Divergencias.'
    })
    $reconciliacao.Add([PSCustomObject][ordered]@{
        Metrica = 'Tamanho (bytes)'
        'Extracao anterior' = $bytesAnteriores
        'Consulta atual da pasta Projetos' = $bytesAtuais
        Diferenca = $bytesAtuais - $bytesAnteriores
        Observacao = 'Valores atualizados pela consulta ao ProjectWise.'
    })

    if ([string]::IsNullOrWhiteSpace($DiretorioSaida)) {
        $DiretorioSaida = Split-Path -Parent (Resolve-Path -LiteralPath $Arquivo)
    }
    if (-not (Test-Path -LiteralPath $DiretorioSaida -PathType Container)) {
        throw "Pasta de saida inexistente: $DiretorioSaida"
    }

    $nomeBase = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo)
    $arquivoSaida = Join-Path $DiretorioSaida ("{0}_Consolidado_{1}.xlsx" -f $nomeBase, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Write-Etapa 'Gerando o Excel consolidado...'
    Exportar-Relatorio -Arquivo $arquivoSaida -Arquivos $linhas -Resumos $resumos -Erros $erros -Reconciliacao $reconciliacao -Divergencias $divergencias
    return [PSCustomObject]@{
        Arquivo = $arquivoSaida
        Documentos = $linhas.Count
        Divergencias = $divergencias.Count
    }
}

try {
    Write-Etapa 'Validando dependencias...'
    Import-Module PWPS_DAB -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop

    $parametrosLogin = @{
        UseGui                     = $true
        DoNotCreateWorkingDirectory = $true
        ErrorAction                = 'Stop'
    }
    if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
        Write-Etapa 'Abrindo login do ProjectWise. Selecione o datasource e Bentley IMS...'
    }
    else {
        Write-Etapa "Abrindo login do ProjectWise no datasource $DatasourceName..."
        $parametrosLogin.DatasourceName = $DatasourceName
    }

    $loginOk = New-PWLogin @parametrosLogin
    if (-not $loginOk) {
        throw 'O login do ProjectWise foi cancelado ou nao foi concluido.'
    }
    $script:LoginRealizado = $true
    Write-Etapa 'Login concluido. Iniciando leitura do datasource...'

    if (-not [string]::IsNullOrWhiteSpace($ArquivoExistente)) {
        $modoOperacao = 'Consolidacao'
        Write-Host ''
        Write-Host 'Modo selecionado por parametro: CONSOLIDACAO' -ForegroundColor Green
    }
    else {
        $modoOperacao = Selecionar-ModoOperacao
        if ($modoOperacao -eq 'Encerrar') {
            Write-Host 'Operacao encerrada pelo usuario.' -ForegroundColor Yellow
            return
        }
        if ($modoOperacao -eq 'Consolidacao') {
            $ArquivoExistente = Selecionar-ArquivoInventario
            Write-Host ''
            Write-Host 'Modo selecionado: CONSOLIDACAO' -ForegroundColor Green
            Write-Host 'O arquivo original nao sera alterado.'
            Write-Host "Arquivo selecionado: $ArquivoExistente"
        }
        else {
            Write-Host ''
            Write-Host 'Modo selecionado: NOVA EXTRACAO' -ForegroundColor Green
        }
    }

    if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
        $PastaSaida = Selecionar-PastaSaida
        if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
            Write-Host 'Selecao da pasta de saida cancelada. Nenhum processamento foi iniciado.' -ForegroundColor Yellow
            return
        }
    }
    elseif (-not (Test-Path -LiteralPath $PastaSaida -PathType Container)) {
        throw "Pasta de saida inexistente: $PastaSaida"
    }
    Write-Host "Pasta de saida: $PastaSaida" -ForegroundColor Green

    Write-Etapa 'Localizando concessoes...'
    $pastaEngenharia = Localizar-Pasta -Pastas (Obter-FilhosRaiz) -Nomes $NomesEngenharia -Descricao 'Pasta Engenharia'
    $concessoes = @(Obter-FilhosPasta (Obter-IdPasta $pastaEngenharia) | Sort-Object { (Obter-RotuloPasta $_).ToLowerInvariant() })
    if ($concessoes.Count -eq 0) { throw 'Nenhuma concessao encontrada.' }

    $concessao = Selecionar-UmItem -Itens $concessoes -Titulo 'Concessoes encontradas'
    $nomeConcessao = Obter-RotuloPasta $concessao

    Write-Etapa "Localizando projetos de '$nomeConcessao'..."
    $filhosConcessao = Obter-FilhosPasta (Obter-IdPasta $concessao)
    $pastaProjetos = Localizar-Pasta -Pastas $filhosConcessao -Nomes $NomesProjetos -Descricao 'Pasta Projetos'

    if ($modoOperacao -eq 'Consolidacao') {
        Write-Host ''
        Write-Host 'Modo de consolidacao de extracao existente.' -ForegroundColor Green
        Write-Host "Arquivo de origem: $ArquivoExistente"
        $confirmacaoConsolidacao = (Read-Host "Consolidar usando a pasta Projetos de '$nomeConcessao'? (S/N)").Trim()
        if ($confirmacaoConsolidacao -notmatch '^[Ss]$') { throw 'Consolidacao cancelada pelo usuario.' }

        $resultadoConsolidacao = Consolidar-ArquivoExistente `
            -Arquivo $ArquivoExistente `
            -PastaProjetos $pastaProjetos `
            -NomeConcessao $nomeConcessao `
            -DiretorioSaida $PastaSaida `
            -ComVersoes:$IncluirVersoes

        Write-Host ''
        Write-Host 'Consolidacao concluida.' -ForegroundColor Green
        Write-Host "Documentos no arquivo consolidado: $($resultadoConsolidacao.Documentos)"
        Write-Host "Divergencias encontradas: $($resultadoConsolidacao.Divergencias)"
        Write-Host "Arquivo: $($resultadoConsolidacao.Arquivo)"
        return
    }

    $projetos = @(Obter-FilhosPasta (Obter-IdPasta $pastaProjetos) | Sort-Object { (Obter-RotuloPasta $_).ToLowerInvariant() })
    if ($projetos.Count -eq 0) { throw 'Nenhum projeto encontrado na concessao selecionada.' }

    $projetosSelecionados = @(Selecionar-VariosItens -Itens $projetos -Titulo 'Projetos encontrados')
    Write-Host ''
    Write-Host "Projetos selecionados: $($projetosSelecionados.Count)" -ForegroundColor Green
    $projetosSelecionados | ForEach-Object { Write-Host " - $(Obter-RotuloPasta $_)" }
    $confirmacao = (Read-Host 'Continuar com a extracao? (S/N)').Trim()
    if ($confirmacao -notmatch '^[Ss]$') { throw 'Operacao cancelada pelo usuario.' }

    $linhas = New-Object 'System.Collections.Generic.List[object]'
    $resumos = New-Object 'System.Collections.Generic.List[object]'
    $erros = New-Object 'System.Collections.Generic.List[object]'
    $reconciliacao = New-Object 'System.Collections.Generic.List[object]'
    $divergencias = New-Object 'System.Collections.Generic.List[object]'
    $totalProjetos = $projetosSelecionados.Count

    for ($i = 0; $i -lt $totalProjetos; $i++) {
        $projeto = $projetosSelecionados[$i]
        $nomeProjeto = Obter-RotuloPasta $projeto
        $inicioProjeto = Get-Date
        Write-Progress -Activity 'Inventario ProjectWise' -Status "$($i + 1)/$totalProjetos - $nomeProjeto" -PercentComplete ((($i + 1) / $totalProjetos) * 100)
        Write-Etapa "[$($i + 1)/$totalProjetos] Consultando '$nomeProjeto'..."

        try {
            Write-Etapa "[$($i + 1)/$totalProjetos] Mapeando pastas de '$nomeProjeto'..."
            $caminhoProjeto = @(
                (Obter-NomePasta $pastaEngenharia),
                (Obter-NomePasta $concessao),
                (Obter-NomePasta $pastaProjetos),
                (Obter-NomePasta $projeto)
            ) -join '\'
            $mapaPastas = Obter-MapaPastasProjeto -Projeto $projeto -CaminhoProjeto $caminhoProjeto

            $documentos = @(Obter-DocumentosProjeto -FolderId (Obter-IdPasta $projeto) -ComVersoes:$IncluirVersoes)
            $inicioLinhas = $linhas.Count
            foreach ($documento in $documentos) {
                $linhas.Add((Nova-LinhaArquivo -Documento $documento -Concessao $nomeConcessao -Projeto $nomeProjeto -MapaPastas $mapaPastas))
            }
            $linhasProjeto = @($linhas | Select-Object -Skip $inicioLinhas)
            $tamanhoTotal = [int64](($linhasProjeto | Measure-Object -Property 'Tamanho (bytes)' -Sum).Sum)
            $comArquivoFisico = @($linhasProjeto | Where-Object { $_.'Possui arquivo fisico?' -eq 'Sim' }).Count
            $semArquivoFisico = $linhasProjeto.Count - $comArquivoFisico
            $semState = @($linhasProjeto | Where-Object { [string]::IsNullOrWhiteSpace($_.State) }).Count
            $resumos.Add([PSCustomObject][ordered]@{
                Concessao          = $nomeConcessao
                Projeto            = $nomeProjeto
                'Total de arquivos'= $documentos.Count
                'Com arquivo fisico' = $comArquivoFisico
                'Documentos logicos' = $semArquivoFisico
                'Sem State'          = $semState
                'Total de subpastas' = [math]::Max(0, ($mapaPastas.Count - 1))
                'Tamanho total (MB)' = [math]::Round(($tamanhoTotal / 1MB), 2)
                Status             = 'Concluido'
                'Duracao (segundos)' = [math]::Round(((Get-Date) - $inicioProjeto).TotalSeconds, 1)
            })
        }
        catch {
            $mensagem = $_.Exception.Message
            $erros.Add([PSCustomObject][ordered]@{
                Concessao = $nomeConcessao
                Projeto   = $nomeProjeto
                Etapa     = 'Consulta de documentos'
                Mensagem  = $mensagem
            })
            $resumos.Add([PSCustomObject][ordered]@{
                Concessao            = $nomeConcessao
                Projeto              = $nomeProjeto
                'Total de arquivos'  = 0
                'Com arquivo fisico' = 0
                'Documentos logicos' = 0
                'Sem State'          = 0
                'Total de subpastas' = 0
                'Tamanho total (MB)' = 0
                Status               = 'Erro'
                'Duracao (segundos)' = [math]::Round(((Get-Date) - $inicioProjeto).TotalSeconds, 1)
            })
            Write-Host "Falha em '$nomeProjeto': $mensagem" -ForegroundColor Yellow
        }
    }
    Write-Progress -Activity 'Inventario ProjectWise' -Completed

    if ($Reconciliar -and $projetosSelecionados.Count -eq $projetos.Count -and $erros.Count -eq 0) {
        Write-Etapa 'Reconciliando projetos com a arvore completa da pasta Projetos...'
        try {
            $caminhoPastaProjetos = @(
                (Obter-NomePasta $pastaEngenharia),
                (Obter-NomePasta $concessao),
                (Obter-NomePasta $pastaProjetos)
            ) -join '\'
            $mapaEscopo = Obter-MapaPastasProjeto -Projeto $pastaProjetos -CaminhoProjeto $caminhoPastaProjetos
            $documentosEscopo = @(Obter-DocumentosProjeto -FolderId (Obter-IdPasta $pastaProjetos) -ComVersoes:$IncluirVersoes)

            $chavesSelecionadas = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($linha in $linhas) { $null = $chavesSelecionadas.Add($linha.'Chave do documento') }

            [decimal]$bytesEscopo = 0
            foreach ($documentoEscopo in $documentosEscopo) {
                $folderIdEscopo = Obter-Texto $documentoEscopo @('ProjectID', 'ProjectId', 'FolderID', 'FolderId')
                $documentIdEscopo = Obter-Texto $documentoEscopo @('DocumentID', 'DocumentId', 'ID', 'Id')
                $chaveEscopo = "$folderIdEscopo-$documentIdEscopo"
                $bytesEscopo += Converter-Int64 (Obter-Valor $documentoEscopo @('FileSize', 'Filesize', 'Size', 'FileLength'))

                if (-not $chavesSelecionadas.Contains($chaveEscopo)) {
                    $linhaDivergente = Nova-LinhaArquivo `
                        -Documento $documentoEscopo `
                        -Concessao $nomeConcessao `
                        -Projeto '[Fora das raizes dos projetos selecionados]' `
                        -MapaPastas $mapaEscopo
                    $linhaDivergente | Add-Member -NotePropertyName 'Tipo de divergencia' -NotePropertyValue 'Documento localizado na pasta Projetos, mas fora das raizes selecionadas'
                    $divergencias.Add($linhaDivergente)
                }
            }

            [decimal]$bytesSelecionados = ($linhas | Measure-Object -Property 'Tamanho (bytes)' -Sum).Sum
            $subpastasSelecionadas = [int](($resumos | Measure-Object -Property 'Total de subpastas' -Sum).Sum)
            $pastasNoEscopo = [math]::Max(0, ($mapaEscopo.Count - 1))

            $reconciliacao.Add([PSCustomObject][ordered]@{
                Metrica = 'Documentos'
                'Projetos selecionados' = $linhas.Count
                'Pasta Projetos completa' = $documentosEscopo.Count
                Diferenca = $documentosEscopo.Count - $linhas.Count
                Observacao = 'A diferenca e detalhada na aba Divergencias.'
            })
            $reconciliacao.Add([PSCustomObject][ordered]@{
                Metrica = 'Tamanho (bytes)'
                'Projetos selecionados' = $bytesSelecionados
                'Pasta Projetos completa' = $bytesEscopo
                Diferenca = $bytesEscopo - $bytesSelecionados
                Observacao = 'Soma da propriedade de tamanho retornada pelo PWPS_DAB.'
            })
            $reconciliacao.Add([PSCustomObject][ordered]@{
                Metrica = 'Pastas abaixo do escopo'
                'Projetos selecionados' = $subpastasSelecionadas
                'Pasta Projetos completa' = $pastasNoEscopo
                Diferenca = $pastasNoEscopo - $subpastasSelecionadas
                Observacao = 'O total completo inclui as pastas-raiz dos projetos.'
            })
        }
        catch {
            $mensagemReconciliacao = $_.Exception.Message
            $erros.Add([PSCustomObject][ordered]@{
                Concessao = $nomeConcessao
                Projeto   = '[Reconciliacao da pasta Projetos]'
                Etapa     = 'Reconciliacao'
                Mensagem  = $mensagemReconciliacao
            })
            $reconciliacao.Add([PSCustomObject]@{
                Metrica = 'Status da reconciliacao'
                'Projetos selecionados' = ''
                'Pasta Projetos completa' = ''
                Diferenca = ''
                Observacao = "Falha: $mensagemReconciliacao"
            })
        }
    }
    elseif ($Reconciliar) {
        $reconciliacao.Add([PSCustomObject]@{
            Metrica = 'Status da reconciliacao'
            'Projetos selecionados' = $projetosSelecionados.Count
            'Pasta Projetos completa' = $projetos.Count
            Diferenca = ''
            Observacao = 'Reconciliacao completa disponivel somente quando todos os projetos sao selecionados e nenhum projeto falha.'
        })
    }
    else {
        $reconciliacao.Add([PSCustomObject]@{
            Metrica = 'Status da reconciliacao'
            'Projetos selecionados' = $projetosSelecionados.Count
            'Pasta Projetos completa' = ''
            Diferenca = ''
            Observacao = 'Nao executada. Use o parametro -Reconciliar somente quando precisar de uma auditoria adicional.'
        })
    }

    Write-Etapa 'Calculando duplicidades sem remover registros...'
    Marcar-Duplicidades -Linhas $linhas

    if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
        Write-Etapa 'Solicitando novamente a pasta de saida...'
        $PastaSaida = Selecionar-PastaSaida
    }
    if ([string]::IsNullOrWhiteSpace($PastaSaida)) { throw 'Nenhuma pasta de saida foi selecionada.' }
    if (-not (Test-Path -LiteralPath $PastaSaida -PathType Container)) {
        throw "Pasta de saida inexistente: $PastaSaida"
    }

    $nomeSeguro = $nomeConcessao -replace '[\\/:*?"<>|]', '_'
    $arquivoSaida = Join-Path $PastaSaida ("Inventario_{0}_{1}_Projetos_{2}.xlsx" -f $nomeSeguro, $totalProjetos, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Write-Etapa 'Gerando arquivo Excel...'
    Exportar-Relatorio `
        -Arquivo $arquivoSaida `
        -Arquivos $linhas `
        -Resumos $resumos `
        -Erros $erros `
        -Reconciliacao $reconciliacao `
        -Divergencias $divergencias

    Write-Host ''
    Write-Host 'Inventario concluido.' -ForegroundColor Green
    Write-Host "Documentos exportados: $($linhas.Count)"
    Write-Host "Divergencias de escopo: $($divergencias.Count)"
    Write-Host "Projetos com erro: $($erros.Count)"
    Write-Host "Arquivo: $arquivoSaida"
}
catch {
    Write-Host ''
    Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($script:LoginRealizado) {
        try { Undo-PWLogin -ErrorAction SilentlyContinue } catch {}
    }
    if ($Host.Name -match 'ConsoleHost') {
        $null = Read-Host 'Pressione Enter para encerrar'
    }
}
