<#
.SYNOPSIS
    Limpa o conteudo de Entrada, Saida e Importacao no ProjectWise.
.DESCRIPTION
    Baseado no script 12: seleciona concessao e projetos e localiza suas Areas
    de Transferencia. Preserva Entrada, Saida e Importacao, removendo documentos
    e subpastas contidos nelas. Somente filhos diretos da Area sao selecionados.
    Os nomes sao comparados sem acentos, ignorando maiusculas e espacos externos.
.PARAMETER SomenteListar
    Lista os alvos e quantidades sem excluir dados.
#>
[CmdletBinding()]
param(
    [string]$DatasourceName = '',
    [switch]$SomenteListar
)
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::MTA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Salve o script e execute o arquivo completo para iniciar a sessao MTA.'
    }
    # Literais PowerShell e comando Unicode preservam espacos, acentos e aspas.
    $caminhoLiteral = $PSCommandPath.Replace("'", "''")
    $datasourceLiteral = $DatasourceName.Replace("'", "''")
    $comandoMTA = "& '$caminhoLiteral' -DatasourceName '$datasourceLiteral'"
    if ($SomenteListar) { $comandoMTA += ' -SomenteListar' }
    $comandoCodificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($comandoMTA))
    $executavelPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    # Janela interativa necessaria para login, selecao e confirmacao da limpeza.
    # NoExit mantem o resultado ou eventual erro visivel ao terminar.
    $processo = Start-Process -FilePath $executavelPowerShell -ArgumentList @(
        '-NoProfile', '-NoExit', '-MTA', '-ExecutionPolicy', 'Bypass',
        '-EncodedCommand', $comandoCodificado
    ) -WindowStyle Normal -PassThru -ErrorAction Stop
    Write-Host "PowerShell MTA iniciado (PID $($processo.Id)). Continue na nova janela." -ForegroundColor Green
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

function Testar-NomeAreaTransferencia {
    param([string]$Nome)
    $normalizado = (Remover-Acentos $Nome).Trim().ToLowerInvariant()
    # Aceita a grafia correta "transferencia", a variante "tranferencia" e
    # espacos antes/depois do nome (removidos pelo Trim acima).
    return $normalizado -match '^(\d+\s*-\s*)?area de trans?ferencia$'
}

function Localizar-AreasTransferenciaNoProjeto {
    param([object]$Projeto)
    # A Area de Transferencia fica diretamente dentro do projeto.
    # Uma unica consulta de filhos imediatos; nao percorre disciplinas.
    $idProjeto = Obter-IdPastaValido $Projeto
    Write-Etapa "Consultando Area de Transferencia diretamente no projeto (ID $idProjeto)..."
    $areas = @(Obter-FilhosPasta $idProjeto | Where-Object {
        Testar-NomeAreaTransferencia (Obter-NomePasta $_)
    })
    if ($areas.Count -gt 1) {
        throw "Mais de uma Area de Transferencia diretamente no projeto ID $idProjeto. Verifique a estrutura antes de limpar."
    }
    if ($areas.Count -eq 1) {
        $idArea = Obter-IdPastaValido $areas[0]
        Write-Etapa "Area encontrada: $(Obter-NomePasta $areas[0]) (ID $idArea). Busca concluida."
    }
    else {
        Write-Etapa "Area de Transferencia nao encontrada diretamente no projeto ID $idProjeto."
    }
    return $areas
}

function Obter-IdPastaValido {
    param([object]$Pasta)
    $idPasta = 0
    if (-not [int]::TryParse((Obter-IdPasta $Pasta), [ref]$idPasta) -or $idPasta -le 0) {
        $detalhes = Obter-DiagnosticoObjeto $Pasta
        throw "Pasta sem ID ProjectWise positivo. Operacao interrompida. Item: $detalhes"
    }
    return $idPasta
}

function Obter-DiagnosticoObjeto {
    param([AllowNull()][object]$Objeto)
    if ($null -eq $Objeto) { return '{"Tipo":"null"}' }
    $propriedades = [ordered]@{}
    foreach ($propriedade in $Objeto.PSObject.Properties) {
        try { $propriedades[$propriedade.Name] = [string]$propriedade.Value }
        catch { $propriedades[$propriedade.Name] = '<falha ao ler propriedade>' }
    }
    return ([ordered]@{
        Tipo = $Objeto.GetType().FullName
        Valor = [string]$Objeto
        Propriedades = $propriedades
    } | ConvertTo-Json -Depth 4 -Compress)
}

function Testar-PastaLimpeza {
    param([object]$Pasta)
    return (Remover-Acentos (Obter-NomePasta $Pasta)).Trim().ToLowerInvariant() -in @('entrada', 'saida', 'importacao')
}

function Obter-InventarioTamanho {
    param([int]$FolderID)
    $itens = @{}
    foreach ($doc in (Get-PWDocumentsBySearch -FolderID $FolderID -GetVersionsToo -ErrorAction Stop)) {
        if ($null -eq $doc) { continue }
        $guid = [guid]$doc.DocumentGUID
        if ($guid -eq [guid]::Empty) { throw 'Documento sem GUID para medir a exclusao.' }
        if ($null -eq $doc.PSObject.Properties['FileSize'] -or $null -eq $doc.FileSize) {
            throw 'Documento sem FileSize para medir a exclusao.'
        }
        $bytes = [decimal]$doc.FileSize
        if ($bytes -lt 0) { throw 'Tamanho de documento invalido.' }
        $itens[('{0}:{1}' -f $guid, $doc.VersionSequence)] = $bytes
    }
    return ,$itens
}

function Medir-Exclusao {
    param([hashtable]$Antes, [hashtable]$Depois)
    [decimal]$bytesAntes = 0
    [decimal]$bytesDepois = 0
    [decimal]$bytesExcluidos = 0
    $quantidade = 0
    foreach ($chave in $Antes.Keys) {
        $bytesAntes += $Antes[$chave]
        if (-not $Depois.ContainsKey($chave)) {
            $bytesExcluidos += $Antes[$chave]
            $quantidade++
        }
    }
    foreach ($valor in $Depois.Values) { $bytesDepois += $valor }
    return [pscustomobject]@{
        BytesAntes = $bytesAntes; BytesDepois = $bytesDepois
        BytesExcluidos = $bytesExcluidos
        MiBExcluidos = [math]::Round($bytesExcluidos / 1MB, 6)
        GiBExcluidos = [math]::Round($bytesExcluidos / 1GB, 6)
        DocumentosVersoesExcluidos = $quantidade
    }
}

function Limpar-ConteudoPasta {
    param([object]$Alvo, [Parameter(Mandatory)][object]$Diagnostico)
    $Diagnostico.Etapa = 'Validar pasta protegida'
    $pastaAtual = @(Obter-FilhosPasta $Alvo.AreaID | Where-Object {
        (Obter-IdPasta $_) -eq [string]$Alvo.PastaID -and (Testar-PastaLimpeza $_)
    })
    if ($pastaAtual.Count -ne 1) { throw 'A pasta alvo mudou ou nao pertence mais a Area de Transferencia.' }

    # Sem RemoveFolders: preserva todas as pastas e remove documentos da arvore.
    $Diagnostico.Etapa = 'Remover documentos'
    $resultado = @(Remove-PWFolder -FolderID $Alvo.PastaID -RemoveDocuments -ProceedWithDelete -ErrorAction Stop)
    if ($resultado -contains $false) { throw 'ProjectWise informou falha ao remover documentos.' }

    # RemoveFolders SOMENTE nos filhos, nunca em Entrada/Saida/Importacao.
    $Diagnostico.Etapa = 'Consultar subpastas apos remover documentos'
    $filhosLimpeza = @(Obter-FilhosPasta $Alvo.PastaID)
    $Diagnostico.ItensRetornados = $filhosLimpeza.Count
    Write-Etapa "Consulta de subpastas da pasta $($Alvo.PastaID): $($filhosLimpeza.Count) item(ns)."
    foreach ($filho in $filhosLimpeza) {
        $Diagnostico.Etapa = 'Validar ID da subpasta'
        $Diagnostico.Item = Obter-DiagnosticoObjeto $filho
        $idFilho = Obter-IdPastaValido $filho
        if ($idFilho -eq $Alvo.PastaID -or $idFilho -eq $Alvo.AreaID) {
            throw 'ID de subpasta corresponde a uma pasta protegida.'
        }
        $Diagnostico.Etapa = "Remover subpasta ID $idFilho"
        Write-Etapa "Removendo subpasta $(Obter-NomePasta $filho) (ID $idFilho)..."
        $resultado = @(Remove-PWFolder -FolderID $idFilho -RemoveDocuments -RemoveFolders -ProceedWithDelete -ErrorAction Stop)
        if ($resultado -contains $false) { throw "Falha ao remover subpasta $idFilho." }
    }

    $Diagnostico.Item = ''
    $Diagnostico.Etapa = 'Verificar documentos restantes'
    $restantes = @(Get-PWDocumentsBySearch -FolderID $Alvo.PastaID -GetVersionsToo -ErrorAction Stop | Where-Object { $null -ne $_ })
    $Diagnostico.Etapa = 'Verificar subpastas restantes'
    $subpastas = @(Obter-FilhosPasta $Alvo.PastaID)
    $Diagnostico.Etapa = 'Verificar preservacao da pasta'
    $preservada = @(Obter-FilhosPasta $Alvo.AreaID | Where-Object { (Obter-IdPasta $_) -eq [string]$Alvo.PastaID })
    if ($preservada.Count -ne 1 -or $restantes.Count -gt 0 -or $subpastas.Count -gt 0) {
        throw "Verificacao incompleta: documentos=$($restantes.Count), subpastas=$($subpastas.Count), pasta preservada=$($preservada.Count)."
    }
    $Diagnostico.Etapa = 'Verificacao final concluida'
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


    $alvos = New-Object 'System.Collections.Generic.List[object]'
    $idsVistos = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($projeto in $projetos) {
        Write-Etapa "Procurando Areas de Transferencia em $(Obter-RotuloPasta $projeto)..."
        $areas = @(Localizar-AreasTransferenciaNoProjeto $projeto)
        if ($areas.Count -eq 0) { Write-Host 'Nenhuma Area de Transferencia encontrada.' -ForegroundColor Yellow }
        foreach ($area in $areas) {
            $idArea = Obter-IdPastaValido $area
            Write-Etapa "Consultando Entrada/Saida/Importacao da area ID $idArea..."
            $pastas = @(Obter-FilhosPasta $idArea | Where-Object { Testar-PastaLimpeza $_ })
            if ($pastas.Count -eq 0) { Write-Host "Nenhuma pasta Entrada/Saida/Importacao na area ID $idArea." -ForegroundColor Yellow }
            foreach ($pasta in $pastas) {
                $idPasta = Obter-IdPastaValido $pasta
                if ($idPasta -eq $idArea) { throw 'Alvo corresponde a propria Area de Transferencia.' }
                if (-not $idsVistos.Add($idPasta)) { continue }
                Write-Etapa "Consultando documentos e versoes de $(Obter-NomePasta $pasta) (ID $idPasta), incluindo subpastas. Aguardando ProjectWise..."
                $documentos = @(Get-PWDocumentsBySearch -FolderID $idPasta -GetVersionsToo -ErrorAction Stop | Where-Object { $null -ne $_ })
                Write-Etapa "$($documentos.Count) documentos/versoes retornados. Consultando subpastas..."
                $filhos = @(Obter-FilhosPasta $idPasta)
                Write-Etapa "Pasta analisada: $(Obter-NomePasta $pasta) | $($filhos.Count) subpastas diretas."
                $alvos.Add([pscustomobject]@{
                    Projeto = Obter-RotuloPasta $projeto
                    AreaID = $idArea; PastaID = $idPasta
                    Pasta = Obter-NomePasta $pasta
                    CaminhoPW = Obter-CaminhoPW $pasta
                    DocumentosComVersoes = $documentos.Count
                    SubpastasDiretas = $filhos.Count
                })
            }
        }
    }
    if ($alvos.Count -eq 0) { Write-Host 'Nenhuma pasta para limpar.'; return }
    $alvos.ToArray() | Format-Table Projeto, Pasta, PastaID, DocumentosComVersoes, SubpastasDiretas, CaminhoPW -AutoSize -Wrap | Out-Host
    if ($SomenteListar) { Write-Host 'SomenteListar: nenhum dado foi excluido.'; return }

    Write-Host 'Todo o conteudo listado sera excluido, incluindo documentos, versoes e subpastas.' -ForegroundColor Yellow
    Write-Host 'As pastas Entrada, Saida e Importacao serao preservadas.'
    while ($true) {
        $confirmacao = (Remover-Acentos (Read-Host 'Deseja iniciar a exclusao? (Sim/Nao)')).Trim().ToLowerInvariant()
        if ($confirmacao -in @('s', 'sim')) { break }
        if ($confirmacao -in @('n', 'nao')) {
            Write-Host 'Operacao cancelada.'
            return
        }
        Write-Host 'Resposta invalida. Digite Sim (S) ou Nao (N).' -ForegroundColor Yellow
    }
    $pastaLog = Join-Path $PSScriptRoot 'Logs'
    if (-not (Test-Path -LiteralPath $pastaLog -PathType Container)) {
        New-Item -ItemType Directory -Path $pastaLog -Force | Out-Null
    }
    $arquivoLog = Join-Path $pastaLog ("Log_Limpeza_Areas_Transferencia_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
    $log = New-Object 'System.Collections.Generic.List[object]'
    # Valida a gravacao do log antes das exclusoes.
    'Projeto,AreaID,PastaID,Pasta,CaminhoPW,Status,Erro' | Set-Content -LiteralPath $arquivoLog -Encoding UTF8
    foreach ($alvo in $alvos) {
        $status = 'Concluido'; $erro = ''
        $diagnostico = [pscustomobject]@{ Etapa = 'Inventario anterior'; ItensRetornados = $null; Item = '' }
        $linhaErro = $null; $pilhaErro = ''; $identificadorErro = ''
        $antes = $null; $medicao = $null
        $statusMedicao = 'Nao confirmado'; $erroMedicao = ''
        try {
            Write-Etapa "Limpando $($alvo.Projeto) / $($alvo.Pasta) (ID $($alvo.PastaID))..."
            $antes = Obter-InventarioTamanho $alvo.PastaID
            Limpar-ConteudoPasta -Alvo $alvo -Diagnostico $diagnostico
        }
        catch {
            $status = 'Erro'; $erro = $_.Exception.Message
            $linhaErro = $_.InvocationInfo.ScriptLineNumber
            $pilhaErro = $_.ScriptStackTrace
            $identificadorErro = $_.FullyQualifiedErrorId
            Write-Host "Etapa da falha: $($diagnostico.Etapa) | Linha: $linhaErro" -ForegroundColor Yellow
            Write-Host "Falha na pasta $($alvo.PastaID): $erro" -ForegroundColor Red
        }
        # Mede inclusive exclusoes parciais. Falha de consulta nao significa zero.
        # Tamanho logico FileSize do PW, nao espaco fisico liberado no servidor.
        # Movimentacoes concorrentes de outros usuarios podem afetar a comparacao.
        if ($null -ne $antes) {
            try {
                $depois = Obter-InventarioTamanho $alvo.PastaID
                $medicao = Medir-Exclusao -Antes $antes -Depois $depois
                $statusMedicao = 'Confirmado por inventario'
            }
            catch { $erroMedicao = $_.Exception.Message }
        }
        else { $erroMedicao = 'Inventario anterior indisponivel. Exclusao nao iniciada.' }
        $log.Add([pscustomobject]@{
            Projeto = $alvo.Projeto; AreaID = $alvo.AreaID; PastaID = $alvo.PastaID
            Pasta = $alvo.Pasta; CaminhoPW = $alvo.CaminhoPW; Status = $status; Erro = $erro
            Etapa = $diagnostico.Etapa; ItensSubpastasRetornados = $diagnostico.ItensRetornados
            ItemDiagnostico = $diagnostico.Item; LinhaErro = $linhaErro
            IdentificadorErro = $identificadorErro; PilhaErro = $pilhaErro
            BytesAntes = if ($null -ne $medicao) { $medicao.BytesAntes } else { $null }
            BytesDepois = if ($null -ne $medicao) { $medicao.BytesDepois } else { $null }
            BytesExcluidos = if ($null -ne $medicao) { $medicao.BytesExcluidos } else { $null }
            MiBExcluidos = if ($null -ne $medicao) { $medicao.MiBExcluidos } else { $null }
            GiBExcluidos = if ($null -ne $medicao) { $medicao.GiBExcluidos } else { $null }
            DocumentosVersoesExcluidos = if ($null -ne $medicao) { $medicao.DocumentosVersoesExcluidos } else { $null }
            StatusMedicao = $statusMedicao; ErroMedicao = $erroMedicao
        })
        $log.ToArray() | Export-Csv -LiteralPath $arquivoLog -NoTypeInformation -Encoding UTF8
    }
    $falhas = @($log.ToArray() | Where-Object Status -eq 'Erro').Count
    Write-Host "Limpeza finalizada. Concluidas: $($log.Count - $falhas). Falhas: $falhas."
    Write-Host "Log: $arquivoLog"
}
finally {
    if ($script:LoginRealizado) {
        try { Undo-PWLogin -ErrorAction Stop | Out-Null }
        catch { Write-Host "Nao foi possivel encerrar a sessao: $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}
