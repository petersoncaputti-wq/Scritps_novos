<#
.SYNOPSIS
    Separa arquivos por SH (SH-01, SH-02, ...) em pastas distintas.

.DESCRIPTION
    Varre a pasta de origem recursivamente. Para cada arquivo, procura no caminho
    (pastas ou nome do arquivo) um identificador de SH: "SH-01", "SH 01", "SH01", "SH_01".
    O arquivo e MOVIDO (ou copiado, com -Copiar) para:

        <Destino>\<nome da pasta origem>\SH-38\<caminho relativo, sem a pasta SH>

    Exemplo (Origem = D:\Lote 21, Destino = D:\Separados):
        D:\Lote 21\GRD-0156-2024-BR493RJ\SH-39\Executivo\b.dwg
        -> D:\Separados\Lote 21\SH-39\GRD-0156-2024-BR493RJ\Executivo\b.dwg

    O SH e identificado SOMENTE pelo nome das pastas (nunca pelo nome do arquivo).

    Arquivos fora de uma pasta SH (pasta "Geral", arquivos soltos na raiz da GRD...):
    o script pergunta UMA VEZ para qual destino enviar (um SH, todos, ou ignorar) e a
    resposta vale para todos eles. Para nao perguntar, use -DestinoSemSH.

    Lixo (.spdownload.part, .part.json, ~$temporarios do Office) NAO e movido:
    fica na origem e aparece no log como ignorado.

    Arquivos com o mesmo nome no destino recebem sufixo _(1), _(2)...
    Um log CSV e gerado na pasta de destino.

.EXAMPLE
    # Sem parametros: abre janelas para escolher origem e destino, depois move
    .\Separar_Arquivos_por_SH.ps1

.EXAMPLE
    # Simulacao (nao move nada, so mostra o que faria)
    .\Separar_Arquivos_por_SH.ps1 -Origem "S:\Ano__Lote_Lote 01_SH__7edd3c95e922\Lote 01" -Destino "S:\Separado_SH" -WhatIf

.EXAMPLE
    # Copia em vez de mover (origem fica intacta)
    .\Separar_Arquivos_por_SH.ps1 -Origem "S:\Ano__Lote_Lote 01_SH__7edd3c95e922\Lote 01" -Destino "S:\Separado_SH" -Copiar

.EXAMPLE
    # Sem perguntas: arquivos sem SH vao para todos os SHs
    .\Separar_Arquivos_por_SH.ps1 -Origem "S:\...\Lote 01" -Destino "S:\Separado_SH" -DestinoSemSH Todos
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Se nao informado, abre uma janela para escolher a pasta
    [string]$Origem,

    # Se nao informado, abre uma janela para escolher a pasta
    [string]$Destino,

    # Copia em vez de mover (o padrao e MOVER)
    [switch]$Copiar,

    # Resposta fixa para arquivos sem SH (nao pergunta):
    # "SH-01", "SH-02", ... | "Todos" | "Ignorar" | "SEM_SH"
    [string]$DestinoSemSH,

    # Mantem a pasta "SH-01" dentro do caminho relativo (por padrao ela e removida)
    [switch]$ManterPastaSH
)

$ErrorActionPreference = 'Stop'
$Mover = -not $Copiar

function Select-Pasta([string]$titulo) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $titulo
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
        return $null
    } catch {
        return (Read-Host $titulo)
    }
}

if (-not $Origem) { $Origem = Select-Pasta 'Selecione a pasta de ORIGEM (onde estao os arquivos misturados)' }
if (-not $Origem) { throw "Nenhuma pasta de origem selecionada." }
if (-not (Test-Path -LiteralPath $Origem -PathType Container)) {
    throw "Pasta de origem nao encontrada: $Origem"
}
$Origem = (Resolve-Path -LiteralPath $Origem).ProviderPath.TrimEnd('\')

if (-not $Destino) { $Destino = Select-Pasta 'Selecione a pasta de DESTINO (onde serao criadas as pastas SH-01, SH-02...)' }
if (-not $Destino) { throw "Nenhuma pasta de destino selecionada." }
$Destino = $Destino.Trim().Trim('"').TrimEnd('\')

# As pastas SH ficam dentro de uma pasta com o nome da origem: <Destino>\Lote 21\SH-38\...
# (se o destino escolhido ja tiver esse nome, nao duplica)
$nomeOrigem = Split-Path $Origem -Leaf
if ((Split-Path $Destino -Leaf) -ne $nomeOrigem) {
    $Destino = Join-Path $Destino $nomeOrigem
}
if ($Destino -ieq $Origem) {
    throw "O destino final ($Destino) e a propria origem. Escolha outra pasta de destino."
}

if ($Destino.StartsWith($Origem + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "O destino nao pode ficar dentro da origem (o script iria reprocessar os proprios arquivos)."
}

# SH seguido de no maximo um separador (espaco, - ou _) e 1-2 digitos.
# A restricao evita falsos positivos como "SH__7edd3c95e922".
$regexSH      = [regex]'(?i)(?<![A-Za-z])SH[\s_-]?(\d{1,2})(?!\d)'
$regexPastaSH = '^(?i)\s*SH[\s_-]?\d{1,2}\s*$'

function Get-CaminhoLivre([string]$caminho) {
    if (-not (Test-Path -LiteralPath $caminho)) { return $caminho }
    $dir  = Split-Path $caminho -Parent
    $base = [IO.Path]::GetFileNameWithoutExtension($caminho)
    $ext  = [IO.Path]::GetExtension($caminho)
    $i = 1
    do {
        $novo = Join-Path $dir ("{0}_({1}){2}" -f $base, $i, $ext)
        $i++
    } while (Test-Path -LiteralPath $novo)
    return $novo
}

$acao = if ($Mover) { 'Mover' } else { 'Copiar' }
Write-Host "Origem : $Origem"
Write-Host "Destino: $Destino"
Write-Host "Acao   : $acao"
Write-Host ""

# ---------------------------------------------------------------------------
# FASE 1 - Classificar cada arquivo
# ---------------------------------------------------------------------------
Write-Host "Analisando arquivos..."
$arquivos = @(Get-ChildItem -LiteralPath $Origem -Recurse -File -Force)
$plano = New-Object System.Collections.Generic.List[object]
$log = New-Object System.Collections.Generic.List[object]

# O log e gravado linha a linha durante a execucao: se o script for interrompido,
# o que ja foi movido continua registrado e pode ser desfeito com Rollback_Separar_Arquivos_por_SH.ps1
$arqLog = $null
if (-not $WhatIfPreference) {
    if (-not (Test-Path -LiteralPath $Destino)) { New-Item -ItemType Directory -Path $Destino -Force | Out-Null }
    $arqLog = Join-Path $Destino ("log_separacao_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
}

function Add-Log([string]$status, [string]$operacao, [string]$sh, [string]$origem, [string]$destino) {
    $linha = [pscustomobject]@{
        DataHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Status   = $status
        Operacao = $operacao
        SH       = $sh
        Origem   = $origem
        Destino  = $destino
    }
    $log.Add($linha)
    if ($script:arqLog) {
        $linha | Export-Csv -LiteralPath $script:arqLog -NoTypeInformation -Delimiter ';' -Encoding UTF8 -Append
    }
}

# Lixo que NAO deve ser movido: downloads incompletos do SharePoint e temporarios do Office
$regexLixo = '(?i)(\.spdownload\.part(\.json)?$|\.part$|\.part\.json$|^~\$)'

foreach ($arq in $arquivos) {
    if ($arq.Name -match $regexLixo) {
        Add-Log 'Ignorado (lixo, mantido na origem)' '' '' $arq.FullName ''
        continue
    }

    $relativo  = $arq.FullName.Substring($Origem.Length + 1)
    $segmentos = $relativo -split '\\'

    # Procura o SH SOMENTE no nome das pastas (o nome do arquivo e ignorado)
    $sh = $null
    $idx = -1
    for ($i = 0; $i -lt $segmentos.Count - 1; $i++) {
        $m = $regexSH.Match($segmentos[$i])
        if ($m.Success) {
            $sh  = 'SH-{0:D2}' -f [int]$m.Groups[1].Value
            $idx = $i
            break
        }
    }

    # Remove do caminho relativo a pasta que e so o SH (ex.: "SH-01"), para nao duplicar
    $partes = [System.Collections.Generic.List[string]]$segmentos
    if (-not $ManterPastaSH -and $idx -ge 0 -and $segmentos[$idx] -match $regexPastaSH) {
        $partes.RemoveAt($idx)
    }

    $plano.Add([pscustomobject]@{
        Arquivo  = $arq
        SH       = $sh
        Relativo = ($partes -join '\')
        Pasta    = $arq.DirectoryName
    })
}

$shsEncontrados = @($plano | Where-Object SH | Select-Object -ExpandProperty SH -Unique | Sort-Object)

# ---------------------------------------------------------------------------
# FASE 2 - Perguntar UMA VEZ o destino dos arquivos sem pasta SH
#          (pasta Geral, arquivos soltos na raiz da GRD, etc.)
# ---------------------------------------------------------------------------
$semSH = @($plano | Where-Object { -not $_.SH })
$destinoSemSHEscolhido = @()   # lista de SHs (vazia = ignorar)

function Convert-Resposta([string]$resp, [string[]]$opcoes) {
    $r = $resp.Trim()
    if ($r -match '^(?i)(t|todos)$')   { return ,$opcoes }
    if ($r -match '^(?i)(i|ignorar)$') { return ,@() }
    if ($r -match '^(?i)SEM_SH$')      { return ,@('SEM_SH') }
    if ($r -match '^\d+$' -and [int]$r -ge 1 -and [int]$r -le $opcoes.Count) { return ,@($opcoes[[int]$r - 1]) }
    $m = $regexSH.Match($r)
    if ($m.Success) { return ,@('SH-{0:D2}' -f [int]$m.Groups[1].Value) }
    return $null
}

if ($semSH.Count -gt 0) {
    $opcoes = if ($shsEncontrados.Count) { $shsEncontrados } else { @('SH-01', 'SH-02') }

    if ($DestinoSemSH) {
        $destinoSemSHEscolhido = Convert-Resposta $DestinoSemSH $opcoes
        if ($null -eq $destinoSemSHEscolhido) { throw "Valor invalido para -DestinoSemSH: $DestinoSemSH" }
    } else {
        # Resumo do que esta sem SH, agrupado pela 1a pasta abaixo da GRD (ex.: "Geral")
        $tamanho = ($semSH | ForEach-Object { $_.Arquivo.Length } | Measure-Object -Sum).Sum
        Write-Host ""
        Write-Host ("{0} arquivo(s) ({1:N1} GB) estao fora de uma pasta SH:" -f $semSH.Count, ($tamanho / 1GB)) -ForegroundColor Yellow
        $semSH | Group-Object {
            $s = $_.Relativo -split '\\'
            if ($s.Count -gt 2) { "pasta '$($s[1])'" } else { 'arquivos soltos na raiz da GRD' }
        } | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
            Write-Host ("   {0,6} - {1}" -f $_.Count, $_.Name)
        }
        Write-Host ""
        Write-Host "Para onde enviar TODOS esses arquivos? (a resposta vale para todos)"
        for ($i = 0; $i -lt $opcoes.Count; $i++) { Write-Host ("   [{0}] {1}" -f ($i + 1), $opcoes[$i]) }
        Write-Host  "   [T] Todos (uma copia em cada SH)"
        Write-Host  "   [I] Ignorar (deixa na origem)"

        $destinoSemSHEscolhido = $null
        while ($null -eq $destinoSemSHEscolhido) {
            $destinoSemSHEscolhido = Convert-Resposta (Read-Host "Opcao") $opcoes
            if ($null -eq $destinoSemSHEscolhido) { Write-Host "Opcao invalida." -ForegroundColor Red }
        }
    }
}

# ---------------------------------------------------------------------------
# FASE 3 - Criar as pastas de destino e mover / copiar
# ---------------------------------------------------------------------------
$pastasSH = @($shsEncontrados) + @($(if ($semSH.Count) { $destinoSemSHEscolhido })) | Where-Object { $_ } | Select-Object -Unique
foreach ($p in $pastasSH) {
    $caminhoSH = Join-Path $Destino $p
    if (-not (Test-Path -LiteralPath $caminhoSH) -and $PSCmdlet.ShouldProcess($caminhoSH, 'Criar pasta')) {
        New-Item -ItemType Directory -Path $caminhoSH -Force | Out-Null
        Write-Host "Pasta criada: $caminhoSH"
    }
}

$contagem = @{}
$total = $plano.Count
$n = 0

foreach ($item in $plano) {
    $n++
    if ($n % 50 -eq 0 -or $n -eq $total) {
        Write-Progress -Activity "Separando por SH" -Status "$n de $total" -PercentComplete (($n / $total) * 100)
    }
    $arq = $item.Arquivo

    [string[]]$alvosSH = @(if ($item.SH) { $item.SH } else { $destinoSemSHEscolhido })
    if ($alvosSH.Count -eq 0) {
        Add-Log 'Ignorado (sem SH, mantido na origem)' '' '' $arq.FullName ''
        continue
    }

    for ($k = 0; $k -lt $alvosSH.Count; $k++) {
        $sh = $alvosSH[$k]
        # Com -Mover e varios destinos: copia para os primeiros, move no ultimo
        $moverAgora = $Mover -and ($k -eq $alvosSH.Count - 1)

        $alvo    = Get-CaminhoLivre (Join-Path (Join-Path $Destino $sh) $item.Relativo)
        $pastaAl = Split-Path $alvo -Parent
        $verbo   = if ($moverAgora) { 'Mover' } else { 'Copiar' }

        if ($PSCmdlet.ShouldProcess($arq.FullName, "$verbo para $alvo")) {
            try {
                if (-not (Test-Path -LiteralPath $pastaAl)) {
                    New-Item -ItemType Directory -Path $pastaAl -Force | Out-Null
                }
                if ($moverAgora) { Move-Item -LiteralPath $arq.FullName -Destination $alvo }
                else             { Copy-Item -LiteralPath $arq.FullName -Destination $alvo }
                $status = 'OK'
            } catch {
                $status = "ERRO: $($_.Exception.Message)"
                Write-Warning "$($arq.FullName) -> $status"
            }
        } else {
            $status = 'Simulado'
        }

        if (-not $item.SH) { $status += ' (sem SH, escolhido)' }
        $contagem[$sh] = 1 + [int]$contagem[$sh]
        Add-Log $status $verbo $sh $arq.FullName $alvo
    }
}

Write-Progress -Activity "Separando por SH" -Completed

# Depois de mover, apaga as pastas que ficaram vazias na origem (a pasta raiz e mantida)
if ($Mover -and -not $WhatIfPreference) {
    Get-ChildItem -LiteralPath $Origem -Recurse -Directory -Force |
        Sort-Object { $_.FullName.Length } -Descending |
        Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force | Select-Object -First 1) } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
}

Write-Host ""
Write-Host "Resumo:" -ForegroundColor Cyan
foreach ($k in ($contagem.Keys | Sort-Object)) {
    Write-Host ("  {0,-8} {1,6} arquivo(s)" -f $k, $contagem[$k])
}
$lixo = @($log | Where-Object { $_.Status -like 'Ignorado (lixo*' }).Count
if ($lixo) { Write-Host "  Lixo nao movido (.part / ~`$): $lixo" }
$ignorados = @($log | Where-Object { $_.Status -like 'Ignorado (sem SH*' }).Count
if ($ignorados) { Write-Host "  Sem SH, mantidos na origem: $ignorados" }
$erros = @($log | Where-Object { $_.Status -like 'ERRO*' }).Count
if ($erros) { Write-Host "  Erros: $erros (veja o log)" -ForegroundColor Red }

if ($arqLog) {
    Write-Host "Log: $arqLog"
    Write-Host "Para desfazer: .\Rollback_Separar_Arquivos_por_SH.ps1 -Log `"$arqLog`"" -ForegroundColor DarkGray
}
