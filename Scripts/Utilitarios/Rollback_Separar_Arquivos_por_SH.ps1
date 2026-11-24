<#
.SYNOPSIS
    Desfaz uma execucao do Separar_Arquivos_por_SH.ps1 usando o log CSV gerado por ele.

.DESCRIPTION
    Le o log (log_separacao_*.csv) e, para cada linha com Status "OK":
      - Operacao "Mover" : move o arquivo de volta do Destino para a Origem.
      - Operacao "Copiar": apaga a copia no Destino, MAS SO SE o original existir
                           na Origem com o mesmo tamanho. Se o original nao existir
                           mais, a copia e movida de volta para a Origem (nada se perde).

    Regras de seguranca:
      - Nunca sobrescreve: se ja existir um arquivo na Origem, a linha e pulada e avisada.
      - Arquivos que nao estao no log nao sao tocados.
      - No final, apaga apenas as pastas que ficaram VAZIAS no destino.
      - Gera um log do rollback (log_rollback_*.csv) ao lado do log original.

.EXAMPLE
    # Simulacao (mostra o que faria, nao mexe em nada)
    .\Rollback_Separar_Arquivos_por_SH.ps1 -Log "D:\Lote 21_Separado\log_separacao_20261123_141721.csv" -WhatIf

.EXAMPLE
    # Rollback de verdade (pede confirmacao)
    .\Rollback_Separar_Arquivos_por_SH.ps1 -Log "D:\Lote 21_Separado\log_separacao_20261123_141721.csv"

.EXAMPLE
    # Sem parametro: abre uma janela para escolher o log
    .\Rollback_Separar_Arquivos_por_SH.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Log,

    # Nao pede confirmacao antes de executar
    [switch]$SemConfirmar
)

$ErrorActionPreference = 'Stop'

if (-not $Log) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title  = 'Selecione o log da separacao (log_separacao_*.csv)'
        $dlg.Filter = 'Log de separacao (log_separacao_*.csv)|log_separacao_*.csv|CSV (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Log = $dlg.FileName }
    } catch {
        $Log = Read-Host 'Caminho do log (log_separacao_*.csv)'
    }
}
if (-not $Log) { throw 'Nenhum log selecionado.' }
$Log = $Log.Trim().Trim('"')
if (-not (Test-Path -LiteralPath $Log)) { throw "Log nao encontrado: $Log" }

$linhas = @(Import-Csv -LiteralPath $Log -Delimiter ';' -Encoding UTF8)
if (-not ($linhas | Get-Member -Name Operacao)) {
    throw 'Este log nao tem a coluna "Operacao" (foi gerado por uma versao antiga do script). Rollback automatico nao e seguro.'
}

# Apenas o que realmente foi feito
$feitos  = @($linhas | Where-Object { $_.Status -like 'OK*' -and $_.Destino })
$movidos = @($feitos | Where-Object Operacao -eq 'Mover')
$copias  = @($feitos | Where-Object Operacao -eq 'Copiar')

$pastaDestino = Split-Path $Log -Parent

Write-Host "Log     : $Log"
Write-Host "Destino : $pastaDestino"
Write-Host ("Desfazer: {0} arquivo(s) movidos (voltam para a origem)" -f $movidos.Count)
Write-Host ("          {0} copia(s) (serao apagadas do destino se o original existir)" -f $copias.Count)
Write-Host ""

if ($feitos.Count -eq 0) { Write-Host 'Nada para desfazer.'; return }

if (-not $SemConfirmar -and -not $WhatIfPreference) {
    if ((Read-Host 'Confirmar o rollback? (S/N)') -notmatch '^(?i)s') { Write-Host 'Cancelado.'; return }
}

$arqLogRb = $null
if (-not $WhatIfPreference) {
    $arqLogRb = Join-Path $pastaDestino ("log_rollback_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
}
$resultado = New-Object System.Collections.Generic.List[object]
function Add-LogRb([string]$status, [string]$acao, [string]$de, [string]$para) {
    $l = [pscustomobject]@{ DataHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Status = $status; Acao = $acao; De = $de; Para = $para }
    $resultado.Add($l)
    if ($script:arqLogRb) { $l | Export-Csv -LiteralPath $script:arqLogRb -NoTypeInformation -Delimiter ';' -Encoding UTF8 -Append }
}

function Restore-Arquivo([string]$de, [string]$para) {
    $pasta = Split-Path $para -Parent
    if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
    Move-Item -LiteralPath $de -Destination $para
}

$total = $feitos.Count
$n = 0

# 1) Primeiro devolve os arquivos MOVIDOS (restaura os originais), do ultimo para o primeiro
for ($i = $movidos.Count - 1; $i -ge 0; $i--) {
    $l = $movidos[$i]; $n++
    if ($n % 50 -eq 0) { Write-Progress -Activity 'Rollback' -Status "$n de $total" -PercentComplete (($n / $total) * 100) }

    if (-not (Test-Path -LiteralPath $l.Destino)) {
        Add-LogRb 'AVISO: arquivo nao esta mais no destino' 'Voltar' $l.Destino $l.Origem
        Write-Warning "Nao encontrado no destino: $($l.Destino)"
        continue
    }
    if (Test-Path -LiteralPath $l.Origem) {
        Add-LogRb 'AVISO: ja existe na origem, nao sobrescrito' 'Voltar' $l.Destino $l.Origem
        Write-Warning "Ja existe na origem (pulado): $($l.Origem)"
        continue
    }
    if ($PSCmdlet.ShouldProcess($l.Destino, "Mover de volta para $($l.Origem)")) {
        try   { Restore-Arquivo $l.Destino $l.Origem; Add-LogRb 'OK' 'Voltar' $l.Destino $l.Origem }
        catch { Add-LogRb "ERRO: $($_.Exception.Message)" 'Voltar' $l.Destino $l.Origem; Write-Warning $_.Exception.Message }
    }
}

# 2) Depois trata as COPIAS
foreach ($l in $copias) {
    $n++
    if ($n % 50 -eq 0) { Write-Progress -Activity 'Rollback' -Status "$n de $total" -PercentComplete (($n / $total) * 100) }

    if (-not (Test-Path -LiteralPath $l.Destino)) {
        Add-LogRb 'AVISO: copia nao esta mais no destino' 'Apagar copia' $l.Destino ''
        continue
    }

    if (Test-Path -LiteralPath $l.Origem) {
        $tamOrig = (Get-Item -LiteralPath $l.Origem -Force).Length
        $tamCop  = (Get-Item -LiteralPath $l.Destino -Force).Length
        if ($tamOrig -ne $tamCop) {
            Add-LogRb 'AVISO: tamanho diferente do original, copia mantida' 'Apagar copia' $l.Destino ''
            Write-Warning "Copia diferente do original (mantida): $($l.Destino)"
            continue
        }
        if ($PSCmdlet.ShouldProcess($l.Destino, 'Apagar copia (original existe na origem)')) {
            try   { Remove-Item -LiteralPath $l.Destino -Force; Add-LogRb 'OK' 'Apagar copia' $l.Destino '' }
            catch { Add-LogRb "ERRO: $($_.Exception.Message)" 'Apagar copia' $l.Destino ''; Write-Warning $_.Exception.Message }
        }
    } else {
        # Original sumiu: a copia vira o original (nada se perde)
        if ($PSCmdlet.ShouldProcess($l.Destino, "Mover de volta para $($l.Origem) (original nao existe)")) {
            try   { Restore-Arquivo $l.Destino $l.Origem; Add-LogRb 'OK' 'Voltar (copia)' $l.Destino $l.Origem }
            catch { Add-LogRb "ERRO: $($_.Exception.Message)" 'Voltar (copia)' $l.Destino $l.Origem; Write-Warning $_.Exception.Message }
        }
    }
}
Write-Progress -Activity 'Rollback' -Completed

# 3) Remove pastas que ficaram vazias no destino (logs e a pasta raiz sao mantidos)
if (-not $WhatIfPreference) {
    Get-ChildItem -LiteralPath $pastaDestino -Recurse -Directory -Force |
        Sort-Object { $_.FullName.Length } -Descending |
        Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force | Select-Object -First 1) } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
}

Write-Host ""
Write-Host 'Resumo do rollback:' -ForegroundColor Cyan
$resultado | Group-Object { "$($_.Acao) - $(($_.Status -split ':')[0])" } | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-30} {1,6}" -f $_.Name, $_.Count)
}
$problemas = @($resultado | Where-Object { $_.Status -ne 'OK' }).Count
if ($problemas) { Write-Host "  $problemas item(ns) com aviso/erro - veja o log do rollback" -ForegroundColor Yellow }
if ($arqLogRb)  { Write-Host "Log do rollback: $arqLogRb" }
