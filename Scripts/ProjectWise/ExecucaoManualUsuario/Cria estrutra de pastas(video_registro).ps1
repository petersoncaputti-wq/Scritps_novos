# Cria uma nova pasta principal e replica somente as subpastas da origem.
# Nenhum arquivo é copiado.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

function Selecionar-Pasta($titulo) {
    $janela = New-Object System.Windows.Forms.FolderBrowserDialog
    $janela.Description = $titulo
    $janela.ShowNewFolderButton = $false

    try {
        if ($janela.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $janela.SelectedPath
        }
        return $null
    }
    finally {
        $janela.Dispose()
    }
}

try {
Write-Host 'Selecione a pasta-modelo. Somente as pastas serao replicadas.' -ForegroundColor Cyan
$origem = Selecionar-Pasta "Selecione a pasta-modelo"
if (-not $origem) { Write-Host 'Operacao cancelada.'; return }
Write-Host "Origem: $origem"

Write-Host 'Selecione a pasta de destino.' -ForegroundColor Cyan
$pastaPaiDestino = Selecionar-Pasta "Selecione onde a nova pasta sera criada"
if (-not $pastaPaiDestino) { Write-Host 'Operacao cancelada.'; return }
Write-Host "Pasta de destino: $pastaPaiDestino"

$nomeNovaPasta = (Read-Host 'Digite aqui no terminal o nome da nova pasta principal (Enter para cancelar)').Trim()
if ([string]::IsNullOrWhiteSpace($nomeNovaPasta)) { Write-Host 'Operacao cancelada.'; return }
if ($nomeNovaPasta -in @('.', '..') -or
    $nomeNovaPasta.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
    $nomeNovaPasta.EndsWith('.') -or
    $nomeNovaPasta -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)') {
    throw 'Nome de pasta invalido. Informe somente um nome, sem caminho ou caracteres especiais.'
}

$destino = Join-Path $pastaPaiDestino $nomeNovaPasta

if (-not (Test-Path -LiteralPath $origem -PathType Container)) {
    Write-Host "A pasta de origem nao foi encontrada." -ForegroundColor Red
    return
}

if (Test-Path -LiteralPath $destino) {
    Write-Host "O destino ja existe. Escolha outro nome ou local." -ForegroundColor Yellow
    return
}

# Impede que a enumeracao recursiva encontre as proprias pastas criadas.
$prefixoOrigem = $origem.TrimEnd('\') + '\'
if ($destino.StartsWith($prefixoOrigem, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Escolha um destino fora da pasta-modelo para evitar uma copia recursiva da propria estrutura.'
}

Write-Host "Criando estrutura em: $destino" -ForegroundColor Cyan
New-Item -ItemType Directory -Path $destino -Force | Out-Null

$quantidade = 0
Write-Host 'Buscando subpastas na origem... Em locais de rede esta etapa pode demorar.' -ForegroundColor Cyan
Get-ChildItem -LiteralPath $origem -Directory -Recurse | ForEach-Object {
    $caminhoRelativo = $_.FullName.Substring($origem.Length).TrimStart('\')
    $novaPasta = Join-Path $destino $caminhoRelativo
    New-Item -ItemType Directory -Path $novaPasta -Force | Out-Null
    $quantidade++
    Write-Progress -Activity 'Replicando estrutura de pastas' -Status "$quantidade subpasta(s) criada(s): $caminhoRelativo"
}

Write-Host "Estrutura criada com sucesso: $quantidade subpasta(s) em $destino" -ForegroundColor Green
}
catch {
    Write-Host "Nao foi possivel concluir a estrutura: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Se alguma pasta ja foi criada, a estrutura pode estar incompleta.' -ForegroundColor Yellow
    throw
}
finally {
    Write-Progress -Activity 'Replicando estrutura de pastas' -Completed
}
