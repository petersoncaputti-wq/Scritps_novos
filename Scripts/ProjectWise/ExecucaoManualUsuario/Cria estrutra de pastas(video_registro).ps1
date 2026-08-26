# Cria uma nova pasta principal e replica somente as subpastas da origem.
# Nenhum arquivo é copiado.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

function Selecionar-Pasta($titulo) {
    $janela = New-Object System.Windows.Forms.FolderBrowserDialog
    $janela.Description = $titulo
    $janela.ShowNewFolderButton = $false

    if ($janela.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $janela.SelectedPath
    }

    return $null
}

$origem = Selecionar-Pasta "Selecione a pasta-modelo"
if (-not $origem) { exit }

$pastaPaiDestino = Selecionar-Pasta "Selecione onde a nova pasta sera criada"
if (-not $pastaPaiDestino) { exit }

$nomeNovaPasta = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Informe o nome da nova pasta principal.",
    "Nome da nova pasta"
)
if ([string]::IsNullOrWhiteSpace($nomeNovaPasta)) { exit }

$destino = Join-Path $pastaPaiDestino $nomeNovaPasta

if (-not (Test-Path -LiteralPath $origem -PathType Container)) {
    Write-Host "A pasta de origem nao foi encontrada." -ForegroundColor Red
    exit 1
}

if (Test-Path -LiteralPath $destino) {
    Write-Host "O destino ja existe. Escolha outro nome ou local." -ForegroundColor Yellow
    exit 1
}

New-Item -ItemType Directory -Path $destino -Force | Out-Null

Get-ChildItem -LiteralPath $origem -Directory -Recurse | ForEach-Object {
    $caminhoRelativo = $_.FullName.Substring($origem.Length).TrimStart('\')
    $novaPasta = Join-Path $destino $caminhoRelativo
    New-Item -ItemType Directory -Path $novaPasta -Force | Out-Null
}

Write-Host "Estrutura criada com sucesso em: $destino" -ForegroundColor Green
