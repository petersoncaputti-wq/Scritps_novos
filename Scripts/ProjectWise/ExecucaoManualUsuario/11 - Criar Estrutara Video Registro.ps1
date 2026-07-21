# Carrega o componente necessário para exibir a janela de seleção
Add-Type -AssemblyName System.Windows.Forms

# Cria a janela de seleção de pasta
$seletorPasta = New-Object System.Windows.Forms.FolderBrowserDialog

$seletorPasta.Description = "Selecione a pasta onde a estrutura de coleta será criada."
$seletorPasta.ShowNewFolderButton = $true

# Exibe a janela
$resultado = $seletorPasta.ShowDialog()

# Encerra caso o usuário cancele
if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Operação cancelada. Nenhuma pasta foi criada." -ForegroundColor Yellow
    exit
}

$pastaDestino = $seletorPasta.SelectedPath

# Solicita confirmação
Write-Host ""
Write-Host "A estrutura será criada em:" -ForegroundColor Cyan
Write-Host $pastaDestino -ForegroundColor White
Write-Host ""

$confirmacao = Read-Host "Deseja continuar? Digite S para confirmar"

if ($confirmacao -notmatch "^[Ss]$") {
    Write-Host "Operação cancelada. Nenhuma pasta foi criada." -ForegroundColor Yellow
    exit
}

# Define as pastas principais
$pastasPrincipais = @(
    "COLETA_DIURNA",
    "COLETA_NOTURNA"
)

# Define as subpastas
$subpastas = @(
    "Dispositivos",
    "Eixo",
    "Faixa Adicional",
    "Pedagio",
    "SPA",
    "Vias"
)

try {
    foreach ($pastaPrincipal in $pastasPrincipais) {

        $caminhoPrincipal = Join-Path `
            -Path $pastaDestino `
            -ChildPath $pastaPrincipal

        # Cria a pasta principal
        New-Item `
            -Path $caminhoPrincipal `
            -ItemType Directory `
            -Force `
            -ErrorAction Stop | Out-Null

        # Cria as subpastas
        foreach ($subpasta in $subpastas) {

            $caminhoSubpasta = Join-Path `
                -Path $caminhoPrincipal `
                -ChildPath $subpasta

            New-Item `
                -Path $caminhoSubpasta `
                -ItemType Directory `
                -Force `
                -ErrorAction Stop | Out-Null
        }
    }

    Write-Host ""
    Write-Host "Estrutura criada com sucesso!" -ForegroundColor Green
    Write-Host "Local: $pastaDestino" -ForegroundColor Green

    # Abre a pasta de destino no Explorador de Arquivos
    Start-Process explorer.exe -ArgumentList "`"$pastaDestino`""
}
catch {
    Write-Host ""
    Write-Host "Ocorreu um erro ao criar a estrutura:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}