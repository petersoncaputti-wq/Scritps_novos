Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

$NumeroThreads = 32
$TentativasErro = 2
$EsperaTentativaSegundos = 2

# ==========================================================
# FUNÇÕES
# ==========================================================

function Selecionar-Pasta {
    param (
        [Parameter(Mandatory)]
        [string]$Titulo,

        [bool]$PermitirCriarPasta = $true
    )

    $janela = New-Object System.Windows.Forms.FolderBrowserDialog
    $janela.Description = $Titulo
    $janela.ShowNewFolderButton = $PermitirCriarPasta

    $resultado = $janela.ShowDialog()

    if ($resultado -eq [System.Windows.Forms.DialogResult]::OK) {
        return $janela.SelectedPath
    }

    return $null
}

function Converter-Tamanho {
    param (
        [long]$Bytes
    )

    if ($Bytes -ge 1TB) {
        return "{0:N2} TB" -f ($Bytes / 1TB)
    }

    if ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }

    if ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }

    if ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }

    return "$Bytes bytes"
}

function Obter-EstatisticasPasta {
    param (
        [Parameter(Mandatory)]
        [string]$Caminho
    )

    $quantidadeArquivos = 0L
    $quantidadePastas = 0L
    $tamanhoTotal = 0L
    $erros = [System.Collections.Generic.List[string]]::new()

    Write-Host ""
    Write-Host "Analisando os arquivos da origem..." -ForegroundColor Cyan

    Write-Progress `
        -Activity "Analisando a pasta de origem" `
        -Status "Localizando arquivos e calculando o tamanho..." `
        -PercentComplete -1

    try {
        Get-ChildItem `
            -LiteralPath $Caminho `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue `
            -ErrorVariable +errosEnumeracao |
        ForEach-Object {
            if ($_.PSIsContainer) {
                $quantidadePastas++
            }
            else {
                $quantidadeArquivos++
                $tamanhoTotal += $_.Length
            }

            if (($quantidadeArquivos + $quantidadePastas) % 10000 -eq 0) {
                Write-Progress `
                    -Activity "Analisando a pasta de origem" `
                    -Status "$quantidadeArquivos arquivos encontrados..." `
                    -PercentComplete -1
            }
        }

        if ($errosEnumeracao) {
            foreach ($erro in $errosEnumeracao) {
                $erros.Add($erro.Exception.Message)
            }
        }
    }
    finally {
        Write-Progress `
            -Activity "Analisando a pasta de origem" `
            -Completed
    }

    return [PSCustomObject]@{
        Arquivos     = $quantidadeArquivos
        Pastas       = $quantidadePastas
        TamanhoBytes = $tamanhoTotal
        Erros        = $erros
    }
}

function Testar-Leitura {
    param (
        [Parameter(Mandatory)]
        [string]$Caminho
    )

    try {
        Get-ChildItem `
            -LiteralPath $Caminho `
            -Force `
            -ErrorAction Stop |
        Select-Object -First 1 |
        Out-Null

        return $true
    }
    catch {
        return $false
    }
}

function Testar-Gravacao {
    param (
        [Parameter(Mandatory)]
        [string]$Caminho
    )

    $arquivoTeste = Join-Path `
        -Path $Caminho `
        -ChildPath ".teste_gravacao_$(New-Guid).tmp"

    try {
        [System.IO.File]::WriteAllText($arquivoTeste, "Teste de gravação")
        Remove-Item -LiteralPath $arquivoTeste -Force -ErrorAction SilentlyContinue

        return $true
    }
    catch {
        return $false
    }
}

function Obter-EspacoLivre {
    param (
        [Parameter(Mandatory)]
        [string]$Caminho
    )

    try {
        $raiz = [System.IO.Path]::GetPathRoot($Caminho)

        if ([string]::IsNullOrWhiteSpace($raiz)) {
            return $null
        }

        $drive = New-Object System.IO.DriveInfo($raiz)

        return [PSCustomObject]@{
            Raiz        = $raiz
            EspacoLivre = $drive.AvailableFreeSpace
            Tamanho     = $drive.TotalSize
        }
    }
    catch {
        # Em alguns compartilhamentos de rede o DriveInfo
        # pode não conseguir consultar o espaço disponível.
        return $null
    }
}

function Confirmar-Operacao {
    param (
        [Parameter(Mandatory)]
        [string]$Mensagem
    )

    while ($true) {
        $resposta = Read-Host "$Mensagem (S/N)"

        switch -Regex ($resposta.Trim()) {
            '^[Ss]$' {
                return $true
            }

            '^[Nn]$' {
                return $false
            }

            default {
                Write-Host "Resposta inválida. Digite S ou N." -ForegroundColor Yellow
            }
        }
    }
}

function Executar-RobocopyComProgresso {
    param (
        [Parameter(Mandatory)]
        [string]$Origem,

        [Parameter(Mandatory)]
        [string]$Destino,

        [Parameter(Mandatory)]
        [string]$ArquivoLog
    )

    $argumentos = @(
        $Origem
        $Destino
        "/E"
        "/Z"
        "/MT:$NumeroThreads"
        "/R:$TentativasErro"
        "/W:$EsperaTentativaSegundos"
        "/COPY:DAT"
        "/DCOPY:DAT"
        "/ETA"
        "/BYTES"
        "/XJ"
        "/TEE"
        "/LOG:$ArquivoLog"
    )

    $arquivoAtual = "Preparando a cópia..."
    $ultimoPercentual = 0

    Write-Host ""
    Write-Host "Cópia iniciada..." -ForegroundColor Cyan
    Write-Host "Não feche esta janela durante a execução." -ForegroundColor Yellow
    Write-Host ""

    & robocopy @argumentos 2>&1 |
    ForEach-Object {
        $linha = $_.ToString()

        # Mantém a saída original do Robocopy na tela.
        Write-Host $linha

        # Tenta identificar o nome ou caminho do arquivo atual.
        if (
            $linha -match '([A-Za-z]:\\.+)' -or
            $linha -match '(\\\\[^\\]+\\.+)'
        ) {
            $possivelCaminho = $Matches[1].Trim()

            if (-not [string]::IsNullOrWhiteSpace($possivelCaminho)) {
                $arquivoAtual = $possivelCaminho
            }
        }

        # Captura percentuais exibidos pelo Robocopy.
        if ($linha -match '(\d{1,3}(?:[.,]\d+)?)\s*%') {
            $textoPercentual = $Matches[1].Replace(",", ".")

            $percentual = 0.0

            if (
                [double]::TryParse(
                    $textoPercentual,
                    [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$percentual
                )
            ) {
                $percentual = [Math]::Min(
                    100,
                    [Math]::Max(0, $percentual)
                )

                $ultimoPercentual = [int]$percentual

                Write-Progress `
                    -Activity "Copiando arquivos com Robocopy" `
                    -Status "$ultimoPercentual% — $arquivoAtual" `
                    -CurrentOperation $arquivoAtual `
                    -PercentComplete $ultimoPercentual
            }
        }
        else {
            Write-Progress `
                -Activity "Copiando arquivos com Robocopy" `
                -Status $arquivoAtual `
                -CurrentOperation "Processando..." `
                -PercentComplete $ultimoPercentual
        }
    }

    $codigoSaida = $LASTEXITCODE

    Write-Progress `
        -Activity "Copiando arquivos com Robocopy" `
        -Completed

    return $codigoSaida
}

# ==========================================================
# INÍCIO DO SCRIPT
# ==========================================================

Clear-Host

Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "       ASSISTENTE DE CÓPIA COM ROBOCOPY       " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# ==========================================================
# 1. SELEÇÃO DA ORIGEM
# ==========================================================

$origem = Selecionar-Pasta `
    -Titulo "Selecione a pasta de ORIGEM" `
    -PermitirCriarPasta $false

if ([string]::IsNullOrWhiteSpace($origem)) {
    Write-Host "Operação cancelada na seleção da origem." -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit
}

# ==========================================================
# 2. SELEÇÃO DO DESTINO
# ==========================================================

$destino = Selecionar-Pasta `
    -Titulo "Selecione a pasta de DESTINO" `
    -PermitirCriarPasta $true

if ([string]::IsNullOrWhiteSpace($destino)) {
    Write-Host "Operação cancelada na seleção do destino." -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit
}

$origemNormalizada = [System.IO.Path]::GetFullPath($origem).TrimEnd('\')
$destinoNormalizado = [System.IO.Path]::GetFullPath($destino).TrimEnd('\')

# ==========================================================
# 3. VALIDAÇÕES BÁSICAS
# ==========================================================

Write-Host ""
Write-Host "Executando validações..." -ForegroundColor Cyan
Write-Host ""

$validacaoFalhou = $false

if (Test-Path -LiteralPath $origem -PathType Container) {
    Write-Host "[OK] A pasta de origem existe." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] A pasta de origem não existe." -ForegroundColor Red
    $validacaoFalhou = $true
}

if (Test-Path -LiteralPath $destino -PathType Container) {
    Write-Host "[OK] A pasta de destino existe." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] A pasta de destino não existe." -ForegroundColor Red
    $validacaoFalhou = $true
}

if ($origemNormalizada -eq $destinoNormalizado) {
    Write-Host "[ERRO] A origem e o destino são a mesma pasta." -ForegroundColor Red
    $validacaoFalhou = $true
}
else {
    Write-Host "[OK] A origem e o destino são diferentes." -ForegroundColor Green
}

if (
    $destinoNormalizado.StartsWith(
        "$origemNormalizada\",
        [System.StringComparison]::OrdinalIgnoreCase
    )
) {
    Write-Host "[ERRO] O destino está dentro da pasta de origem." -ForegroundColor Red
    Write-Host "       Isso poderia provocar uma cópia recursiva." -ForegroundColor Red
    $validacaoFalhou = $true
}
else {
    Write-Host "[OK] O destino não está dentro da origem." -ForegroundColor Green
}

if (Testar-Leitura -Caminho $origem) {
    Write-Host "[OK] Permissão de leitura na origem." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Não foi possível ler a pasta de origem." -ForegroundColor Red
    $validacaoFalhou = $true
}

if (Testar-Gravacao -Caminho $destino) {
    Write-Host "[OK] Permissão de gravação no destino." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Não foi possível gravar na pasta de destino." -ForegroundColor Red
    $validacaoFalhou = $true
}

if ($validacaoFalhou) {
    Write-Host ""
    Write-Host "A operação não pode continuar devido aos erros acima." -ForegroundColor Red
    Read-Host "Pressione Enter para fechar"
    exit 1
}

# ==========================================================
# 4. ANÁLISE DA ORIGEM
# ==========================================================

$estatisticas = Obter-EstatisticasPasta -Caminho $origem
$espacoDestino = Obter-EspacoLivre -Caminho $destino

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "             RESUMO DA ORIGEM                 " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Origem..............: $origem"
Write-Host "Destino.............: $destino"
Write-Host ""
Write-Host ("Arquivos............: {0:N0}" -f $estatisticas.Arquivos)
Write-Host ("Pastas..............: {0:N0}" -f $estatisticas.Pastas)
Write-Host "Tamanho total.......: $(Converter-Tamanho $estatisticas.TamanhoBytes)"

if ($estatisticas.Erros.Count -gt 0) {
    Write-Host ""
    Write-Host "Aviso: alguns itens não puderam ser analisados." -ForegroundColor Yellow
    Write-Host "Quantidade de erros.: $($estatisticas.Erros.Count)" -ForegroundColor Yellow
}

if ($null -ne $espacoDestino) {
    Write-Host ""
    Write-Host "Unidade de destino..: $($espacoDestino.Raiz)"
    Write-Host "Espaço livre........: $(Converter-Tamanho $espacoDestino.EspacoLivre)"

    $espacoRestante = $espacoDestino.EspacoLivre - $estatisticas.TamanhoBytes

    if ($espacoDestino.EspacoLivre -ge $estatisticas.TamanhoBytes) {
        Write-Host "Espaço suficiente...: SIM" -ForegroundColor Green
        Write-Host "Livre após a cópia..: $(Converter-Tamanho $espacoRestante)"
    }
    else {
        Write-Host "Espaço suficiente...: NÃO" -ForegroundColor Red
        Write-Host ""
        Write-Host "Faltam aproximadamente:" -ForegroundColor Red
        Write-Host "$(Converter-Tamanho ([Math]::Abs($espacoRestante)))" -ForegroundColor Red
        Write-Host ""
        Read-Host "Pressione Enter para fechar"
        exit 1
    }
}
else {
    Write-Host ""
    Write-Host "Não foi possível consultar o espaço livre do destino." -ForegroundColor Yellow
    Write-Host "Isso pode acontecer em alguns compartilhamentos de rede." -ForegroundColor Yellow
}

if ($estatisticas.Arquivos -eq 0) {
    Write-Host ""
    Write-Host "Nenhum arquivo foi encontrado na origem." -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit
}

# ==========================================================
# 5. SIMULAÇÃO DO ROBOCOPY
# ==========================================================

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "            SIMULAÇÃO DO ROBOCOPY              " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "O Robocopy verificará o que seria copiado." -ForegroundColor Cyan
Write-Host "Nenhum arquivo será alterado nesta etapa." -ForegroundColor Yellow
Write-Host ""

$dataHora = Get-Date -Format "yyyyMMdd_HHmmss"

$logSimulacao = Join-Path `
    -Path $destino `
    -ChildPath "Robocopy_Simulacao_$dataHora.log"

$argumentosSimulacao = @(
    $origem
    $destino
    "/E"
    "/L"
    "/R:0"
    "/W:0"
    "/COPY:DAT"
    "/DCOPY:DAT"
    "/BYTES"
    "/XJ"
    "/TEE"
    "/LOG:$logSimulacao"
)

Write-Progress `
    -Activity "Simulação do Robocopy" `
    -Status "Analisando diferenças entre origem e destino..." `
    -PercentComplete -1

& robocopy @argumentosSimulacao

$codigoSimulacao = $LASTEXITCODE

Write-Progress `
    -Activity "Simulação do Robocopy" `
    -Completed

Write-Host ""
Write-Host "Log da simulação:" -ForegroundColor Cyan
Write-Host $logSimulacao

if ($codigoSimulacao -ge 8) {
    Write-Host ""
    Write-Host "A simulação encontrou um erro grave." -ForegroundColor Red
    Write-Host "Código de saída do Robocopy: $codigoSimulacao" -ForegroundColor Red
    Write-Host "Consulte o log antes de continuar." -ForegroundColor Red
    Read-Host "Pressione Enter para fechar"
    exit $codigoSimulacao
}

# ==========================================================
# 6. CONFIRMAÇÃO FINAL
# ==========================================================

Write-Host ""
Write-Host "==============================================" -ForegroundColor Yellow
Write-Host "              CONFIRMAÇÃO FINAL               " -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Origem : $origem"
Write-Host "Destino: $destino"
Write-Host ""
Write-Host "Arquivos analisados: $($estatisticas.Arquivos)"
Write-Host "Volume analisado...: $(Converter-Tamanho $estatisticas.TamanhoBytes)"
Write-Host ""

if (-not (Confirmar-Operacao -Mensagem "Deseja iniciar a cópia real agora?")) {
    Write-Host ""
    Write-Host "Cópia cancelada pelo usuário." -ForegroundColor Yellow
    Read-Host "Pressione Enter para fechar"
    exit
}

# ==========================================================
# 7. EXECUÇÃO REAL
# ==========================================================

$logCopia = Join-Path `
    -Path $destino `
    -ChildPath "Robocopy_Copia_$dataHora.log"

$inicio = Get-Date

$codigoRobocopy = Executar-RobocopyComProgresso `
    -Origem $origem `
    -Destino $destino `
    -ArquivoLog $logCopia

$fim = Get-Date
$duracao = $fim - $inicio

# ==========================================================
# 8. RESULTADO
# ==========================================================

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "             RESULTADO DA CÓPIA               " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Código do Robocopy.: $codigoRobocopy"
Write-Host "Tempo total........: $($duracao.ToString('dd\.hh\:mm\:ss'))"
Write-Host "Log da execução....: $logCopia"
Write-Host ""

# Códigos de 0 a 7 não representam falha grave.
if ($codigoRobocopy -lt 8) {
    Write-Host "Processo concluído sem falhas graves." -ForegroundColor Green

    switch ($codigoRobocopy) {
        0 {
            Write-Host "Nenhum arquivo precisou ser copiado." -ForegroundColor Green
        }

        1 {
            Write-Host "Arquivos copiados com sucesso." -ForegroundColor Green
        }

        default {
            Write-Host "A cópia terminou com avisos ou diferenças." -ForegroundColor Yellow
            Write-Host "Consulte o resumo no final do log." -ForegroundColor Yellow
        }
    }
}
else {
    Write-Host "O Robocopy encontrou uma falha grave." -ForegroundColor Red
    Write-Host "Consulte o arquivo de log para identificar o problema." -ForegroundColor Red
}

Write-Host ""

if (Confirmar-Operacao -Mensagem "Deseja abrir a pasta de destino?") {
    Start-Process explorer.exe -ArgumentList "`"$destino`""
}

Write-Host ""
Read-Host "Pressione Enter para fechar"