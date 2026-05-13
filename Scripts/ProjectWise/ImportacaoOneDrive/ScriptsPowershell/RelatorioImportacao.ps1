# Função para tratar o tamanho dos arquivos em bytes para GB
function ConvertBytesToGB($bytes) {
    return [math]::round($bytes / 1GB, 2)
}

# Função para ler os CSVs e gerar os relatórios
function GenerateReport {
    # Definir o caminho da pasta onde os CSVs estão armazenados
    $pastaCSVs = Resolve-Path -Path "$PSScriptRoot\..\Mapeamento\" | Select-Object -ExpandProperty "Path"

    # Obter todos os arquivos CSV na pasta
    $csvFiles = Get-ChildItem -Path $pastaCSVs | Where-Object { $_.Name -match '^OD_.*_(Analisado|Importado)\.csv$' }
    $csvFiles = $csvFiles | Sort-Object {
        if ($_.Name -match "_Analisado.csv$") {
            1
        }
        else {
            0
        }
    }

    # Inicializar um hash para armazenar os dados
    $relatorio = @{}
    
    # Lista para armazenar os documentos importados
    $importados = @{}

    # Inicializar hash para armazenar o total de bytes por contrato
    $tamanhosPorContrato = @{}

    # Processar cada arquivo CSV
    foreach ($csvFile in $csvFiles) {
        # Identificar o contrato com base no nome do arquivo
        $contrato = $csvFile.Name -replace "OD_", ''
        $contrato = $contrato -replace "_Analisado.csv", ''
        $contrato = $contrato -replace "_Importado.csv", ''

        # Inicializar estruturas se ainda não existentes
        if (-not $relatorio.ContainsKey($contrato)) {
            $relatorio[$contrato] = @{}
        }
        if (-not $tamanhosPorContrato.ContainsKey($contrato)) {
            $tamanhosPorContrato[$contrato] = @{ "Importados" = 0; "Analisados" = 0 }
        }

        # Importar o conteúdo do CSV
        $dados = Import-Csv -Path $csvFile.FullName -Encoding 'Default'

        # Se o arquivo for do tipo "Importado"
        if ($csvFile.Name -match "_Importado.csv$") {
            $relatorio[$contrato]["Importado"] = $dados.Count

            foreach ($linha in $dados) {
                $importados[$linha.NomeAjustado] = $true

                if ($linha.tamanho_bytes -match '^\d+$') {
                    $tamanhosPorContrato[$contrato]["Importados"] += [int64]$linha.tamanho_bytes
                }
            }
        }
        # Se o arquivo for do tipo "Analisado"
        elseif ($csvFile.Name -match "_Analisado.csv$") {
            foreach ($linha in $dados) {
                # Contabilizar tamanho mesmo se já foi importado
                if ($linha.tamanho_bytes -match '^\d+$') {
                    $tamanhosPorContrato[$contrato]["Analisados"] += [int64]$linha.tamanho_bytes
                }

                # Contar status apenas se não foi importado
                if ($importados.ContainsKey($linha.NomeAjustado)) {
                    continue
                }

                $status = $linha.Status

                if (-not $relatorio[$contrato].ContainsKey($status)) {
                    $relatorio[$contrato][$status] = 0
                }

                $relatorio[$contrato][$status]++
            }
        }
    }

    # Gerar o relatório
    $reportPath = "$PSScriptRoot\relatorio.txt"
    $reportContent = "Resumo Geral:`n"

    foreach ($contrato in $relatorio.Keys) {
        $gbAnalisados = ConvertBytesToGB($tamanhosPorContrato[$contrato]["Analisados"])
        $gbImportados = ConvertBytesToGB($tamanhosPorContrato[$contrato]["Importados"])

        $reportContent += "`n$contrato " + "($gbAnalisados GB Baixados | $gbImportados GB Importados)`n"
        
        foreach ($status in $relatorio[$contrato].Keys) {
            $quantidade = $relatorio[$contrato][$status]
            $reportContent += "$quantidade documentos no status ""$status""`n"
        }

    }

    # Salvar o relatório em arquivo .txt
    Set-Content -Path $reportPath -Value $reportContent -Encoding 'UTF8'
    Write-Host "Relatório gerado com sucesso em: $reportPath"

    # Gerar relatório em CSV
    $csvPath = "$PSScriptRoot\relatorio.csv"

    # Descobrir todos os status únicos (exceto "Importados") corretamente
    $todosStatus = @()
    foreach ($dadosContrato in $relatorio.Values) {
        foreach ($status in $dadosContrato.Keys) {
            if (-not $todosStatus.Contains($status)) {
                $todosStatus += $status
            }
        }
    }

    # Criar cabeçalho do CSV
    $colunas = @("NumeroContrato") + $todosStatus + @("GBs Baixados", "GBs Importados")
    $linhasCSV = @()

    foreach ($contrato in $relatorio.Keys) {
        $linha = @{}
        $linha["NumeroContrato"] = $contrato

        # Preencher os status com valores (ou 0 se não existir)
        foreach ($status in $todosStatus) {
            if ($relatorio[$contrato].ContainsKey($status)) {
                $linha[$status] = $relatorio[$contrato][$status]
            }
            else {
                $linha[$status] = 0
            }
        }

        # Adicionar os valores de tamanho convertidos
        $linha["GBs Baixados"] = ConvertBytesToGB($tamanhosPorContrato[$contrato]["Analisados"])
        $linha["GBs Importados"] = ConvertBytesToGB($tamanhosPorContrato[$contrato]["Importados"])

        $linhasCSV += New-Object PSObject -Property $linha
    }

    # Exportar para CSV
    $linhasCSV | Select-Object $colunas | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation

    Write-Host "Relatório CSV gerado com sucesso em: $csvPath"

    # Obter todos os arquivos CSV na pasta
    $csvFiles = Get-ChildItem -Path $pastaCSVs | Where-Object { $_.Name -match '^OD_.*_(Analisado|Importado)\.csv$' }
    $csvFiles = $csvFiles | Sort-Object {
        if ($_.Name -match "_Analisado.csv$") {
            0
        }
        else {
            1
        }
    }
 
    # Inicializar um hash para armazenar os dados
    $relatorio = @{}
 
    # Processar cada arquivo CSV
    foreach ($csvFile in $csvFiles) {
        $contrato = $csvFile.Name -replace "OD_", ''
        $contrato = $contrato -replace "_Analisado.csv", ''
        $contrato = $contrato -replace "_Importado.csv", ''
 
        # Importar o conteúdo do CSV
        $dados = Import-Csv -Path $csvFile.FullName -Encoding 'Default'
 
        foreach ($linha in $dados) {
            $relatorio["$contrato-$($linha.nome_arquivo)"] = [PScustomObject]@{
                Contrato       = $contrato
                nome_arquivo   = $linha.nome_arquivo
                Status         = $linha.Status.Replace("Importar", "Importado")
                extensao       = $linha.extensao
                tamanho_bytes  = $linha.tamanho_bytes
                full_path      = $linha.full_path
                Observacoes    = $linha.Observacoes
                NomeAjustado   = $linha.NomeAjustado
                DataImportacao = $linha.'Data Importacao'
            }
        }
    }
 
    # Exportar os valores (objetos) do hash $relatorio para um CSV
    $csvLitaPath = "$PSScriptRoot\relatorioListaDocumentos.csv"
    $relatorio.Values | Export-Csv -Path $csvLitaPath -NoTypeInformation -Encoding UTF8
    Write-Host "Relatório total de documentos gerado com sucesso em: $csvLitaPath"
}

# Chamar a função para gerar o relatório
GenerateReport
