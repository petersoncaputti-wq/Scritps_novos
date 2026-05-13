Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

function JanelaSeletorExcel {
    param ()

    $janela = New-Object System.Windows.Forms.OpenFileDialog

    $janela.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $janela.Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
    $janela.Title = "Selecione Excel com os dados dos projetos a serem criados"

    return $janela
}
function SelecionarPlanilhaCadastrosDeProjetos {
    param ()
    
    $janela = JanelaSeletorExcel

    if ($janela.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $enderecoArquivoExcel = $janela.FileName     
        $dados = Import-Excel -Path $enderecoArquivoExcel

        return $dados
    }
    return $null
}
function mapearArquivosProjeto {
    param ($dadosProjeto)

    $nomeProjeto = $dadosProjeto."Nome Projeto"
    $pastaRaiz = $dadosProjeto."Link Rede"
    
    if (-Not (Test-Path -Path $pastaRaiz)) {
        Write-Error "O caminho '$pastaRaiz' não existe."
        return $null
    }
    $arquivos = Get-ChildItem -Path $pastaRaiz -Recurse -File -ErrorAction SilentlyContinue
    $dados = $arquivos | ForEach-Object {
        [PSCustomObject]@{
            nome_arquivo = $_.Name
            extensao     = $_.Extension
            tamanho_bytes = $_.Length
            path         = $_.DirectoryName
            full_path    = $_.FullName
        }
    }

    return $dados
}

$dados = SelecionarPlanilhaCadastrosDeProjetos

foreach ($projeto in $dados) {
    $nomeProjeto = $projeto."Nome Projeto"
    $nomeProjeto = $nomeProjeto.Replace("*","(Asterisco)")
    $arquivosMapeados = mapearArquivosProjeto -dadosProjeto $projeto
    $pathExportacao = Resolve-Path -Path "$PSScriptRoot\..\Mapeamento\" | Select-Object -ExpandProperty "Path"
    $arquivosMapeados | Export-Csv -Path "$pathExportacao\PR_$nomeProjeto.csv" -Encoding UTF8 -Delimiter ";" -NoTypeInformation
    Write-Host "Relatório gerado em: $pathExportacao\PR_$nomeProjeto.csv"
}



