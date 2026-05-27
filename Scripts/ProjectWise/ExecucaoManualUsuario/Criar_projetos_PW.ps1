param(
    $noGUI,
    $nomeConcessao,
    $siglaConcessao,
    $projeto,
    $descricao,
    $projetista,
    $gestorEngenharia,
    $assistenteEngenharia,
    $poderConcedente,
    $th
)

Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

$my_other_script = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\Funcoes\CriarProjeto.ps1"))

if (-not (Test-Path $my_other_script)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Não foi possível localizar o arquivo CriarProjeto.ps1.`r`nCaminho esperado:`r`n$my_other_script",
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

Write-Host 'Carregando funcoes de criacao de projeto: ' $my_other_script
. $my_other_script

#-------------------------------------------------------
# Funções auxiliares
#-------------------------------------------------------
function Normalizar-Texto {
    param([object]$Valor)

    if ($null -eq $Valor) {
        return ''
    }

    return ([string]$Valor).Trim()
}

function Escrever-LogProjeto {
    param(
        [string]$NomeConcessao,
        [string]$SiglaConcessao,
        [string]$Projeto,
        [string]$Descricao,
        [string]$Projetista,
        [string]$GestorEngenharia,
        [string]$AssistenteEngenharia,
        [string]$PoderConcedente,
        [string]$TH
    )

    Write-Host "--------------------------------------------------"
    Write-Host "Nome Concessao       : [$NomeConcessao]"
    Write-Host "Sigla Concessao      : [$SiglaConcessao]"
    Write-Host "Nome Projeto         : [$Projeto]"
    Write-Host "Descricao            : [$Descricao]"
    Write-Host "Projetista           : [$Projetista]"
    Write-Host "Gestor Engenharia    : [$GestorEngenharia]"
    Write-Host "Assistente Engenharia: [$AssistenteEngenharia]"
    Write-Host "Poder Concedente     : [$PoderConcedente]"
    Write-Host "TH                   : [$TH]"
    Write-Host "--------------------------------------------------"
}

#-------------------------------------------------------
# Form para selecionar o Excel com os dados de projeto
#-------------------------------------------------------
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

        try {
            $dados = Import-Excel -Path $enderecoArquivoExcel
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Erro ao ler a planilha Excel.`r`n$($_.Exception.Message)",
                "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        if (-not $dados -or $dados.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "A planilha selecionada não possui dados para processamento.",
                "Aviso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        CriarProjetos -DadosProjetos $dados

        [System.Windows.Forms.MessageBox]::Show(
            "Criação de pastas de projeto finalizada!",
            "Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}

#-------------------------------------------------------
# Função que consolida todas as tarefas de criação do projeto
#-------------------------------------------------------
function CriarProjetos {
    param ($dadosProjetos)

    foreach ($dadosProjeto in $dadosProjetos) {

        $nomeConcessao        = Normalizar-Texto $dadosProjeto.'Nome Concessao'
        $siglaConcessao       = Normalizar-Texto $dadosProjeto.'Sigla Concessao'
        $projeto              = Normalizar-Texto $dadosProjeto.'Nome Projeto'
        $descricao            = Normalizar-Texto $dadosProjeto.'Descricao'
        $projetista           = Normalizar-Texto $dadosProjeto.'Projetista'
        $gestorEngenharia     = Normalizar-Texto $dadosProjeto.'Gestor Engenharia'
        $assistenteEngenharia = Normalizar-Texto $dadosProjeto.'Assistente Engenharia'
        $poderConcedente      = Normalizar-Texto $dadosProjeto.'Poder Concedente'
        $th                   = Normalizar-Texto $dadosProjeto.'TH'

        Escrever-LogProjeto `
            -NomeConcessao $nomeConcessao `
            -SiglaConcessao $siglaConcessao `
            -Projeto $projeto `
            -Descricao $descricao `
            -Projetista $projetista `
            -GestorEngenharia $gestorEngenharia `
            -AssistenteEngenharia $assistenteEngenharia `
            -PoderConcedente $poderConcedente `
            -TH $th

        if ([string]::IsNullOrWhiteSpace($nomeConcessao) -or
            [string]::IsNullOrWhiteSpace($siglaConcessao) -or
            [string]::IsNullOrWhiteSpace($projeto) -or
            [string]::IsNullOrWhiteSpace($descricao) -or
            [string]::IsNullOrWhiteSpace($projetista) -or
            [string]::IsNullOrWhiteSpace($gestorEngenharia) -or
            [string]::IsNullOrWhiteSpace($assistenteEngenharia) -or
            [string]::IsNullOrWhiteSpace($poderConcedente)) {

            $mensagemCamposObrigatorios = @"
Há campos obrigatórios vazios na linha atual da planilha.

Projeto: [$projeto]
Concessão: [$nomeConcessao]

Verifique os campos:
- Nome Concessao
- Sigla Concessao
- Nome Projeto
- Descricao
- Projetista
- Gestor Engenharia
- Assistente Engenharia
- Poder Concedente
"@

            Write-Host $mensagemCamposObrigatorios

            if ($noGUI -eq 'true') {
                continue
            }

            $retorno = [System.Windows.Forms.MessageBox]::Show(
                $mensagemCamposObrigatorios + "`r`nClique em OK para ignorar este projeto e continuar.`r`nClique em CANCELAR para encerrar.",
                "Campos obrigatórios não preenchidos",
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($retorno -eq 'OK') {
                continue
            }
            else {
                return
            }
        }

        if ($projeto.Length -gt 28) {
            $mensagemErroNomeProjetoGrande = "O nome do Projeto não pode ter mais que 28 caracteres!`r`n`r`nClique em OK para ignorar este projeto e continuar criando os demais.`r`n`r`nClique em CANCELAR para encerrar a execução do script.`r`n`r`nProjeto: ""$projeto"" possui $($projeto.Length) caracteres."
            Write-Host $mensagemErroNomeProjetoGrande

            if ($noGUI -eq 'true') {
                continue
            }

            $retorno = [System.Windows.Forms.MessageBox]::Show(
                $mensagemErroNomeProjetoGrande,
                "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            if ($retorno -eq 'OK') {
                continue
            }
            else {
                return
            }
        }

        try {
            $resultado = CriarProjeto `
                -PoderConcedente $poderConcedente `
                -NomeConcessao $nomeConcessao `
                -SiglaConcessao $siglaConcessao `
                -Projeto $projeto `
                -Descricao $descricao `
                -Projetista $projetista `
                -GestorEngenharia $gestorEngenharia `
                -AssistenteEngenharia $assistenteEngenharia
        }
        catch {
            Write-Host "Erro ao criar projeto [$projeto]: $($_.Exception.Message)"
            $resultado = $null
        }

        if ($resultado) {
            Write-Host "Projeto [$projeto] da concessão [$nomeConcessao] criado com sucesso."
        }
        else {
            Write-Host "Projeto [$projeto] da concessão [$nomeConcessao] não foi criado por este script."
        }
    }
}

#-------------------------------------------------------
# Início da execução
#-------------------------------------------------------
$login = $null

try {
    $login = New-PWLogin

    if (-not $login) {
        Write-Host "Login no ProjectWise não realizado."
        return
    }

    if ($noGUI -eq 'true') {
        $dados = @(
            [PSCustomObject]@{
                'Nome Concessao'        = (Normalizar-Texto $nomeConcessao)
                'Sigla Concessao'       = (Normalizar-Texto $siglaConcessao)
                'Nome Projeto'          = (Normalizar-Texto $projeto)
                'Descricao'             = (Normalizar-Texto $descricao)
                'Projetista'            = (Normalizar-Texto $projetista)
                'Gestor Engenharia'     = (Normalizar-Texto $gestorEngenharia)
                'Assistente Engenharia' = (Normalizar-Texto $assistenteEngenharia)
                'Poder Concedente'      = (Normalizar-Texto $poderConcedente)
                'TH'                    = (Normalizar-Texto $th)
            }
        )

        CriarProjetos -DadosProjetos $dados
    }
    else {
        SelecionarPlanilhaCadastrosDeProjetos
    }
}
finally {
    if ($login) {
        Undo-PWLogin
    }
}