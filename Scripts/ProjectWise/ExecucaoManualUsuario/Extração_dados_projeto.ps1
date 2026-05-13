# =======================================================
# Extração 1 - Projetos criados em uma concessão ProjectWise
# Versão usando a mesma lógica do script de gestão de acessos:
# - Get-PWFoldersImmediateChildren -Root
# - Get-PWFoldersImmediateChildren -FolderID
# =======================================================

Add-Type -AssemblyName System.Windows.Forms

# =======================================================
# Configurações
# =======================================================
$DatasourceName = '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01'
$UserName = 'admin'
$PasswordPlainText = '123456'

$NomesPossiveisEngenhariaRaiz = @('Engenharia', 'Engineering')
$NomesPossiveisPastaProjetos = @('Projetos', 'Projeto', 'Projects', 'Project')

# False = extrai apenas os projetos diretamente dentro da pasta Projetos.
# True  = extrai também subpastas abaixo dos projetos.
$BuscarProjetosRecursivo = $false
$ProfundidadeMaxima = 5

# =======================================================
# Utilitários
# =======================================================
function Obter-ValorSeguroPropriedade {
    param(
        [object]$Objeto,
        [string[]]$PossiveisNomes
    )

    if ($null -eq $Objeto) { return '' }

    foreach ($nome in $PossiveisNomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = $propriedade.Value.ToString().Trim()
            if ($valor -ne '') { return $valor }
        }
    }

    return ''
}

function Obter-IdPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        'ProjectID', 'FolderID', 'Id', 'ID'
    )
}

function Obter-NomePasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        'Name', 'FolderName', 'ProjectName', 'ObjectName', 'FolderPath', 'FullPath', 'Path'
    )
}

function Obter-DescricaoPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @(
        'Description', 'Descricao', 'ProjectDescription', 'FolderDescription'
    )
}

function Normalizar-Texto {
    param([string]$Texto)

    if ([string]::IsNullOrWhiteSpace($Texto)) { return '' }

    return (($Texto.Trim().ToLowerInvariant()) -replace '\s+', ' ')
}

function Localizar-PastaPorPossiveisNomes {
    param(
        [array]$Pastas,
        [string[]]$NomesPossiveis,
        [string]$Descricao = 'Pasta'
    )

    if (-not $Pastas -or $Pastas.Count -eq 0) {
        throw "$Descricao não encontrada. Nenhuma pasta foi retornada."
    }

    $nomesNormalizados = @($NomesPossiveis | ForEach-Object { Normalizar-Texto $_ })

    $encontrada = $Pastas | Where-Object {
        $nome = Normalizar-Texto (Obter-NomePasta -Pasta $_)
        $nomesNormalizados -contains $nome
    } | Select-Object -First 1

    if (-not $encontrada) {
        $encontrada = $Pastas | Where-Object {
            $nome = Normalizar-Texto (Obter-NomePasta -Pasta $_)
            foreach ($possivel in $nomesNormalizados) {
                if ($nome -like "*$possivel*") { return $true }
            }
            return $false
        } | Select-Object -First 1
    }

    if (-not $encontrada) {
        $opcoes = @($Pastas | ForEach-Object { Obter-NomePasta -Pasta $_ }) -join ', '
        throw "$Descricao não encontrada. Nomes esperados: $($NomesPossiveis -join ' / '). Encontradas: $opcoes"
    }

    return $encontrada
}

function Selecionar-PastaSaida {
    Write-Host ''
    Write-Host 'Selecione onde deseja salvar o arquivo Excel'
    Write-Host '----------------------------------------'
    Write-Host '1 - Escolher pasta pela janela'
    Write-Host '2 - Informar caminho manualmente'

    do {
        $opcao = Read-Host 'Escolha uma opção (1/2)'
        $opcao = $opcao.Trim()
    } while ($opcao -notin @('1','2'))

    if ($opcao -eq '1') {
        $dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialogo.Description = 'Selecione a pasta onde o Excel será salvo'
        $dialogo.ShowNewFolderButton = $true

        if ($dialogo.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if (Test-Path -Path $dialogo.SelectedPath -PathType Container) {
                return $dialogo.SelectedPath
            }
        }
    }

    Write-Host ''
    Write-Host 'Informe o caminho completo da pasta de saída.'
    Write-Host 'Exemplos:'
    Write-Host 'C:\Temp'
    Write-Host '\\servidor\relatorios\projectwise'
    Write-Host ''

    $tentativas = 0
    do {
        $tentativas++
        $caminho = (Read-Host 'Caminho').Trim().Trim('"')

        if (-not [string]::IsNullOrWhiteSpace($caminho)) {
            if (-not (Test-Path -Path $caminho -PathType Container)) {
                try { New-Item -Path $caminho -ItemType Directory -Force | Out-Null } catch {}
            }

            if (Test-Path -Path $caminho -PathType Container) { return $caminho }
        }

        Write-Host 'Caminho inválido ou inacessível. Tente novamente.'
    } while ($tentativas -lt 3)

    return $null
}

function Selecionar-ItemConsole {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Itens,

        [Parameter(Mandatory = $true)]
        [string]$Titulo
    )

    if (-not $Itens -or $Itens.Count -eq 0) { return $null }

    Write-Host ''
    Write-Host $Titulo
    Write-Host '----------------------------------------'

    for ($i = 0; $i -lt $Itens.Count; $i++) {
        $numero = $i + 1
        $nome = Obter-NomePasta -Pasta $Itens[$i]
        $descricao = Obter-DescricaoPasta -Pasta $Itens[$i]

        if ([string]::IsNullOrWhiteSpace($descricao)) {
            Write-Host "$numero - $nome"
        }
        else {
            Write-Host "$numero - $nome | $descricao"
        }
    }

    Write-Host ''

    do {
        $entrada = Read-Host 'Digite o número desejado'
        $indice = 0
        $ok = [int]::TryParse($entrada, [ref]$indice)
    } while (-not $ok -or $indice -lt 1 -or $indice -gt $Itens.Count)

    return $Itens[$indice - 1]
}

function Get-PWRootFoldersSafe {
    Write-Host 'Listando pastas da raiz do datasource...'
    $pastas = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    $pastas = @($pastas | Where-Object { $_ -ne $null })

    if ($pastas.Count -eq 0) {
        throw 'Nenhuma pasta foi retornada na raiz do datasource.'
    }

    return $pastas
}

function Get-PWChildFoldersSafe {
    param([string]$FolderId)

    if ([string]::IsNullOrWhiteSpace($FolderId)) {
        throw 'FolderId vazio ao tentar buscar subpastas.'
    }

    $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
    return @($pastas | Where-Object { $_ -ne $null })
}

function Obter-ProjetosRecursivo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderIdInicial,

        [int]$ProfundidadeMaxima = 5
    )

    $resultado = New-Object System.Collections.Generic.List[object]
    $fila = New-Object System.Collections.Queue
    $fila.Enqueue([PSCustomObject]@{ FolderId = $FolderIdInicial; Nivel = 0 })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $filhos = @(Get-PWChildFoldersSafe -FolderId $atual.FolderId)

        foreach ($filho in $filhos) {
            $resultado.Add($filho)

            $idFilho = Obter-IdPasta -Pasta $filho
            if (-not [string]::IsNullOrWhiteSpace($idFilho) -and $atual.Nivel -lt $ProfundidadeMaxima) {
                $fila.Enqueue([PSCustomObject]@{ FolderId = $idFilho; Nivel = ($atual.Nivel + 1) })
            }
        }
    }

    return @($resultado)
}

function Converter-ProjetoParaLinhaExcel {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Projeto,

        [Parameter(Mandatory = $true)]
        [string]$NomeConcessao
    )

    [PSCustomObject]@{
        Concessao         = $NomeConcessao
        NomeProjeto       = Obter-NomePasta -Pasta $Projeto
        Descricao         = Obter-DescricaoPasta -Pasta $Projeto
        ProjectID         = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('ProjectID','FolderID','Id','ID')
        ProjectGUID       = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('ProjectGUID','FolderGUID','GUID')
        ProjectGUIDString = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('ProjectGUIDString','FolderGUIDString','GUIDString')
        FullPath          = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('FullPath','FolderPath','Path')
        CreateDate        = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('CreateDate','CreationDate')
        CreatorName       = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('CreatorName','CreatedBy','OwnerName')
        UpdateDate        = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('UpdateDate','ModifyDate','ModifiedDate')
        UpdaterName       = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('UpdaterName','ModifiedBy')
        IsRichProject     = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('IsRichProject')
        Workflow          = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('Workflow')
        WorkflowState     = Obter-ValorSeguroPropriedade -Objeto $Projeto -PossiveisNomes @('WorkflowState')
    }
}

# =======================================================
# Execução
# =======================================================
try {
    Import-Module PWPS_DAB -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop

    $pastaSaida = Selecionar-PastaSaida
    if (-not $pastaSaida) { throw 'Nenhuma pasta de saída válida foi informada.' }

    $securePassword = ConvertTo-SecureString $PasswordPlainText -AsPlainText -Force

    Write-Host ''
    Write-Host 'Conectando ao ProjectWise...'
    New-PWLogin -DatasourceName $DatasourceName -UserName $UserName -Password $securePassword -ErrorAction Stop
    Write-Host 'Login realizado com sucesso.'

    # 1. Lista a raiz real do datasource.
    $pastasRaiz = Get-PWRootFoldersSafe

    # 2. Localiza Engenharia na raiz.
    $pastaEngenharia = Localizar-PastaPorPossiveisNomes `
        -Pastas $pastasRaiz `
        -NomesPossiveis $NomesPossiveisEngenhariaRaiz `
        -Descricao 'Pasta de engenharia na raiz'

    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenharia
    if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
        throw 'Não foi possível obter o ID da pasta Engenharia.'
    }

    # 3. Lista apenas os filhos imediatos de Engenharia. Estes são as concessões.
    Write-Host ''
    Write-Host 'Buscando concessões dentro de Engenharia...'
    $concessoes = Get-PWChildFoldersSafe -FolderId $idEngenharia
    $concessoes = @($concessoes | Sort-Object { (Obter-NomePasta -Pasta $_).ToLowerInvariant() })

    if (-not $concessoes -or $concessoes.Count -eq 0) {
        throw 'Nenhuma concessão encontrada abaixo da pasta Engenharia.'
    }

    # 4. Usuário escolhe a concessão.
    $concessaoSelecionada = Selecionar-ItemConsole -Itens $concessoes -Titulo 'Concessões encontradas:'
    if (-not $concessaoSelecionada) { throw 'Nenhuma concessão selecionada.' }

    $nomeConcessao = Obter-NomePasta -Pasta $concessaoSelecionada
    $idConcessao = Obter-IdPasta -Pasta $concessaoSelecionada
    if ([string]::IsNullOrWhiteSpace($idConcessao)) {
        throw 'Não foi possível obter o ID da concessão selecionada.'
    }

    # 5. Localiza a pasta Projetos dentro da concessão.
    Write-Host ''
    Write-Host "Buscando pasta de projetos da concessão '$nomeConcessao'..."
    $filhosConcessao = Get-PWChildFoldersSafe -FolderId $idConcessao

    $pastaProjetos = Localizar-PastaPorPossiveisNomes `
        -Pastas $filhosConcessao `
        -NomesPossiveis $NomesPossiveisPastaProjetos `
        -Descricao 'Pasta de projetos da concessão'

    $idPastaProjetos = Obter-IdPasta -Pasta $pastaProjetos
    if ([string]::IsNullOrWhiteSpace($idPastaProjetos)) {
        throw 'Não foi possível obter o ID da pasta Projetos.'
    }

    # 6. Busca projetos somente dentro da pasta Projetos da concessão escolhida.
    Write-Host ''
    Write-Host "Buscando projetos da concessão '$nomeConcessao'..."

    if ($BuscarProjetosRecursivo) {
        $projetos = Obter-ProjetosRecursivo -FolderIdInicial $idPastaProjetos -ProfundidadeMaxima $ProfundidadeMaxima
    }
    else {
        $projetos = Get-PWChildFoldersSafe -FolderId $idPastaProjetos
    }

    $projetos = @($projetos | Sort-Object { (Obter-NomePasta -Pasta $_).ToLowerInvariant() })

    if (-not $projetos -or $projetos.Count -eq 0) {
        throw 'Nenhum projeto encontrado dentro da pasta Projetos.'
    }

    Write-Host "Projetos encontrados: $($projetos.Count)"

    # 7. Monta relatório.
    $dados = foreach ($projeto in $projetos) {
        Converter-ProjetoParaLinhaExcel -Projeto $projeto -NomeConcessao $nomeConcessao
    }

    $nomeArquivoSeguro = $nomeConcessao -replace '[\\/:*?"<>|]', '_'
    $arquivoSaida = Join-Path $pastaSaida ('Extracao_Projetos_{0}_{1}.xlsx' -f $nomeArquivoSeguro, (Get-Date -Format 'yyyyMMdd_HHmmss'))

    Write-Host ''
    Write-Host 'Exportando Excel...'

    $dados | Export-Excel `
        -Path $arquivoSaida `
        -WorksheetName 'Projetos' `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow `
        -AutoFilter

    Write-Host ''
    Write-Host 'Extração concluída com sucesso.'
    Write-Host "Arquivo gerado: $arquivoSaida"
}
catch {
    Write-Host ''
    Write-Host 'Ocorreu um erro durante a extração.'
    Write-Host $_.Exception.Message
}
finally {
    try { Undo-PWLogin -ErrorAction SilentlyContinue } catch {}
    Read-Host 'Pressione Enter para encerrar'
}
