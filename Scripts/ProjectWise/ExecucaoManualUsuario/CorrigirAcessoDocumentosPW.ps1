param(
    $noGUI,
    [string]$NomeGrupo = "PW-CORRECAO-ACESSO-DOCUMENTOS",
    [string]$NomeConcessao,
    [string[]]$Projetos,
    [switch]$AplicarWorkflowStates,
    [switch]$NaoPausarNoFinal
)

# Script corretivo para aplicar acesso somente leitura/download em pastas e
# documentos de projetos existentes no ProjectWise.
# Nao remove permissoes existentes, nao altera workflow/state e nao concede
# permissao de edicao, criacao, exclusao ou escrita.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module PWPS_DAB -ErrorAction Stop -WarningAction SilentlyContinue

$transcriptIniciado = $false
$diretorioLogs = Join-Path $PSScriptRoot "Logs"
$arquivoLog = Join-Path $diretorioLogs ("CorrigirAcessoDocumentosPW_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$nomesPossiveisEngenhariaRaiz = @("Engenharia", "Engineering")
$nomesPossiveisPastaProjetos = @("Projetos", "Projeto", "Projects", "Project")

$folderAccess = @("r")
$documentAccess = @("r", "fr")

$pastasProjeto = @(
    "",
    "0 - Documentos Gerais",
    "1 - Area de Trabalho",
    "2 - Unidade",
    "Previsao de Documentos (LD)",
    "Area de Transferencia"
)

$workflowStatesConhecidos = @(
    @{
        Workflow = "Workflow - Engenharia"
        States = @(
            "Em elaboracao",
            "Em analise do Assistente",
            "Em analise da Engenharia",
            "Em analise Especifica",
            "Em consolidacao pela Engenharia",
            "Em Envio para Projetista",
            "Em revisao pelo Projetista",
            "Em Envio para Unidade",
            "Concluido",
            "Superado"
        )
    },
    @{
        Workflow = "Workflow - Engenharia - Unidade"
        States = @(
            "Em analise pela Unidade",
            "Em analise pela Engenharia",
            "Enviado ao Poder Concedente",
            "Concluido pela Unidade",
            "Nova emissao sendo analisada pela Engenharia",
            "Enviado ao Poder Concedente - Nova emissao em analise Eng",
            "Concluido - Nova emissao em analise Eng",
            "Superado"
        )
    }
)

$resumo = [ordered]@{
    ProjetosProcessados = 0
    PastasEncontradas = 0
    PastasNaoEncontradas = 0
    PermissoesPastaAplicadas = 0
    PermissoesDocumentoAplicadas = 0
    PermissoesWorkflowStateAplicadas = 0
    Erros = 0
}

#-------------------------------------------------------
# Funcoes auxiliares
#-------------------------------------------------------
function Normalizar-Texto {
    param([object]$Valor)

    if ($null -eq $Valor) {
        return ''
    }

    return ([string]$Valor).Trim()
}

function Escrever-LogParametro {
    param(
        [string]$Nome,
        [object]$Valor
    )

    Write-Host ("{0,-35}: [{1}]" -f $Nome, ([string]$Valor))
}

function Incrementar-Contador {
    param(
        [string]$Nome,
        [int]$Valor = 1
    )

    $script:resumo[$Nome] = [int]$script:resumo[$Nome] + $Valor
}

function Obter-ValorSeguroPropriedade {
    param(
        [object]$Objeto,
        [string[]]$PossiveisNomes
    )

    if ($null -eq $Objeto) {
        return ""
    }

    foreach ($nome in $PossiveisNomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = ([string]$propriedade.Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($valor)) {
                return $valor
            }
        }
    }

    return ""
}

function Obter-IdPasta {
    param([object]$Pasta)

    return Obter-ValorSeguroPropriedade -Objeto $Pasta -PossiveisNomes @("ProjectID", "FolderID", "Id", "ID")
}

function Obter-NomeAmigavelItem {
    param([object]$Item)

    return Obter-ValorSeguroPropriedade -Objeto $Item -PossiveisNomes @("Name", "FolderName", "ProjectName", "ObjectName", "FolderPath", "FullPath", "Path")
}

function Ordenar-ItensPorNome {
    param([array]$Itens)

    return @(
        $Itens | Sort-Object -Property @{
            Expression = {
                (Obter-NomeAmigavelItem -Item $_).ToLower()
            }
        }
    )
}

function SolicitarNomeGrupo {
    param([string]$ValorPadrao)

    $ValorPadrao = Normalizar-Texto $ValorPadrao

    Write-Host ""
    Write-Host "===== GRUPO GLOBAL DE CORRECAO ====="
    Write-Host "O nome do grupo sempre deve ser informado/confirmado pelo usuario."
    $nomeInformado = Read-Host "Nome do grupo [$ValorPadrao]"
    $nomeInformado = Normalizar-Texto $nomeInformado

    if ([string]::IsNullOrWhiteSpace($nomeInformado)) {
        return $ValorPadrao
    }

    return $nomeInformado
}

function SolicitarNomeConcessao {
    param([string]$ValorAtual)

    $ValorAtual = Normalizar-Texto $ValorAtual

    if (-not [string]::IsNullOrWhiteSpace($ValorAtual)) {
        return $ValorAtual
    }

    Write-Host ""
    Write-Host "===== CONCESSAO ====="
    return (Normalizar-Texto (Read-Host "Informe o NomeConcessao exatamente como esta no caminho ENGENHARIA\<NomeConcessao>\Projetos"))
}

function SolicitarProjetos {
    param([string[]]$ValoresAtuais)

    $projetosNormalizados = @($ValoresAtuais | ForEach-Object { Normalizar-Texto $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($projetosNormalizados.Count -gt 0) {
        return $projetosNormalizados
    }

    Write-Host ""
    Write-Host "===== PROJETOS ====="
    Write-Host "Informe um ou mais projetos separados por virgula."
    $entrada = Read-Host "Projetos"
    $entrada = Normalizar-Texto $entrada

    if ([string]::IsNullOrWhiteSpace($entrada)) {
        return @()
    }

    return @($entrada -split "," | ForEach-Object { Normalizar-Texto $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Selecionar-ItemDialog {
    param(
        [array]$Itens,
        [string]$Titulo,
        [string]$Descricao = ""
    )

    if (-not $Itens -or $Itens.Count -eq 0) {
        throw "Nenhum item disponivel para selecao."
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Titulo
    $form.Size = New-Object System.Drawing.Size(560, 540)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Descricao
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(15, 45)
    $listBox.Size = New-Object System.Drawing.Size(510, 390)
    $listBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($listBox)

    foreach ($item in $Itens) {
        $nome = Obter-NomeAmigavelItem -Item $item
        $null = $listBox.Items.Add([PSCustomObject]@{
            Texto = $nome
            Valor = $item
        })
    }

    $listBox.DisplayMember = "Texto"

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(350, 455)
    $btnOk.Size = New-Object System.Drawing.Size(80, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(445, 455)
    $btnCancelar.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancelar)

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancelar

    $resultado = $form.ShowDialog()

    if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Selecao cancelada pelo operador."
    }

    if (-not $listBox.SelectedItem) {
        throw "Nenhum item foi selecionado."
    }

    return $listBox.SelectedItem.Valor
}

function Selecionar-ProjetosDialog {
    param(
        [array]$Projetos,
        [string]$NomeConcessao
    )

    if (-not $Projetos -or $Projetos.Count -eq 0) {
        throw "Nenhum projeto disponivel para selecao."
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selecionar projetos"
    $form.Size = New-Object System.Drawing.Size(720, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Concessao: $NomeConcessao"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $form.Controls.Add($label)

    $labelFiltro = New-Object System.Windows.Forms.Label
    $labelFiltro.Text = "Filtro:"
    $labelFiltro.AutoSize = $true
    $labelFiltro.Location = New-Object System.Drawing.Point(15, 48)
    $form.Controls.Add($labelFiltro)

    $txtFiltro = New-Object System.Windows.Forms.TextBox
    $txtFiltro.Location = New-Object System.Drawing.Point(65, 44)
    $txtFiltro.Size = New-Object System.Drawing.Size(610, 24)
    $txtFiltro.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($txtFiltro)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(15, 80)
    $checkedList.Size = New-Object System.Drawing.Size(660, 405)
    $checkedList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $checkedList.CheckOnClick = $true
    $checkedList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($checkedList)

    $btnTodos = New-Object System.Windows.Forms.Button
    $btnTodos.Text = "Marcar todos"
    $btnTodos.Location = New-Object System.Drawing.Point(15, 505)
    $btnTodos.Size = New-Object System.Drawing.Size(105, 30)
    $btnTodos.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $form.Controls.Add($btnTodos)

    $btnLimpar = New-Object System.Windows.Forms.Button
    $btnLimpar.Text = "Limpar"
    $btnLimpar.Location = New-Object System.Drawing.Point(130, 505)
    $btnLimpar.Size = New-Object System.Drawing.Size(80, 30)
    $btnLimpar.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $form.Controls.Add($btnLimpar)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(505, 505)
    $btnOk.Size = New-Object System.Drawing.Size(80, 30)
    $btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(595, 505)
    $btnCancelar.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancelar.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancelar)

    $mapaMarcados = @{}

    function Atualizar-ListaProjetos {
        foreach ($item in $checkedList.CheckedItems) {
            if ($item -and $item.PSObject.Properties['Nome']) {
                $mapaMarcados[$item.Nome] = $true
            }
        }

        $checkedList.Items.Clear()
        $filtro = (Normalizar-Texto $txtFiltro.Text)

        foreach ($projeto in $Projetos) {
            $nomeProjeto = Obter-NomeAmigavelItem -Item $projeto
            if (-not [string]::IsNullOrWhiteSpace($filtro) -and (Normalizar-Texto $nomeProjeto) -notlike "*$filtro*") {
                continue
            }

            $itemLista = [PSCustomObject]@{
                Nome = $nomeProjeto
                Valor = $projeto
            }

            $indice = $checkedList.Items.Add($itemLista)
            if ($mapaMarcados.ContainsKey($nomeProjeto)) {
                $checkedList.SetItemChecked($indice, $true)
            }
        }

        $checkedList.DisplayMember = "Nome"
    }

    $txtFiltro.Add_TextChanged({ Atualizar-ListaProjetos })
    $checkedList.Add_ItemCheck({
        param($sender, $e)

        $item = $sender.Items[$e.Index]
        if ($item -and $item.PSObject.Properties['Nome']) {
            if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
                $mapaMarcados[$item.Nome] = $true
            }
            else {
                $mapaMarcados.Remove($item.Nome)
            }
        }
    })
    $btnTodos.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $true)
        }
    })
    $btnLimpar.Add_Click({
        $mapaMarcados.Clear()
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $false)
        }
    })

    Atualizar-ListaProjetos

    $resultado = $form.ShowDialog()

    if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Selecao cancelada pelo operador."
    }

    foreach ($item in $checkedList.CheckedItems) {
        if ($item -and $item.PSObject.Properties['Nome']) {
            $mapaMarcados[$item.Nome] = $true
        }
    }

    $selecionados = @($Projetos | Where-Object {
        $nomeProjeto = Obter-NomeAmigavelItem -Item $_
        $mapaMarcados.ContainsKey($nomeProjeto)
    })
    if ($selecionados.Count -eq 0) {
        throw "Nenhum projeto foi selecionado."
    }

    return $selecionados
}

function Localizar-PastaPorPossiveisNomes {
    param(
        [array]$Pastas,
        [string[]]$NomesPossiveis,
        [string]$Descricao = "Pasta"
    )

    if (-not $Pastas -or $Pastas.Count -eq 0) {
        throw "$Descricao nao encontrada. Nenhuma pasta foi retornada para pesquisa."
    }

    $nomesNormalizados = @($NomesPossiveis | ForEach-Object { Normalizar-Texto $_ })

    $encontrada = $Pastas | Where-Object {
        $nome = Normalizar-Texto (Obter-NomeAmigavelItem -Item $_)
        $nomesNormalizados -contains $nome
    } | Select-Object -First 1

    if (-not $encontrada) {
        $encontrada = $Pastas | Where-Object {
            $nome = Normalizar-Texto (Obter-NomeAmigavelItem -Item $_)
            foreach ($possivel in $nomesNormalizados) {
                if ($nome -like "*$possivel*") {
                    return $true
                }
            }
            return $false
        } | Select-Object -First 1
    }

    if (-not $encontrada) {
        $opcoes = @($Pastas | ForEach-Object { Obter-NomeAmigavelItem -Item $_ }) -join ", "
        throw "$Descricao nao encontrada. Nomes esperados: $($NomesPossiveis -join ' / '). Encontradas: $opcoes"
    }

    return $encontrada
}

function Get-PWRootFoldersSafe {
    $pastas = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    $pastas = @($pastas | Where-Object { $_ -ne $null })

    if ($pastas.Count -eq 0) {
        throw "Nenhuma pasta retornada na raiz."
    }

    return $pastas
}

function Get-PWChildFoldersSafe {
    param([string]$FolderId)

    $pastas = @(Get-PWFoldersImmediateChildren -FolderID $FolderId -ErrorAction Stop)
    $pastas = @($pastas | Where-Object { $_ -ne $null })

    if ($pastas.Count -eq 0) {
        throw "Nenhuma subpasta retornada para o FolderID $FolderId."
    }

    return $pastas
}

function ObterProjetosParaProcessamento {
    param(
        [string]$NomeConcessao,
        [string[]]$Projetos
    )

    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projetos = @($Projetos | ForEach-Object { Normalizar-Texto $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if (-not [string]::IsNullOrWhiteSpace($NomeConcessao) -and $Projetos.Count -gt 0) {
        return @($Projetos | ForEach-Object {
            [PSCustomObject]@{
                NomeConcessao = $NomeConcessao
                Projeto = $_
            }
        })
    }

    if ($noGUI -eq 'true') {
        $NomeConcessao = SolicitarNomeConcessao -ValorAtual $NomeConcessao
        $Projetos = SolicitarProjetos -ValoresAtuais $Projetos

        return @($Projetos | ForEach-Object {
            [PSCustomObject]@{
                NomeConcessao = $NomeConcessao
                Projeto = $_
            }
        })
    }

    Write-Host "Selecionando concessao e projetos a partir do ProjectWise..."
    $pastasRaiz = Get-PWRootFoldersSafe
    $pastaEngenhariaRaiz = Localizar-PastaPorPossiveisNomes -Pastas $pastasRaiz -NomesPossiveis $nomesPossiveisEngenhariaRaiz -Descricao "Pasta de engenharia na raiz"
    $idEngenharia = Obter-IdPasta -Pasta $pastaEngenhariaRaiz

    if ([string]::IsNullOrWhiteSpace($idEngenharia)) {
        throw "Nao foi possivel obter o ID da pasta de engenharia na raiz."
    }

    $concessoes = Get-PWChildFoldersSafe -FolderId $idEngenharia
    $concessoes = Ordenar-ItensPorNome -Itens $concessoes
    $concessaoSelecionada = Selecionar-ItemDialog -Itens $concessoes -Titulo "Selecionar concessao" -Descricao "Escolha a concessao:"
    $nomeConcessaoSelecionada = Obter-NomeAmigavelItem -Item $concessaoSelecionada

    $idConcessao = Obter-IdPasta -Pasta $concessaoSelecionada
    if ([string]::IsNullOrWhiteSpace($idConcessao)) {
        throw "Nao foi possivel obter o ID da concessao selecionada."
    }

    $pastasConcessao = Get-PWChildFoldersSafe -FolderId $idConcessao
    $pastaProjetos = Localizar-PastaPorPossiveisNomes -Pastas $pastasConcessao -NomesPossiveis $nomesPossiveisPastaProjetos -Descricao "Pasta de projetos da concessao"
    $idPastaProjetos = Obter-IdPasta -Pasta $pastaProjetos

    if ([string]::IsNullOrWhiteSpace($idPastaProjetos)) {
        throw "Nao foi possivel obter o ID da pasta Projetos."
    }

    $projetosDisponiveis = Get-PWChildFoldersSafe -FolderId $idPastaProjetos
    $projetosDisponiveis = Ordenar-ItensPorNome -Itens $projetosDisponiveis
    $projetosSelecionados = Selecionar-ProjetosDialog -Projetos $projetosDisponiveis -NomeConcessao $nomeConcessaoSelecionada

    return @($projetosSelecionados | ForEach-Object {
        [PSCustomObject]@{
            NomeConcessao = $nomeConcessaoSelecionada
            Projeto = (Obter-NomeAmigavelItem -Item $_)
        }
    })
}

function MontarCaminhoProjeto {
    param(
        [string]$NomeConcessao,
        [string]$Projeto,
        [string]$PastaRelativa
    )

    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projeto = Normalizar-Texto $Projeto
    $PastaRelativa = Normalizar-Texto $PastaRelativa

    $raizProjeto = "ENGENHARIA\$NomeConcessao\Projetos\$Projeto"

    if ([string]::IsNullOrWhiteSpace($PastaRelativa)) {
        return $raizProjeto
    }

    return "$raizProjeto\$PastaRelativa"
}

function ExibirResumoFinal {
    Write-Host ""
    Write-Host "===== RESUMO FINAL ====="
    Escrever-LogParametro -Nome "Projetos processados"                -Valor $resumo.ProjetosProcessados
    Escrever-LogParametro -Nome "Pastas encontradas"                  -Valor $resumo.PastasEncontradas
    Escrever-LogParametro -Nome "Pastas nao encontradas"              -Valor $resumo.PastasNaoEncontradas
    Escrever-LogParametro -Nome "Permissoes de pasta aplicadas"       -Valor $resumo.PermissoesPastaAplicadas
    Escrever-LogParametro -Nome "Permissoes de documento aplicadas"   -Valor $resumo.PermissoesDocumentoAplicadas
    Escrever-LogParametro -Nome "Permissoes workflow/state aplicadas" -Valor $resumo.PermissoesWorkflowStateAplicadas
    Escrever-LogParametro -Nome "Erros"                               -Valor $resumo.Erros
    Write-Host "========================="
    Write-Host ""
}

function Pausar-NoFinal {
    param([switch]$NaoPausar)

    if ($NaoPausar) {
        return
    }

    Write-Host ""
    Write-Host "Execucao finalizada. Revise as mensagens acima."
    Write-Host "Log gravado em: $arquivoLog"
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
}

#-------------------------------------------------------
# Funcoes de grupo
#-------------------------------------------------------
function GarantirGrupoCorrecao {
    param([string]$NomeGrupo)

    $NomeGrupo = Normalizar-Texto $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($NomeGrupo)) {
        throw "Nome do grupo nao pode ficar vazio."
    }

    Write-Host ""
    Write-Host "Validando grupo global de correcao: $NomeGrupo"

    try {
        $grupo = Get-PWGroups -GroupName $NomeGrupo -ErrorAction Stop

        if (-not $grupo) {
            Write-Host "Grupo nao encontrado. Criando grupo: $NomeGrupo"
            $grupo = New-PWGroupByName -GroupName $NomeGrupo -GroupDescription "Grupo global para correcao de acesso somente leitura/download em documentos de projetos." -ErrorAction Stop
            Write-Host "Grupo criado com sucesso: $NomeGrupo"
        }
        else {
            Write-Host "Grupo ja existe e sera reutilizado: $NomeGrupo"
        }

        return $grupo
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao garantir grupo [$NomeGrupo]: $($_.Exception.Message)"
        throw
    }
}

#-------------------------------------------------------
# Funcoes de permissao
#-------------------------------------------------------
function TestarPastaPW {
    param([string]$FolderPath)

    $FolderPath = Normalizar-Texto $FolderPath

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        Incrementar-Contador -Nome "PastasNaoEncontradas"
        Write-Host "FolderPath vazio. Permissoes nao serao aplicadas."
        return $false
    }

    try {
        $pasta = Get-PWFolders -FolderPath $FolderPath -PopulatePaths -JustOne -ErrorAction Stop

        if (-not $pasta) {
            Incrementar-Contador -Nome "PastasNaoEncontradas"
            Write-Host "Pasta nao encontrada: $FolderPath"
            return $false
        }

        Incrementar-Contador -Nome "PastasEncontradas"
        Write-Host "Pasta encontrada: $FolderPath"
        return $true
    }
    catch {
        Incrementar-Contador -Nome "PastasNaoEncontradas"
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao validar pasta [$FolderPath]: $($_.Exception.Message)"
        return $false
    }
}

function AplicarFolderSecurity {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    try {
        Write-Host "Aplicando FolderSecurity para grupo [$NomeGrupo] em [$FolderPath] com acesso: $($folderAccess -join ', ')"

        Update-PWFolderSecurity `
            -Verbose `
            -InputFolder $FolderPath `
            -FolderSecurity `
            -MemberType Group `
            -MemberName $NomeGrupo `
            -MemberAccess $folderAccess `
            -ErrorAction Stop

        Incrementar-Contador -Nome "PermissoesPastaAplicadas"
        Write-Host "FolderSecurity aplicada com sucesso."
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao aplicar FolderSecurity em [$FolderPath]: $($_.Exception.Message)"
    }
}

function AplicarDocumentSecurityBasica {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    try {
        Write-Host "Aplicando DocumentSecurity para grupo [$NomeGrupo] em [$FolderPath] com acesso: $($documentAccess -join ', ')"

        Update-PWFolderSecurity `
            -Verbose `
            -InputFolder $FolderPath `
            -DocumentSecurity `
            -MemberType Group `
            -MemberName $NomeGrupo `
            -MemberAccess $documentAccess `
            -ErrorAction Stop

        Incrementar-Contador -Nome "PermissoesDocumentoAplicadas"
        Write-Host "DocumentSecurity aplicada com sucesso."
    }
    catch {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Erro ao aplicar DocumentSecurity em [$FolderPath]: $($_.Exception.Message)"
    }
}

function AplicarDocumentSecurityWorkflowState {
    param(
        [string]$FolderPath,
        [string]$NomeGrupo
    )

    foreach ($workflowConfig in $workflowStatesConhecidos) {
        $workflow = $workflowConfig.Workflow

        foreach ($state in $workflowConfig.States) {
            try {
                Write-Host "Aplicando DocumentSecurity por workflow/state em [$FolderPath]"
                Escrever-LogParametro -Nome "Workflow" -Valor $workflow
                Escrever-LogParametro -Nome "State"    -Valor $state

                Update-PWFolderSecurity `
                    -Verbose `
                    -InputFolder $FolderPath `
                    -DocumentSecurity `
                    -WorkFlowName $workflow `
                    -StateName $state `
                    -MemberType Group `
                    -MemberName $NomeGrupo `
                    -MemberAccess $documentAccess `
                    -ErrorAction Stop

                Incrementar-Contador -Nome "PermissoesWorkflowStateAplicadas"
                Write-Host "DocumentSecurity workflow/state aplicada com sucesso."
            }
            catch {
                Incrementar-Contador -Nome "Erros"
                Write-Host "Erro ao aplicar workflow/state [$workflow / $state] em [$FolderPath]: $($_.Exception.Message)"
            }
        }
    }
}

#-------------------------------------------------------
# Funcao principal por projeto
#-------------------------------------------------------
function CorrigirAcessoProjeto {
    param(
        [string]$NomeConcessao,
        [string]$Projeto,
        [string]$NomeGrupo,
        [switch]$AplicarWorkflowStates
    )

    $NomeConcessao = Normalizar-Texto $NomeConcessao
    $Projeto = Normalizar-Texto $Projeto
    $NomeGrupo = Normalizar-Texto $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($Projeto)) {
        Incrementar-Contador -Nome "Erros"
        Write-Host "Projeto vazio recebido. Item ignorado."
        return
    }

    Incrementar-Contador -Nome "ProjetosProcessados"

    Write-Host ""
    Write-Host "===== INICIO DA CORRECAO DO PROJETO ====="
    Escrever-LogParametro -Nome "NomeConcessao" -Valor $NomeConcessao
    Escrever-LogParametro -Nome "Projeto"       -Valor $Projeto
    Escrever-LogParametro -Nome "Grupo"         -Valor $NomeGrupo
    Write-Host "========================================="

    foreach ($pastaRelativa in $pastasProjeto) {
        $folderPath = MontarCaminhoProjeto -NomeConcessao $NomeConcessao -Projeto $Projeto -PastaRelativa $pastaRelativa

        try {
            Write-Host ""
            Write-Host "----- Pasta alvo -----"
            Write-Host $folderPath

            if (-not (TestarPastaPW -FolderPath $folderPath)) {
                continue
            }

            AplicarFolderSecurity -FolderPath $folderPath -NomeGrupo $NomeGrupo
            AplicarDocumentSecurityBasica -FolderPath $folderPath -NomeGrupo $NomeGrupo

            if ($AplicarWorkflowStates) {
                AplicarDocumentSecurityWorkflowState -FolderPath $folderPath -NomeGrupo $NomeGrupo
            }
        }
        catch {
            Incrementar-Contador -Nome "Erros"
            Write-Host "Erro inesperado na pasta [$folderPath]: $($_.Exception.Message)"
            continue
        }
    }
}

#-------------------------------------------------------
# Inicio da execucao
#-------------------------------------------------------
$login = $null

try {
    if (-not (Test-Path $diretorioLogs)) {
        New-Item -Path $diretorioLogs -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $arquivoLog -Append | Out-Null
    $transcriptIniciado = $true

    $login = New-PWLogin -ErrorAction Stop

    if (-not $login) {
        Write-Host "Login no ProjectWise nao realizado."
        return
    }

    $dadosProjetos = ObterProjetosParaProcessamento -NomeConcessao $NomeConcessao -Projetos $Projetos

    Write-Host ""
    Write-Host "===== PARAMETROS RECEBIDOS ====="
    Escrever-LogParametro -Nome "Projetos para processar" -Valor $dadosProjetos.Count
    Escrever-LogParametro -Nome "AplicarWorkflowStates"  -Valor $AplicarWorkflowStates
    Write-Host "================================"

    if (-not $dadosProjetos -or $dadosProjetos.Count -eq 0) {
        throw "Nenhum projeto valido foi informado ou selecionado. A planilha deve conter as colunas 'Nome Concessao' e 'Nome Projeto'."
    }

    $NomeGrupo = SolicitarNomeGrupo -ValorPadrao $NomeGrupo

    if ([string]::IsNullOrWhiteSpace($NomeGrupo)) {
        throw "Nome do grupo nao pode ficar vazio."
    }

    $null = GarantirGrupoCorrecao -NomeGrupo $NomeGrupo

    foreach ($dadosProjeto in $dadosProjetos) {
        CorrigirAcessoProjeto `
            -NomeConcessao $dadosProjeto.NomeConcessao `
            -Projeto $dadosProjeto.Projeto `
            -NomeGrupo $NomeGrupo `
            -AplicarWorkflowStates:$AplicarWorkflowStates
    }
}
catch {
    Incrementar-Contador -Nome "Erros"
    Write-Host "Erro geral na execucao: $($_.Exception.Message)"
}
finally {
    if ($login) {
        Undo-PWLogin
    }

    ExibirResumoFinal

    if ($transcriptIniciado) {
        try {
            Stop-Transcript | Out-Null
        }
        catch {
            Write-Host "Nao foi possivel encerrar o transcript: $($_.Exception.Message)"
        }
    }

    Pausar-NoFinal -NaoPausar:$NaoPausarNoFinal
}
