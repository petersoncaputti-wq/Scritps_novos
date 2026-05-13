function ObterDocumentosValidados {
    Get-PWDocumentsBySearch -SearchName 'Scripts\Docs GRD Validados' -GetAttributes
}

# ------- Helpers mínimos (PS5) -------
function Resolve-PWFolderPath {
    param([string]$Path)
    if (-not $Path) { return $null }
    $p = $Path.Trim('\')
    if ($p -match '^pw:\\[^\\]+\\(.+)$') { return $Matches[1] } # remove "pw:\Datasource\"
    return $p
}

function Get-PWParentPath {
    param([string]$Path)
    if (-not $Path) { return $null }
    $n = (Resolve-PWFolderPath $Path).TrimEnd('\')
    $i = $n.LastIndexOf('\')
    if ($i -lt 0) { return $null }
    $n.Substring(0, $i)
}

function Get-FolderMetadata {
    param($Folder)
    $env = $null; $wf = $null; $state = $null
    if ($Folder) {
        if ($Folder.PSObject.Properties['EnvironmentName'] -and $Folder.EnvironmentName) {
            $env = $Folder.EnvironmentName
        }
        elseif ($Folder.PSObject.Properties['Environment'] -and $Folder.Environment) {
            if ($Folder.Environment -is [string]) { $env = $Folder.Environment }
            elseif ($Folder.Environment.PSObject.Properties['Name']) { $env = $Folder.Environment.Name }
        }

        if ($Folder.PSObject.Properties['WorkflowName'] -and $Folder.WorkflowName) {
            $wf = $Folder.WorkflowName
        }
        elseif ($Folder.PSObject.Properties['Workflow'] -and $Folder.Workflow) {
            if ($Folder.Workflow -is [string]) { $wf = $Folder.Workflow }
            elseif ($Folder.Workflow.PSObject.Properties['Name']) { $wf = $Folder.Workflow.Name }
        }

        if ($Folder.PSObject.Properties['WorkflowState'] -and $Folder.WorkflowState) {
            $state = $Folder.WorkflowState
        }
        elseif ($Folder.PSObject.Properties['StateName'] -and $Folder.StateName) {
            $state = $Folder.StateName
        }
    }
    [PSCustomObject]@{ Environment = $env; Workflow = $wf; State = $state }
}

# ------- Criação de pasta herdando do pai -------
function GarantirPasta {
    param(
        [string]$Caminho,
        [string]$EnvironmentOverride,
        [string]$WorkflowOverride
    )

    if (-not $Caminho) { return $null }

    $pwPath = Resolve-PWFolderPath $Caminho
    if (-not $pwPath) { return $null }

    $folder = Get-PWFolders -FolderPath $pwPath -JustOne -Slow -ErrorAction SilentlyContinue
    if (-not $folder) {
        $parentPath = Get-PWParentPath $pwPath
        $parentPath = if ($parentPath -match "\\Modelos Autorais$") { Get-PWParentPath $parentPath } else { $parentPath }
        $parentPath = if ($parentPath -match "\\Modelo BIM$") { Get-PWParentPath $parentPath } else { $parentPath }
        $envToApply = $EnvironmentOverride
        $wfToApply = $WorkflowOverride

        if ($parentPath) {
            $parentFolder = Get-PWFolders -FolderPath $parentPath -JustOne -Slow -ErrorAction SilentlyContinue
            if ($parentFolder) {
                $pmeta = Get-FolderMetadata $parentFolder
                if (-not $envToApply -and $pmeta.Environment) { $envToApply = $pmeta.Environment }
                if (-not $wfToApply -and $pmeta.Workflow) { $wfToApply = $pmeta.Workflow }
            }
        }

        $newParams = @{ FolderPath = $pwPath }
        if ($envToApply) { $newParams.Environment = $envToApply }
        if ($wfToApply) { $newParams.Workflow = $wfToApply }

        try { $folder = New-PWFolder @newParams -ErrorAction Stop }
        catch { Write-Warning "Falha ao criar a pasta '$pwPath': $($_.Exception.Message)"; return $null }
    }
    else {
        if ($EnvironmentOverride) {
            try { Set-PWFolderEnvironment -InputFolder $folder -NewEnvironment $EnvironmentOverride -ErrorAction Stop }
            catch { Write-Warning "Falha ao aplicar environment '$EnvironmentOverride' em '$pwPath': $($_.Exception.Message)"; return $null }
        }
        if ($WorkflowOverride) {
            try { Set-PWFolderWorkflow -InputFolder $folder -NewWorkflow $WorkflowOverride -Force -ErrorAction Stop }
            catch { Write-Warning "Falha ao aplicar workflow '$WorkflowOverride' em '$pwPath': $($_.Exception.Message)"; return $null }
        }
    }
    $Caminho
}

function GarantirEstruturaPastaDestino {
    param([string]$pastaDestino)

    $pastaDestino = (Resolve-PWFolderPath $pastaDestino).TrimEnd('\')

    if ($pastaDestino -match '\\Modelo BIM') { return $true }

    # 1) Pasta da disciplina
    $principal = GarantirPasta -Caminho $pastaDestino
    if (-not $principal) { Write-Warning "Estrutura incompleta: '$pastaDestino'."; return $false }

    $disciplinaFolder = Get-PWFolders -FolderPath $pastaDestino -JustOne -Slow -ErrorAction SilentlyContinue
    $disciplinaMeta = Get-FolderMetadata $disciplinaFolder

    # 2) Superados (herda da disciplina)
    $superados = GarantirPasta -Caminho ($pastaDestino + '\Superados') `
        -EnvironmentOverride $disciplinaMeta.Environment `
        -WorkflowOverride    $disciplinaMeta.Workflow
    if (-not $superados) { Write-Warning "Estrutura incompleta: '$pastaDestino\Superados'."; return $false }

    # 3) xfdf__ (env dmsXFDF + wf Workflow - Comentarios)
    $xfdf = GarantirPasta -Caminho ($pastaDestino + '\xfdf__') `
        -EnvironmentOverride 'dmsXFDF' `
        -WorkflowOverride    'Workflow - Comentarios'
    if (-not $xfdf) { Write-Warning "Estrutura incompleta: '$pastaDestino\xfdf__'."; return $false }

    $true
}

# ------- Regras de negócio -------
function DefineDisciplinaPai {
    param ($cadastroDisciplina, $poderConcedente)
    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) { return $null }
    if ($poderConcedente -eq 'ARTESP') { return $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[1] }
    $cadastroDisciplina.Attributes[0].Relacionamento
}

function DefineFaseProjeto {
    param ($cadastroDisciplina, $cadastroFaseProjeto, $poderConcedente)
    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.Attributes -or $null -eq $cadastroDisciplina.Attributes[0].Relacionamento) { return $null }
    if ($poderConcedente -eq 'ARTESP') { return $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[0] }
    $cadastroFaseProjeto.Attributes[0].Relacionamento
}

function ObtemCadastroDisciplina {
    param ($documento, $poderConcedente)
    if ($null -eq $documento -or $null -eq $documento.Attributes.Disciplina) { return $null }
    $disciplina = $documento.Attributes.Disciplina
    $tipoRegistro = switch ($poderConcedente) {
        'ARTESP' { 'Disciplinas ARTESP' }
        Default { 'Disciplinas ANTT' }
    }
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = $tipoRegistro; Codigo = $disciplina } -GetAttributes
}

function ObtemCadastroFaseProjeto {
    param ($documento, $poderConcedente)
    if ($null -eq $documento -or $null -eq $documento.Attributes.Disciplina) { return $null }
    if ($poderConcedente -eq 'ARTESP') { return $null }
    $faseProjeto = $documento.Attributes.FaseProjeto
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = 'Tipos de Projeto ANTT'; Codigo = $faseProjeto } -GetAttributes
}

# ------- Espelho em 2 - Unidade -------
function GarantirEspelhoEmUnidade {
    param([string]$pastaDestinoAT)

    if (-not $pastaDestinoAT) { return $false }

    $destAT = (Resolve-PWFolderPath $pastaDestinoAT).TrimEnd('\')

    # troca segura de "1 - Area de Trabalho" por "2 - Unidade"
    $destUN = $destAT.Replace('\1 - Area de Trabalho\', '\2 - Unidade\')
    if ($destUN -eq $destAT) {
        $destUN = $destAT.Replace('\1 - Area de Trabalho', '\2 - Unidade')
    }

    if (-not (GarantirPasta -Caminho $destUN)) { return $false }
    if (-not (GarantirEstruturaPastaDestino -PastaDestino $destUN)) { return $false }

    return $true
}

# ------- Cálculo do destino + criação de estrutura (AT) e espelho (UN) -------
function CalculaPastaDestino {
    param ($documento)

    if (-not $documento) { return $null }

    $poderConcedente = $documento.Attributes[0].PoderConcedente
    $cadastroDisciplina = ObtemCadastroDisciplina -Documento $documento -PoderConcedente $poderConcedente

    $disciplina = DefineDisciplinaPai -CadastroDisciplina $cadastroDisciplina -PoderConcedente $poderConcedente
    if (-not $disciplina) { return $null }

    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -Documento $documento -PoderConcedente $poderConcedente
    $faseProjeto = DefineFaseProjeto -CadastroDisciplina $cadastroDisciplina -CadastroFaseProjeto $cadastroFaseProjeto -PoderConcedente $poderConcedente
    if (-not $faseProjeto) { return $null }

    $volume = $documento.Attributes[0].Volume
    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $documento
    if (-not $pastaRaizProjeto) { return $null }

    $base = $pastaRaizProjeto.FullPath.TrimEnd('\')

    # Melhoria: Exceção para Modelos Autorais
    $caminhoEspecialModeloAutoral = if ($documento.Attributes[0].Disciplina -ne 'U4' -and (ValidaSeModeloFederadoAutoral $documento)) { 'Modelo BIM\Modelos Autorais' } else { '' }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -eq 'MI') { 
        $caminhoEspecialModeloAutoral = 'Modelo BIM' 
        $disciplina = '' # MI, por ser Modelo BIM fica direto na raiz dessa pasta e ignora a disciplina
    }

    $destino = @($base, '1 - Area de Trabalho', $faseProjeto, $volume, $caminhoEspecialModeloAutoral, $disciplina) | Where-Object { $_ }
    $destino = $destino -join '\'

    # garante AT
    if (-not (GarantirPasta -Caminho $destino)) { return $null }
    if (-not (GarantirEstruturaPastaDestino -PastaDestino $destino)) { return $null }

    # espelha UN
    if (-not (GarantirEspelhoEmUnidade -pastaDestinoAT $destino)) {
        Write-Warning "Falha ao garantir espelho em '2 - Unidade' para '$destino'."
    }

    $destino
}

# ------- Utilidades de movimentação e states (mantidas para fase 2) -------
function ObtemDocumentosAnterioresNaPastaDestino {
    param ($pastaDestino, $numeroDocumento, $sequencialEmissao)

    $doclist = Get-PWDocumentsBySearch -FolderPath (Resolve-PWFolderPath $pastaDestino) -JustThisFolder `
        -Attributes @{ NumeroPoderConcedente = $numeroDocumento } -GetAttributes
    $doclist.Where({ $_.Attributes[0].SequencialEmissao -lt $sequencialEmissao })
}

function MovimentaDocumentosAnterioresParaPastaSuperado {
    param ($pastaDestino, $documento)

    $numero = $documento.Attributes[0].NumeroPoderConcedente
    $seq = $documento.Attributes[0].SequencialEmissao

    $anteriores = ObtemDocumentosAnterioresNaPastaDestino -PastaDestino $pastaDestino -NumeroDocumento $numero -SequencialEmissao $seq
    if (-not $anteriores) { return }

    $pastaSuperadosPath = (Resolve-PWFolderPath ($pastaDestino + '\Superados'))
    $pastaSuperadosObj = Get-PWFolders -FolderPath $pastaSuperadosPath -JustOne -Slow -ErrorAction SilentlyContinue
    if (-not $pastaSuperadosObj) {
        Write-Warning "Pasta 'Superados' não encontrada após garantir estrutura."
        return
    }

    $moved = Move-PWDocumentsToFolder -InputDocument $anteriores -TargetFolderPath $pastaSuperadosPath
    Set-PWDocumentState -InputDocuments $moved -State 'Superado' -Force
}

function CalculaState {
    param ($stateAtual)
    switch ($stateAtual) {
        'Emitido pela Engenharia' { 'Nova emissao sendo analisada pela Engenharia' }
        'Solicitado reanalise da Engenharia' { 'Nova emissao sendo analisada pela Engenharia' }
        'Enviado ao Poder Concedente' { 'Enviado ao Poder Concedente - Nova emissao em analise Eng' }
        'Concluido pela Unidade' { 'Concluido - Nova emissao em analise Eng' }
        Default { $null }
    }
}

function AtualizaStateDocumentosEmitidosParaUnidade {
    param ($documento, $pastaDestino)

    $pastaUnidade = (Resolve-PWFolderPath $pastaDestino).Replace('1 - Area de Trabalho', '2 - Unidade')
    $numero = $documento.Attributes[0].NumeroPoderConcedente
    $docsUn = Get-PWDocumentsBySearch -FolderPath $pastaUnidade -JustThisFolder `
        -Attributes @{ NumeroPoderConcedente = $numero } -GetAttributes

    if (-not $docsUn -or $docsUn.Attributes[0].SequencialEmissao -ge $documento.Attributes[0].SequencialEmissao) { return }

    $stateAtual = $docsUn[0].WorkflowState
    $proximoState = CalculaState -StateAtual $stateAtual
    if ($proximoState) { Set-PWDocumentState -InputDocuments $docsUn -State $proximoState -Force }
}
function ValidaSeModeloFederadoAutoral {
    param($documento)

    if ($documento.Attributes[0].PoderConcedente -eq 'ANTT' -and $documento.Attributes[0].Sequencial -as [int] -in 800..999) { return $true }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

# ---------------- Execução (fase 1: calcular e criar AT + espelho UN) ----------------
try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch {}

$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-01' -Password $SecurePassword -UserName 'admin' | Out-Null
if (-not (Get-PWCurrentDatasource)) { throw "Login ProjectWise não estabelecido." }

$documentos = ObterDocumentosValidados
foreach ($documento in $documentos) {
    $pastaDestino = CalculaPastaDestino -Documento $documento

    if ($pastaDestino) {
        $destNoDocs = Resolve-PWFolderPath $pastaDestino
        $destObj = Get-PWFolders -FolderPath $destNoDocs -JustOne -Slow -ErrorAction SilentlyContinue
        if ($destObj) {
            Write-Host ("OK: Pastas garantidas (AT e UN) -> {0}" -f $destObj.FullPath)
        }
        else {
            Write-Warning ("Falha ao garantir pasta AT -> {0}" -f $pastaDestino)
        }
    }
    else {
        Write-Warning "Pasta destino não calculada/criada para o documento atual."
    }
}

Undo-PWLogin
# ---------------- fim ----------------
