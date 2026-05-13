# -----------------------------------------------
# Cria Disciplina em "2 - Unidade" (PS5.1)
# -----------------------------------------------

# ===== Config =====
$SearchName_Unidade = 'Scripts\Envio p Unidade'   # ajuste se quiser outra pesquisa

# ===== Entrada (pesquisa salva) =====
function ObterDocumentosParaCriarUnidade {
    Get-PWDocumentsBySearch -SearchName $SearchName_Unidade -GetAttributes
}

# ===== Helpers mínimos =====
function Resolve-PWFolderPath {
    param([string]$Path)
    if (-not $Path) { return $null }
    $p = $Path.Trim('\')
    if ($p -match '^pw:\\[^\\]+\\(.+)$') { return $Matches[1] } # remove "pw:\Datasource\"
    $p
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

function GarantirEstruturaPastaUnidade {
    param([string]$pastaUnidade)

    $pastaUnidade = (Resolve-PWFolderPath $pastaUnidade).TrimEnd('\')

    if ($pastaDestino -match '\\Modelo BIM') { return $true }

    # 1) Pasta da disciplina (UN)
    $principal = GarantirPasta -Caminho $pastaUnidade
    if (-not $principal) { Write-Warning "Estrutura incompleta: '$pastaUnidade'."; return $false }

    $unFolder = Get-PWFolders -FolderPath $pastaUnidade -JustOne -Slow -ErrorAction SilentlyContinue
    $unMeta = Get-FolderMetadata $unFolder

    # 2) Superados (herda env/wf da UN)
    $superados = GarantirPasta -Caminho ($pastaUnidade + '\Superados') `
        -EnvironmentOverride $unMeta.Environment `
        -WorkflowOverride    $unMeta.Workflow
    if (-not $superados) { Write-Warning "Estrutura incompleta: '$pastaUnidade\Superados'."; return $false }

    # 3) xfdf__ (env dmsXFDF + wf de comentários)
    $xfdf = GarantirPasta -Caminho ($pastaUnidade + '\xfdf__') `
        -EnvironmentOverride 'dmsXFDF' `
        -WorkflowOverride    'Workflow - Comentarios'
    if (-not $xfdf) { Write-Warning "Estrutura incompleta: '$pastaUnidade\xfdf__'."; return $false }

    $true
}

# ===== Regras de negócio p/ calcular disciplina/rota =====
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
    if ($null -eq $cadastroFaseProjeto -or $null -eq $cadastroFaseProjeto.Attributes -or $null -eq $cadastroFaseProjeto.Attributes[0].Relacionamento) { return $null }
    $cadastroFaseProjeto.Attributes[0].Relacionamento
}

function ObtemCadastroDisciplina {
    param ($documentoBase, $poderConcedente)
    if ($null -eq $documentoBase -or $null -eq $documentoBase.Attributes -or $null -eq $documentoBase.Attributes[0].Disciplina) { return $null }
    $disciplina = $documentoBase.Attributes[0].Disciplina
    $tipoRegistro = switch ($poderConcedente) {
        'ARTESP' { 'Disciplinas ARTESP' }
        Default { 'Disciplinas ANTT' }
    }
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = $tipoRegistro; Codigo = $disciplina } -GetAttributes
}

function ObtemCadastroFaseProjeto {
    param ($documentoBase, $poderConcedente)
    if ($poderConcedente -eq 'ARTESP') { return $null }
    if ($null -eq $documentoBase -or $null -eq $documentoBase.Attributes -or $null -eq $documentoBase.Attributes[0].FaseProjeto) { return $null }
    $faseProjeto = $documentoBase.Attributes[0].FaseProjeto
    Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{ TipoRegistro = 'Tipos de Projeto ANTT'; Codigo = $faseProjeto } -GetAttributes
}

function TemAtributosParaRoteamento {
    param ($doc, $poderConcedente)
    if ($null -eq $doc -or $null -eq $doc.Attributes -or $null -eq $doc.Attributes[0]) { return $false }
    $a = $doc.Attributes[0]
    if ($poderConcedente -eq 'ARTESP') { return [bool]$a.Disciplina }
    ([bool]$a.Disciplina -and [bool]$a.FaseProjeto)
}

function TentarEncontrarHospedeiro {
    param ($documento)
    $rels = $null
    try { $rels = Get-PWDocumentRelationships -InputDocuments $documento -ReferencedBy -ErrorAction Stop }
    catch {
        try { $rels = Get-PWDocumentReferences -InputDocuments $documento -ReferencedBy -ErrorAction Stop }
        catch { $rels = $null }
    }
    if (-not $rels) { return $null }
    $hospedeiros = $rels | Where-Object { $_.FileName -match '.(dgn|dwg)$' }
    if (-not $hospedeiros) { return $rels[0] }
    $hospedeiros[0]
}

# ===== Cálculo do destino em "2 - Unidade" =====
function CalculaPastaUnidade {
    param ($documento, [ref]$documentoBaseUsado)

    if (-not $documento) { return $null }

    $poderConcedente = $documento.Attributes[0].PoderConcedente
    if (-not $poderConcedente) { return $null }

    $docBase = $documento
    if (-not (TemAtributosParaRoteamento -doc $documento -poderConcedente $poderConcedente)) {
        $hosp = TentarEncontrarHospedeiro -documento $documento
        if ($hosp -and (TemAtributosParaRoteamento -doc $hosp -poderConcedente $poderConcedente)) { $docBase = $hosp } else { return $null }
    }
    $documentoBaseUsado.Value = $docBase

    $cadastroDisciplina = ObtemCadastroDisciplina -documentoBase $docBase -PoderConcedente $poderConcedente
    $disciplina = DefineDisciplinaPai -CadastroDisciplina $cadastroDisciplina -PoderConcedente $poderConcedente
    if (-not $disciplina) { return $null }

    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -documentoBase $docBase -PoderConcedente $poderConcedente
    $faseProjeto = DefineFaseProjeto -CadastroDisciplina $cadastroDisciplina -CadastroFaseProjeto $cadastroFaseProjeto -PoderConcedente $poderConcedente
    if (-not $faseProjeto) { return $null }

    $volume = $docBase.Attributes[0].Volume

    $pastaRaizProjeto = Get-PWRichProjectForDocument -InputDocument $docBase
    if (-not $pastaRaizProjeto) { return $null }

    $base = $pastaRaizProjeto.FullPath.TrimEnd('\')

    # Melhoria: Exceção para Modelos Autorais
    $caminhoEspecialModeloAutoral = if ($documento.Attributes[0].Disciplina -ne 'U4' -and (ValidaSeModeloFederadoAutoral $documento)) { 'Modelo BIM\Modelos Autorais' } else { '' }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -eq 'MI') { 
        $caminhoEspecialModeloAutoral = 'Modelo BIM' 
        $disciplina = '' # MI, por ser Modelo BIM fica direto na raiz dessa pasta e ignora a disciplina
    }

    $destino = @($base, '2 - Unidade', $faseProjeto, $volume, $caminhoEspecialModeloAutoral, $disciplina) | Where-Object { $_ }
    $destino = $destino -join '\'

    return $destino
}
function ValidaSeModeloFederadoAutoral {
    param($documento)

    if ($documento.Attributes[0].PoderConcedente -eq 'ANTT' -and $documento.Attributes[0].Sequencial -as [int] -in 800..999) { return $true }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

# -----------------------------------------------
# Execução (só cria/garante a estrutura na UNIDADE)
# -----------------------------------------------
try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch {}

$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-01' -Password $SecurePassword -UserName 'admin' | Out-Null
if (-not (Get-PWCurrentDatasource)) { throw "Login ProjectWise não estabelecido." }

$documentos = ObterDocumentosParaCriarUnidade

foreach ($documento in $documentos) {
    $docBaseRef = New-Object PSObject
    $pastaUnidade = CalculaPastaUnidade -Documento $documento -documentoBaseUsado ([ref]$docBaseRef)
    if (-not $pastaUnidade) {
        Write-Warning "Destino UN não calculado para: $($documento.FileName)"
        continue
    }

    if (-not (GarantirPasta -Caminho $pastaUnidade)) {
        Write-Warning "Falha ao garantir pasta UN: $pastaUnidade"
        continue
    }
    if (-not (GarantirEstruturaPastaUnidade -pastaUnidade $pastaUnidade)) {
        Write-Warning "Falha ao garantir estrutura em: $pastaUnidade"
        continue
    }

    $okFolder = Get-PWFolders -FolderPath (Resolve-PWFolderPath $pastaUnidade) -JustOne -Slow -ErrorAction SilentlyContinue
    if ($okFolder) {
        Write-Host ("OK: UN criada/garantida -> {0}" -f $okFolder.FullPath)
    }
    else {
        Write-Warning ("Estrutura UN não localizada após criação -> {0}" -f $pastaUnidade)
    }
}

Undo-PWLogin
# ---------------- fim ----------------
