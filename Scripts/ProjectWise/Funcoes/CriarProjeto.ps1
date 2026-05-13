$acessosEngenharia = @(
    'ASSISTENTE ENG',
    'CONSULTA',
    'ESPECIALISTA',
    'GESTOR ENG',
    'PROJETISTA'
)
$acessosUnidade = @(
    'CONSULTA',
    'Lider Unidade',
    'Poder Concedente',
    'GESTOR ENG'
)

# -------------------------------------------------------
# NOVA ESTRUTURA DE PASTAS
# -------------------------------------------------------
$pastasDocumentosGerais = @(
    'Atas de Reuniao',
    'Cadastro de interferências',
    'Checklist de Qualidade',
    'Cronograma',
    'EVTEA',
    'Meio ambiente',
    'Oficios',
    'Sondagens',
    'Stage Gate',
    'Topografia'
)

$fasesPadrao = @(
    '01 - Funcional',
    '02 - Anteprojeto',
    '03 - Executivo'
)

$volumesAnteprojeto = @(
    'Volume I',
    'Volume II'
)

$volumesExecutivo = @(
    'Volume I',
    'Volume II',
    'Volume III',
    'Volume IV'
)

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

function Escrever-LogParametro {
    param(
        [string]$Nome,
        [object]$Valor
    )

    Write-Host ("{0,-25}: [{1}]" -f $Nome, ([string]$Valor))
}

#-------------------------------------------------------
# Funções de criação de pastas
#-------------------------------------------------------
function DefinicaoEnvironmentEngenharia {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') { 
        return 'dmsEngenhariaARTESP'
    }
    
    if ($poderConcedente -eq 'ANTT') {
        return 'dmsEngenhariaANTT'
    }

    return ''
}

function DefinicaoEnvironmentUnidade {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') {
        return 'dmsUnidadeARTESP'
    }
    
    if ($poderConcedente -eq 'ANTT') {
        return 'dmsUnidadeANTT'
    }

    return ''
}

function DefinicaoEnvironmentGRD {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') {
        return 'dmsDocumentosGRDARTESP'
    }

    if ($poderConcedente -eq 'ANTT') {
        return 'dmsDocumentosGRDANTT'
    }

    return ''
}

function DefinicaoEnvironmentGRDRetorno {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') {
        return 'dmsGRDRetornoARTESP'
    }
    
    if ($poderConcedente -eq 'ANTT') {
        return 'dmsGRDRetornoANTT'
    }

    return ''
}

function DefinicaoEnvironmentLD {
    param ($poderConcedente)

    if ($poderConcedente -eq 'ARTESP') {
        return 'dmsListaDocumentosARTESP'
    }

    if ($poderConcedente -eq 'ANTT') {
        return 'dmsListaDocumentosANTT'
    }

    return ''
}

function CriarPastaRaizProjeto {
    param (
        [string]$NomeConcessao,
        [string]$SiglaConcessao,
        [string]$Projeto,
        [string]$Descricao,
        [string]$GestorEngenharia,
        [string]$AssistenteEngenharia,
        [string]$PoderConcedente,
        [string]$Projetista
    )

    $NomeConcessao        = Normalizar-Texto $NomeConcessao
    $SiglaConcessao       = Normalizar-Texto $SiglaConcessao
    $Projeto              = Normalizar-Texto $Projeto
    $Descricao            = Normalizar-Texto $Descricao
    $GestorEngenharia     = Normalizar-Texto $GestorEngenharia
    $AssistenteEngenharia = Normalizar-Texto $AssistenteEngenharia
    $PoderConcedente      = Normalizar-Texto $PoderConcedente
    $Projetista           = Normalizar-Texto $Projetista

    $nomePasta = "ENGENHARIA\$NomeConcessao\Projetos\$Projeto"

    Write-Host ""
    Write-Host "===== DADOS RECEBIDOS PARA CRIACAO DA PASTA RAIZ ====="
    Escrever-LogParametro -Nome "NomeConcessao"        -Valor $NomeConcessao
    Escrever-LogParametro -Nome "SiglaConcessao"       -Valor $SiglaConcessao
    Escrever-LogParametro -Nome "Projeto"              -Valor $Projeto
    Escrever-LogParametro -Nome "Descricao"            -Valor $Descricao
    Escrever-LogParametro -Nome "GestorEngenharia"     -Valor $GestorEngenharia
    Escrever-LogParametro -Nome "AssistenteEngenharia" -Valor $AssistenteEngenharia
    Escrever-LogParametro -Nome "PoderConcedente"      -Valor $PoderConcedente
    Escrever-LogParametro -Nome "Projetista"           -Valor $Projetista
    Write-Host "======================================================"
    Write-Host ""

    $null = New-PWRichProject -NewFolderPath $nomePasta -ProjectType "Projetos_" -StorageArea "Storage" -UpgradeIfExists -ProjectProperties @{
        PROJECT_Sigla_da_Concesso    = $SiglaConcessao
        PROJECT_Poder_Concedente_    = $PoderConcedente
        PROJECT_Nome_Projetista      = $Projetista
        PROJECT_Gestor_da_Engenharia = $GestorEngenharia
        PROJECT_Assistente           = $AssistenteEngenharia
        PROJECT_Projeto              = $Projeto
    }

    Start-Sleep -Milliseconds 500

    $pastaRaizProjeto = Get-PWFolders -FolderPath $nomePasta -PopulatePaths -JustOne

    if (-not $pastaRaizProjeto) {
        Write-Host "Nao foi possivel localizar a pasta raiz criada: $nomePasta"
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($Descricao)) {
        Write-Host "Aplicando descricao na pasta raiz: $Descricao"

        try {
            $null = Update-PWFolderNameProps -InputFolder $pastaRaizProjeto -NewDescription $Descricao
            Start-Sleep -Milliseconds 300

            $pastaRaizProjeto = Get-PWFolders -FolderPath $nomePasta -PopulatePaths -JustOne
            Write-Host "Descricao aplicada com sucesso."
        }
        catch {
            Write-Host "Erro ao aplicar descricao na pasta raiz: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "ATENCAO: descricao veio vazia. O ProjectWise pode manter o nome do projeto como descricao."
    }

    $pastas = Get-PWFolders -FolderPath $nomePasta -PopulatePaths -JustOne
    return $pastas
}

function CriarPasta {
    param ($nomePasta, $pastasProjeto, $environment, $workflow)
    
    $pasta = $pastasProjeto.Where({ $_.Fullpath -eq $nomePasta })

    if (-not $pasta) {
        Write-Host "Criando pasta: $nomePasta"

        if (-not [string]::IsNullOrWhiteSpace($environment) -and -not [string]::IsNullOrWhiteSpace($workflow)) {
            New-PWFolder -FolderPath $nomePasta -Environment $environment -Workflow $workflow
        }
        elseif (-not [string]::IsNullOrWhiteSpace($environment)) {
            New-PWFolder -FolderPath $nomePasta -Environment $environment
        }
        elseif (-not [string]::IsNullOrWhiteSpace($workflow)) {
            New-PWFolder -FolderPath $nomePasta -Workflow $workflow
        }
        else {
            New-PWFolder -FolderPath $nomePasta
        }
    }
    else {
        Write-Host "Pasta já estava criada: $nomePasta"
    }
}

function CriarEstruturaInicialPastasProjeto {
    param ($pastasProjeto, $concessao, $projeto, $environmentEngenharia, $environmentUnidade, $environmentLD, $environmentGRD, $environmentGRDRetorno)

    $raizProjeto = "ENGENHARIA\$concessao\Projetos\$projeto"

    # 0 - Documentos Gerais
    $nomePasta = "$raizProjeto\0 - Documentos Gerais"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"

    foreach ($subpasta in $pastasDocumentosGerais) {
        $nomeSubpasta = "$raizProjeto\0 - Documentos Gerais\$subpasta"
        CriarPasta -NomePasta $nomeSubpasta -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"
    }

    # 1 - Area de Trabalho
    $nomePasta = "$raizProjeto\1 - Area de Trabalho"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"

    foreach ($fase in $fasesPadrao) {
        $nomeFase = "$raizProjeto\1 - Area de Trabalho\$fase"
        CriarPasta -NomePasta $nomeFase -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"

        if ($fase -eq '02 - Anteprojeto') {
            foreach ($volume in $volumesAnteprojeto) {
                $nomeVolume = "$nomeFase\$volume"
                CriarPasta -NomePasta $nomeVolume -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"
            }
        }

        if ($fase -eq '03 - Executivo') {
            foreach ($volume in $volumesExecutivo) {
                $nomeVolume = "$nomeFase\$volume"
                CriarPasta -NomePasta $nomeVolume -PastasProjeto $pastasProjeto -Environment $environmentEngenharia -Workflow "Workflow - Engenharia"
            }
        }
    }

    # 2 - Unidade
    $nomePasta = "$raizProjeto\2 - Unidade"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentUnidade -Workflow "Workflow - Engenharia - Unidade"

    foreach ($fase in $fasesPadrao) {
        $nomeFase = "$raizProjeto\2 - Unidade\$fase"
        CriarPasta -NomePasta $nomeFase -PastasProjeto $pastasProjeto -Environment $environmentUnidade -Workflow "Workflow - Engenharia - Unidade"

        if ($fase -eq '02 - Anteprojeto') {
            foreach ($volume in $volumesAnteprojeto) {
                $nomeVolume = "$nomeFase\$volume"
                CriarPasta -NomePasta $nomeVolume -PastasProjeto $pastasProjeto -Environment $environmentUnidade -Workflow "Workflow - Engenharia - Unidade"
            }
        }

        if ($fase -eq '03 - Executivo') {
            foreach ($volume in $volumesExecutivo) {
                $nomeVolume = "$nomeFase\$volume"
                CriarPasta -NomePasta $nomeVolume -PastasProjeto $pastasProjeto -Environment $environmentUnidade -Workflow "Workflow - Engenharia - Unidade"
            }
        }
    }

    # Area de Transferencia
    $nomePasta = "$raizProjeto\Area de Transferencia"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Workflow "Workflow - Transferencia"

    $nomePasta = "$raizProjeto\Area de Transferencia\Entrada"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentGRD -Workflow "Workflow - Transferencia"

    $nomePasta = "$raizProjeto\Area de Transferencia\Saida"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentGRDRetorno -Workflow "Workflow - Transferencia"

    $nomePasta = "$raizProjeto\Area de Transferencia\Importacao"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentGRD -Workflow "Importacao"

    # Previsao de Documentos (LD)
    $nomePasta = "$raizProjeto\Previsao de Documentos (LD)"
    CriarPasta -NomePasta $nomePasta -PastasProjeto $pastasProjeto -Environment $environmentLD
}

function CriarPastasArea {
    param ($pastasProjeto, $poderConcedente, $concessao, $projeto, $area, $environment, $workflow, $workflowSuperados)

    # Mantida apenas para compatibilidade.
    # Na nova estrutura, as subpastas já são criadas em CriarEstruturaInicialPastasProjeto.
    return
}

function CriarPastasProjeto {
    param (
        [string]$NomeConcessao,
        [string]$SiglaConcessao,
        [string]$Projeto,
        [string]$Descricao,
        [string]$PoderConcedente,
        [string]$Projetista,
        [string]$GestorEngenharia,
        [string]$AssistenteEngenharia
    )

    $NomeConcessao        = Normalizar-Texto $NomeConcessao
    $SiglaConcessao       = Normalizar-Texto $SiglaConcessao
    $Projeto              = Normalizar-Texto $Projeto
    $Descricao            = Normalizar-Texto $Descricao
    $PoderConcedente      = Normalizar-Texto $PoderConcedente
    $Projetista           = Normalizar-Texto $Projetista
    $GestorEngenharia     = Normalizar-Texto $GestorEngenharia
    $AssistenteEngenharia = Normalizar-Texto $AssistenteEngenharia

    $environmentLD         = DefinicaoEnvironmentLD -PoderConcedente $PoderConcedente
    $environmentGRD        = DefinicaoEnvironmentGRD -PoderConcedente $PoderConcedente
    $environmentUnidade    = DefinicaoEnvironmentUnidade -PoderConcedente $PoderConcedente
    $environmentEngenharia = DefinicaoEnvironmentEngenharia -PoderConcedente $PoderConcedente
    $environmentGRDRetorno = DefinicaoEnvironmentGRDRetorno -PoderConcedente $PoderConcedente

    if (-not $environmentEngenharia -or -not $environmentUnidade -or -not $environmentGRD -or -not $environmentLD -or -not $environmentGRDRetorno) {
        Write-Host "Nao foi possivel definir os environments para o poder concedente informado: $PoderConcedente"
        return $null
    }   
    
    $pastaRaiz = CriarPastaRaizProjeto `
        -NomeConcessao $NomeConcessao `
        -SiglaConcessao $SiglaConcessao `
        -Projeto $Projeto `
        -Descricao $Descricao `
        -PoderConcedente $PoderConcedente `
        -Projetista $Projetista `
        -GestorEngenharia $GestorEngenharia `
        -AssistenteEngenharia $AssistenteEngenharia

    if (-not $pastaRaiz) {
        return $null
    }
    
    $null = CriarEstruturaInicialPastasProjeto `
        -PastasProjeto $pastaRaiz `
        -Concessao $NomeConcessao `
        -Projeto $Projeto `
        -EnvironmentEngenharia $environmentEngenharia `
        -EnvironmentUnidade $environmentUnidade `
        -EnvironmentGRD $environmentGRD `
        -EnvironmentLD $environmentLD `
        -EnvironmentGRDRetorno $environmentGRDRetorno

    # Mantido por compatibilidade, mas sem ação
    $null = CriarPastasArea `
        -PastasProjeto $pastaRaiz `
        -PoderConcedente $PoderConcedente `
        -Concessao $NomeConcessao `
        -Projeto $Projeto `
        -Area "1 - Area de Trabalho" `
        -Environment $environmentEngenharia `
        -Workflow "Workflow - Engenharia" `
        -WorkflowSuperados "Workflow - Engenharia - Superados"

    $null = CriarPastasArea `
        -PastasProjeto $pastaRaiz `
        -PoderConcedente $PoderConcedente `
        -Concessao $NomeConcessao `
        -Projeto $Projeto `
        -Area "2 - Unidade" `
        -Environment $environmentUnidade `
        -Workflow "Workflow - Engenharia - Unidade" `
        -WorkflowSuperados "Workflow - Engenharia - Superados"

    return $pastaRaiz
}

#-------------------------------------------------------
# Funções para definições de acesso
#-------------------------------------------------------
function DefineAcessosProjeto {
    param ($nomeConcessao, $siglaConcessao, $projeto)
    
    DefinirAcessoRaizProjeto -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto
    DefinirAcessoEngenharia -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto
    DefinirAcessoUnidade -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto
    DefinirAcessoLD -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto
    DefinirAcessoGRD -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto
}

function DefinirAcessoRaizProjeto {
    param ($nomeConcessao, $siglaConcessao, $projeto)
    
    $nomePastaRaiz = "ENGENHARIA\$nomeConcessao\Projetos\$projeto"
    $nomeUserList = "$siglaConcessao-$projeto"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "Engenharia GPR" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$siglaConcessao-ENGENHARIA UNIDADE" -MemberAccess "r" 

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Everyone -MemberAccess "na" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType UserList -MemberName $nomeUserList -MemberAccess "r" 

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-ASSISTENTE ENG" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-ENG CONSULTA TODAS REV" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-ENG CONSULTA ULTIMA REV" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-ESPECIALISTA" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-GESTOR ENG" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-GESTOR UNIDADE" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-PROJETISTA" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-CONSULTA TODAS REV" -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$nomeUserList-CONSULTA ULTIMA REV" -MemberAccess "r" 

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType Everyone -MemberAccess "na" 
}

function DefinirAcessoEngenharia {
    param ($nomeConcessao, $siglaConcessao, $projeto)

    $pastaEngenharia = "ENGENHARIA\$nomeConcessao\Projetos\$projeto\1 - Area de Trabalho"
    $nomeUserListEngenharia = "$siglaConcessao-$projeto-ENGENHARIA"
    
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -FolderSecurity -MemberType Everyone -MemberAccess "na" 
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -FolderSecurity -MemberType UserList -MemberName $nomeUserListEngenharia -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -MemberType User -MemberName "admin" -MemberAccess "fc" 

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListEngenharia-ESPECIALISTA" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListEngenharia-ASSISTENTE ENG" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em elaboracao" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise do Assistente" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise do Assistente" -MemberType UserList -MemberName "$nomeUserListEngenharia-ASSISTENTE ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise do Assistente" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise do Assistente" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise da Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise da Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise da Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise Especifica" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise Especifica" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise Especifica" -MemberType UserList -MemberName "$nomeUserListEngenharia-ESPECIALISTA" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em analise Especifica" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em consolidacao pela Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em consolidacao pela Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em consolidacao pela Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-PROJETISTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Projetista" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em revisao pelo Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em revisao pelo Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em revisao pelo Projetista" -MemberType UserList -MemberName "$nomeUserListEngenharia-PROJETISTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em revisao pelo Projetista" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Ressalvas" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Pendencias do Especialista" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Pendencias do Especialista" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em Envio para Unidade c/ Pendencias do Especialista" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade com Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade com Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade com Ressalvas" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade e Projetista com Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade e Projetista com Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade e Projetista com Ressalvas" -MemberType UserList -MemberName "$nomeUserListEngenharia-PROJETISTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade e Projetista com Ressalvas" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias do Especialista" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias do Especialista" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias do Especialista" -MemberType UserList -MemberName "$nomeUserListEngenharia-ESPECIALISTA" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias do Especialista" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias da Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias da Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Enviado para Unidade c/ Pendencias da Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Devolvido pela Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Devolvido pela Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Devolvido pela Unidade" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em devolucao para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em devolucao para Unidade" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Em devolucao para Unidade" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Arquivo interno Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Arquivo interno Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Arquivo interno Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Concluido pela Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Concluido pela Engenharia" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Concluido pela Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Superado" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia" -StateName "Superado" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType UserList -MemberName "$nomeUserListEngenharia-GESTOR ENG" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType UserList -MemberName "$nomeUserListEngenharia-ESPECIALISTA" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Superados" -StateName "Superado" -MemberType UserList -MemberName "$nomeUserListEngenharia-CONSULTA-TODAS REV" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaEngenharia -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Superados" -StateName "Superado" -MemberType Everyone -MemberAccess "na"
}

function DefinirAcessoUnidade {
    param ($nomeConcessao, $siglaConcessao, $projeto)
 
    $pastaUnidade = "ENGENHARIA\$nomeConcessao\Projetos\$projeto\2 - Unidade"
    $nomeUserListUnidade = "$siglaConcessao-$projeto-UNIDADE"
    
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -FolderSecurity -MemberType Everyone -MemberAccess "na" 
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -FolderSecurity -MemberType UserList -MemberName $nomeUserListUnidade -MemberAccess "r" 
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -MemberType User -MemberName "admin" -MemberAccess "fc" 

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Em elaboracao" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Em elaboracao" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Emitido pela Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Emitido pela Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Emitido pela Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Solicitado reanalise da Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Solicitado reanalise da Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Solicitado reanalise da Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-GESTOR ENG" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Solicitado reanalise da Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente" -MemberType UserList -MemberName "$nomeUserListUnidade-Poder Concedente" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido pela Unidade" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido pela Unidade" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido pela Unidade" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Nova emissao sendo analisada pela Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Nova emissao sendo analisada pela Engenharia" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Nova emissao sendo analisada pela Engenharia" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente - Nova emissao em analise Eng" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente - Nova emissao em analise Eng" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente - Nova emissao em analise Eng" -MemberType UserList -MemberName "$nomeUserListUnidade-Poder Concedente" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Enviado ao Poder Concedente - Nova emissao em analise Eng" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido - Nova emissao em analise Eng" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido - Nova emissao em analise Eng" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "r", "w", "cw", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Concluido - Nova emissao em analise Eng" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Superado" -MemberType UserList -MemberName "$nomeUserListUnidade-CONSULTA-TODAS REV" -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Engenharia - Unidade" -StateName "Superado" -MemberType Everyone -MemberAccess "na"

    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType UserList -MemberName "$nomeUserListUnidade-Lider Unidade" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType UserList -MemberName "$nomeUserListUnidade-GESTOR ENG" -MemberAccess "c", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastaUnidade -DocumentSecurity -WorkFlowName "Workflow - Comentarios" -StateName "Markup" -MemberType Everyone -MemberAccess "na"
}

function DefinirAcessoLD {
    param ($nomeConcessao, $siglaConcessao, $projeto)
    
    $nomePastaRaiz = "ENGENHARIA\$nomeConcessao\Projetos\$projeto\Previsao de Documentos (LD)"
    $nomeUserList = "$siglaConcessao-$projeto-LD"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "Engenharia GPR" -MemberAccess "fc"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$siglaConcessao-ENGENHARIA UNIDADE" -MemberAccess "fc"
    
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Everyone -MemberAccess "na"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType UserList -MemberName $nomeUserList -MemberAccess "r"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType UserList -MemberName "$nomeUserList-EDICAO" -MemberAccess "r", "c", "w"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType Everyone -MemberAccess "na"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType UserList -MemberName $nomeUserList -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType UserList -MemberName "$nomeUserList-EDICAO" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
}

function DefinirAcessoGRD {
    param ($nomeConcessao, $siglaConcessao, $projeto)
    
    $nomePastaRaiz = "ENGENHARIA\$nomeConcessao\Projetos\$projeto\Area de Transferencia"
    $nomeUserList = "$siglaConcessao-$projeto-GRD"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "Engenharia GPR" -MemberAccess "fc"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Group -MemberName "$siglaConcessao-ENGENHARIA UNIDADE" -MemberAccess "fc"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType Everyone -MemberAccess "na"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType UserList -MemberName $nomeUserList -MemberAccess "r"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -FolderSecurity -MemberType UserList -MemberName "$nomeUserList-EDICAO" -MemberAccess "r", "c", "w"

    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType Everyone -MemberAccess "na"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType UserList -MemberName $nomeUserList -MemberAccess "r", "fr"
    Update-PWFolderSecurity -Verbose -InputFolder $nomePastaRaiz -DocumentSecurity -MemberType UserList -MemberName "$nomeUserList-EDICAO" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
}

function DefinirAcessoXFDF {
    param ($nomeConcessao, $siglaConcessao, $projeto)
    
    $nomePastaRaiz = "ENGENHARIA\$nomeConcessao\Projetos\$projeto"
    
    $pastasXFDF = Get-PWFolders -FolderPath $nomePastaRaiz
    $pastasXFDF = $pastasXFDF | Where-Object { $_.Name.ToLower() -eq 'xfdf__' }

    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-ASSISTENTE ENG" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-CONSULTA TODAS REV" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-CONSULTA ULTIMA REV" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-ENG CONSULTA TODAS REV" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-ENG CONSULTA ULTIMA REV" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-ESPECIALISTA" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-GESTOR ENG" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-GESTOR UNIDADE" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
    Update-PWFolderSecurity -Verbose -InputFolder $pastasXFDF -DocumentSecurity -MemberType Group -MemberName "$siglaConcessao-$projeto-PROJETISTA" -MemberAccess "c", "d", "r", "w", "cw", "fr", "fw"
}

#-------------------------------------------------------
# Funções para criação de User Lists
#-------------------------------------------------------
function CriarUserLists {
    param ($siglaConcessao, $projeto)

    $userListConcessao = CriarUserListConcessao -SiglaConcessao $siglaConcessao
    $userListProjeto = CriarUserListProjeto -UserListConcessao $userListConcessao -Projeto $projeto
    $userListArea = CriarUserListArea -UserListProjeto $userListProjeto -NomeArea 'ENGENHARIA' -Acessos $acessosEngenharia
    $userListArea = CriarUserListArea -UserListProjeto $userListProjeto -NomeArea 'UNIDADE' -Acessos $acessosUnidade
    CriarUserListLD -UserListProjeto $userListProjeto
    CriarUserListGRD -UserListProjeto $userListProjeto
}

function CriarUserListConcessao {
    param ($siglaConcessao)
    
    $userlist = CriarUserList -Nome $siglaConcessao
    return $userlist
}

function CriarUserListProjeto {
    param ($userListConcessao, $projeto)

    if (-not $userListConcessao) {
        return $null
    }

    $nomeUserList = $userListConcessao.Name + "-$projeto"
    $userlist = CriarUserList -Nome $nomeUserList -UserListPai $userListConcessao

    return $userlist
}

function CriarUserListArea {
    param ($userListProjeto, $nomeArea, $acessos)
    
    if (-not $userListProjeto) {
        return $null
    }

    $nomeUserListArea = $userListProjeto.Name + "-$nomeArea"
    $userListArea = CriarUserList -Nome $nomeUserListArea -UserListPai $userListProjeto
    
    foreach ($acesso in $acessos) {
        $nome = $nomeUserListArea + "-$acesso"
        $userlist = CriarUserList -Nome $nome -UserListPai $userListArea
        if ($acesso -eq 'CONSULTA') {
            $nomeTodasRev = $nome + "-TODAS REV"
            CriarUserList -Nome $nomeTodasRev -UserListPai $userlist 
        }
    }

    return $userListArea
}

function CriarUserListLD {
    param ($userListProjeto)

    if (-not $userListProjeto) {
        return
    }

    $nomeUL = $userListProjeto.Name + "-LD"
    $ulLD = CriarUserList -Nome $nomeUL -UserListPai $userListProjeto

    $nomeUL = $userListProjeto.Name + "-LD-CONSULTA"
    CriarUserList -Nome $nomeUL -UserListPai $ulLD

    $nomeUL = $userListProjeto.Name + "-LD-EDICAO"
    CriarUserList -Nome $nomeUL -UserListPai $ulLD

    return $ulLD
}

function CriarUserListGRD {
    param ($userListProjeto)

    if (-not $userListProjeto) {
        return
    }

    $nomeUL = $userListProjeto.Name + "-GRD"
    $ulGRD = CriarUserList -Nome $nomeUL -UserListPai $userListProjeto

    $nomeUL = $userListProjeto.Name + "-GRD-CONSULTA"
    CriarUserList -Nome $nomeUL -UserListPai $ulGRD

    $nomeUL = $userListProjeto.Name + "-GRD-EDICAO"
    CriarUserList -Nome $nomeUL -UserListPai $ulGRD

    return $ulGRD
}

function CriarUserList {
    param ($nome, $userListPai)

    $userlist = Get-PWUserLists -UserListName $nome
    $userlist = $userlist.Where({$_.Name -eq $nome})

    if (-not $userlist) {
        $userlist = New-PWUserListByName -UserList $nome
    }

    if ($userListPai) {
        Add-PWMemberToUserList -UserList $userListPai.Name -UserListNames $nome
    }

    return $userlist
}

#-------------------------------------------------------
# Funções para criação de Grupos de Usuários
#-------------------------------------------------------
function CriarGruposUsuarios {
    param ($siglaConcessao, $projeto)

    CriarGrupoEngenhariaGPR                   -siglaConcessao $siglaConcessao -projeto $projeto
    CriarGrupoEngenhariaUnidade               -SiglaConcessao $siglaConcessao -projeto $projeto
    CriarGrupoAssistenteEngenharia            -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoGestorEngenharia                -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoEspecialista                    -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoProjetista                      -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoConsultaUltimaRevisaoEngenharia -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoConsultaTodasRevisoesEngenharia -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoConsultaUltimaRevisaoUnidade    -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoConsultaTodasRevisoesUnidade    -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoGestorUnidade                   -SiglaConcessao $siglaConcessao -Projeto $projeto
    CriarGrupoPoderConcedente                 -SiglaConcessao $siglaConcessao -Projeto $projeto
}

function CriarGrupoEngenhariaGPR {
    param ($siglaConcessao, $projeto)

    $nomeGrupo = "Engenharia GPR"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-GESTOR ENG",
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-GESTOR ENG",
        "$siglaConcessao-$projeto-GRD-CONSULTA",
        "$siglaConcessao-$projeto-LD-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoPoderConcedente {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-PODER CONCEDENTE"

    $userLists = @(
        "$siglaConcessao-$projeto-UNIDADE-Poder Concedente"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoEngenhariaUnidade {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-ENGENHARIA UNIDADE"

    $userLists = @(
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-Lider Unidade"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -userListsMembros $userLists
    return $grupo
}

function CriarGrupoAssistenteEngenharia {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-ASSISTENTE ENG"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-ASSISTENTE ENG",
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-Lider Unidade",
        "$siglaConcessao-$projeto-GRD-EDICAO",
        "$siglaConcessao-$projeto-LD-EDICAO"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoGestorEngenharia {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-GESTOR ENG"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-GESTOR ENG",
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-GESTOR ENG",
        "$siglaConcessao-$projeto-GRD-CONSULTA",
        "$siglaConcessao-$projeto-LD-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoEspecialista {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-ESPECIALISTA"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-ESPECIALISTA",
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-GRD-CONSULTA",
        "$siglaConcessao-$projeto-LD-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoProjetista {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-PROJETISTA"

    $userLists = @(
        "$siglaConcessao-$projeto-GRD-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoConsultaUltimaRevisaoEngenharia {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-ENG CONSULTA ULTIMA REV"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoConsultaTodasRevisoesEngenharia {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-ENG CONSULTA TODAS REV"

    $userLists = @(
        "$siglaConcessao-$projeto-ENGENHARIA-CONSULTA-TODAS REV"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoConsultaUltimaRevisaoUnidade {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-CONSULTA ULTIMA REV"

    $userLists = @(
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -Descricao '(OBRA, PREFEITURA E ETC)' -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoConsultaTodasRevisoesUnidade {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-CONSULTA TODAS REV"

    $userLists = @(
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -Descricao '(OBRA, PREFEITURA E ETC)' -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoGestorUnidade {
    param ($siglaConcessao, $projeto)
    
    $nomeGrupo = "$siglaConcessao-$projeto-GESTOR UNIDADE"

    $userLists = @(
        "$siglaConcessao-$projeto-UNIDADE-CONSULTA-TODAS REV",
        "$siglaConcessao-$projeto-UNIDADE-Lider Unidade"
    )

    $grupo = CriarGrupoUsuario -Nome $nomeGrupo -UserListsMembros $userLists
    return $grupo
}

function CriarGrupoUsuario {
    param ($nome, $descricao, $userListsMembros)

    if (-not $descricao) { $descricao = $nome }

    $grupo = Get-PWGroups -GroupName $nome

    if (-not $grupo) {
        $grupo = New-PWGroupByName -GroupName $nome -GroupDescription $descricao
    }

    foreach ($ul in $userListsMembros) {
        Add-PWMemberToUserList -UserList $ul -GroupNames $nome
    }
}

function CriarProjeto {
    param (
        [string]$PoderConcedente,
        [string]$NomeConcessao,
        [string]$SiglaConcessao,
        [string]$Projeto,
        [string]$Descricao,
        [string]$Projetista,
        [string]$GestorEngenharia,
        [string]$AssistenteEngenharia
    )

    $PoderConcedente      = Normalizar-Texto $PoderConcedente
    $NomeConcessao        = Normalizar-Texto $NomeConcessao
    $SiglaConcessao       = Normalizar-Texto $SiglaConcessao
    $Projeto              = Normalizar-Texto $Projeto
    $Descricao            = Normalizar-Texto $Descricao
    $Projetista           = Normalizar-Texto $Projetista
    $GestorEngenharia     = Normalizar-Texto $GestorEngenharia
    $AssistenteEngenharia = Normalizar-Texto $AssistenteEngenharia

    Write-Host ""
    Write-Host "===== INICIO DA CRIACAO DO PROJETO ====="
    Escrever-LogParametro -Nome "PoderConcedente"      -Valor $PoderConcedente
    Escrever-LogParametro -Nome "NomeConcessao"        -Valor $NomeConcessao
    Escrever-LogParametro -Nome "SiglaConcessao"       -Valor $SiglaConcessao
    Escrever-LogParametro -Nome "Projeto"              -Valor $Projeto
    Escrever-LogParametro -Nome "Descricao"            -Valor $Descricao
    Escrever-LogParametro -Nome "Projetista"           -Valor $Projetista
    Escrever-LogParametro -Nome "GestorEngenharia"     -Valor $GestorEngenharia
    Escrever-LogParametro -Nome "AssistenteEngenharia" -Valor $AssistenteEngenharia
    Write-Host "========================================"
    Write-Host ""

    $pastaRaiz = CriarPastasProjeto `
        -NomeConcessao $NomeConcessao `
        -SiglaConcessao $SiglaConcessao `
        -Projeto $Projeto `
        -Descricao $Descricao `
        -PoderConcedente $PoderConcedente `
        -Projetista $Projetista `
        -GestorEngenharia $GestorEngenharia `
        -AssistenteEngenharia $AssistenteEngenharia

    if (-not $pastaRaiz) {
        return $null
    }
    
    $null = CriarUserLists       -SiglaConcessao $SiglaConcessao -Projeto $Projeto
    $null = CriarGruposUsuarios  -SiglaConcessao $SiglaConcessao -Projeto $Projeto
    $null = DefineAcessosProjeto -SiglaConcessao $SiglaConcessao -Projeto $Projeto -NomeConcessao $NomeConcessao   

    return $pastaRaiz
}