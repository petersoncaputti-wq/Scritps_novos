<#
.SYNOPSIS
Concede acesso ARTESP no projeto, Unidade e documentos selecionados por workflow/estado.
.DESCRIPTION
Seleciona concessao, projeto, Unidade, disciplina, workflow/estado e documentos.
Projeto e Unidade: Read, somente na pasta selecionada. Documentos: Read + File Read
no workflow/estado atual, preservando os demais participantes e direitos.
Planeja todo o lote antes de confirmar, valida cada documento apos aplicar e
interrompe na primeira falha. Consulta apenas documentos atuais da pasta selecionada.
.EXAMPLE
powershell.exe -NoProfile -MTA -File ".\18 - Preparar Acesso ARTESP.ps1"
#>
[CmdletBinding()]
param(
    [string]$DatasourceName,
    [string]$PastaSaida
)
$ErrorActionPreference = 'Stop'
$caminhoScriptARTESP = $PSCommandPath
if ([string]::IsNullOrWhiteSpace($caminhoScriptARTESP)) {
    $caminhoScriptARTESP = $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($caminhoScriptARTESP)) {
    throw 'Nao foi possivel identificar o script. Execute o arquivo .ps1 salvo em disco.'
}
if ([string]::IsNullOrWhiteSpace($PastaSaida)) {
    $PastaSaida = Join-Path (Split-Path -LiteralPath $caminhoScriptARTESP) 'RelatoriosAcessoARTESP'
}
if ($PSVersionTable.PSEdition -ne 'Desktop' -or -not [Environment]::Is64BitProcess -or
    [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'MTA') {
    $comandoARTESP = "& '" + $caminhoScriptARTESP.Replace("'", "''") + "'"
    foreach ($parametroARTESP in @('DatasourceName', 'PastaSaida')) {
        $valorARTESP = [string](Get-Variable -Name $parametroARTESP -ValueOnly)
        if (-not [string]::IsNullOrWhiteSpace($valorARTESP)) {
            $comandoARTESP += " -$parametroARTESP '" + $valorARTESP.Replace("'", "''") + "'"
        }
    }
    # Mantem o diagnostico visivel inclusive quando o processo filho falha.
    $comandoARTESP = "try { $comandoARTESP } catch { Write-Host (`$_ | Out-String) -ForegroundColor Red; Read-Host 'Pressione Enter para encerrar'; exit 1 }; Read-Host 'Execucao encerrada. Pressione Enter para fechar'"
    $codigoARTESP = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($comandoARTESP))
    $executavelARTESP = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        $executavelARTESP = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    }
    Write-Host 'Iniciando Windows PowerShell 64 bits em MTA...'
    & $executavelARTESP -NoProfile -MTA -EncodedCommand $codigoARTESP
    if ($LASTEXITCODE -ne 0) { throw 'A execucao em Windows PowerShell MTA falhou.' }
    return
}

function Texto($Objeto, [string[]]$Nomes) {
    foreach ($nome in $Nomes) {
        $p = $Objeto.PSObject.Properties[$nome]
        if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) { return [string]$p.Value }
    }
    return ''
}
function Id-Pasta($Pasta) {
    foreach ($campo in @('ProjectID','FolderID','ID')) {
        $id = 0
        if ([int]::TryParse((Texto $Pasta @($campo)), [ref]$id) -and $id -gt 0) { return $id }
    }
    throw 'Pasta sem identificador valido.'
}
function Nome-Pasta($Pasta) { Texto $Pasta @('Name','FolderName','ProjectName','ObjectName') }
function Filhos($Pasta) {
    $avisos = @()
    $resultado = @(Get-PWFoldersImmediateChildren -FolderID (Id-Pasta $Pasta) -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable avisos)
    foreach ($aviso in $avisos) {
        if ([string]$aviso -notmatch '^\s*No immediate child folders found\.?\s*$') { throw "Falha ao listar subpastas: $aviso" }
    }
    @($resultado | Where-Object { $null -ne $_ -and $_ -isnot [bool] } | ForEach-Object {
        $null = Id-Pasta $_
        $_
    } | Sort-Object { Nome-Pasta $_ })
}
function Escolher([object[]]$Itens, [string]$Titulo, [switch]$Varios) {
    $Itens = @($Itens | Where-Object { $null -ne $_ -and $_ -isnot [bool] })
    if ($Itens.Count -eq 0) { throw "Nenhum item: $Titulo" }
    Write-Host "`n$Titulo" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Itens.Count; $i++) {
        $rotulo = Texto $Itens[$i] @('Rotulo','Description','Name','FolderName','ProjectName')
        Write-Host ('{0} - {1}' -f ($i + 1), $rotulo)
    }
    while ($true) {
        $entrada = Read-Host 'Informe numero(s) separados por virgula; Q cancela'
        if ($entrada -eq 'Q') { throw 'Operacao cancelada pelo usuario.' }
        $partes = @($entrada.Split(',') | ForEach-Object { $_.Trim() })
        $indices = @(); $valido = $true
        foreach ($parte in $partes) {
            $numero = 0
            if (-not [int]::TryParse($parte, [ref]$numero) -or $numero -lt 1 -or $numero -gt $Itens.Count) {
                $valido = $false; break
            }
            $indices += $numero - 1
        }
        if (-not $Varios -and $indices.Count -ne 1) { $valido = $false }
        if ($valido) {
            foreach ($indice in @($indices | Select-Object -Unique)) { $Itens[$indice] }
            return
        }
        Write-Host 'Selecao invalida.' -ForegroundColor Yellow
    }
}
function Localizar([object[]]$Pastas, [string]$Nome) {
    $encontradas = @($Pastas | Where-Object { (Nome-Pasta $_) -eq $Nome })
    if ($encontradas.Count -eq 1) { return $encontradas[0] }
    Escolher $Pastas "Selecione a pasta $Nome"
}
function Exportar-Seguranca($Pasta, [string]$Nome) {
    # Falhas sao registradas e interrompem a analise; nao equivalem a ausencia de ACL.
    $dados = Get-PWFolderSecurity -InputFolder $Pasta -ErrorAction Stop
    if ($dados -is [System.Data.DataTable]) { $dados = $dados.Rows }
    $dados | Export-Clixml -LiteralPath (Join-Path $script:Destino "$Nome.xml") -Depth 8
    $dados | Export-Csv -LiteralPath (Join-Path $script:Destino "$Nome.csv") -NoTypeInformation -Encoding UTF8
    Write-Host "Seguranca exportada: $Nome"
}
function Testar-SessaoPW {
    try {
        return [bool](Get-PWLoginStatus -ErrorAction Stop)
    }
    catch {
        # PWPS_DAB 24.0.2 sinaliza ausencia de login por excecao.
        if ($_.Exception.Message -match 'Not logged into a ProjectWise datasource') {
            return $false
        }
        throw
    }
}
function Conferir-ARTESPPasta($Projeto, [string]$Etapa) {
    if (@($Projeto).Count -ne 1) { throw 'Esta etapa exige exatamente uma pasta, sem subpastas.' }
    Exportar-Seguranca $Projeto ("Seguranca${Etapa}Antes")
    $dados = Get-PWFolderSecurity -InputFolder $Projeto -ErrorAction Stop
    if ($dados -is [System.Data.DataTable]) { $dados = $dados.Rows }
    $entradas = @($dados | Where-Object {
        $_.Name -eq 'ARTESP' -and $_.Type -eq 'Group' -and
        $_.SecurityType -match '^(Project|Folder|Project/Folder)$' -and
        [string]::IsNullOrWhiteSpace([string]$_.Workflow) -and
        [string]::IsNullOrWhiteSpace([string]$_.State)
    })
    Write-Host ('{0}: {1} (ID {2})' -f $Etapa,(Nome-Pasta $Projeto),(Id-Pasta $Projeto)) -ForegroundColor Cyan
    if ($entradas.Count -gt 0) {
        $entradas | Format-Table Name,Type,SecurityType,Access_Control_Settings,InheritingFrom -AutoSize | Out-Host
        if (@($entradas | Where-Object { $_.Access_Control_Settings -ne '----r------' }).Count -gt 0) {
            throw 'ARTESP possui uma configuracao diferente de Read. Revise o relatorio antes de substituir direitos existentes.'
        }
        Write-Host 'ARTESP ja possui Read. Nenhuma alteracao necessaria nesta pasta.'
        return
    }
    Write-Host 'Alteracao proposta: adicionar grupo ARTESP com Read na seguranca da pasta selecionada.' -ForegroundColor Yellow
    Write-Host 'Escopo: somente esta pasta. Subpastas podem herdar; suas permissoes proprias nao serao substituidas.' -ForegroundColor Yellow
    if ((Read-Host 'Digite APLICAR para executar esta etapa; outra resposta encerra').Trim() -ine 'APLICAR') {
        throw 'Etapa da pasta cancelada pelo usuario.'
    }
    # Equivalente ao escopo "Apply changes to this folder only": somente
    # o objeto selecionado e fornecido, sem enumerar ou regravar descendentes.
    Update-PWFolderSecurity -InputFolder $Projeto -FolderSecurity -MemberType Group -MemberName 'ARTESP' -MemberAccess 'r' -ErrorAction Stop | Out-Host
    Exportar-Seguranca $Projeto ("Seguranca${Etapa}Depois")
    $depois = Get-PWFolderSecurity -InputFolder $Projeto -ErrorAction Stop
    if ($depois -is [System.Data.DataTable]) { $depois = $depois.Rows }
    $confirmadas = @($depois | Where-Object {
        $_.Name -eq 'ARTESP' -and $_.Type -eq 'Group' -and
        $_.SecurityType -match '^(Project|Folder|Project/Folder)$' -and
        [string]::IsNullOrWhiteSpace([string]$_.Workflow) -and
        [string]::IsNullOrWhiteSpace([string]$_.State) -and
        $_.Access_Control_Settings -eq '----r------'
    })
    if ($confirmadas.Count -eq 0) { throw 'Read nao confirmado apos a aplicacao. Confira os relatorios e o ProjectWise.' }
    Write-Host 'Read de ARTESP confirmado nesta pasta.' -ForegroundColor Green
}
function Testar-AcessoDocumento([object[]]$Linhas) {
    . (Join-Path $PSScriptRoot 'Funcoes\PreservarSegurancaARTESP.ps1')
    $linha = Escolher $Linhas 'Selecione UM documento para testar preservacao completa de acessos'
    if ([string]::IsNullOrWhiteSpace($linha.DocumentGUID)) { throw 'Documento sem GUID.' }
    $pastaTeste = Join-Path $Destino ('TestePreservacao_' + [guid]::NewGuid().ToString('N').Substring(0,6))
    $recuperar = $linha.DocumentGUID -eq '5fc0819d-db6f-4841-a979-75a387bb1d01'
    Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino $pastaTeste -RecuperarOrigem:$recuperar
    Write-Host 'O plano acima preserva os participantes e direitos anteriores e acrescenta ARTESP. A heranca sera materializada no documento quando necessario.' -ForegroundColor Yellow
    if ((Read-Host 'Digite TESTAR UM para aplicar somente neste arquivo; Enter encerra').Trim() -ine 'TESTAR UM') { return }
    Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino (Join-Path $pastaTeste 'Aplicacao') -RecuperarOrigem:$recuperar -Aplicar
    Write-Host 'Teste concluido. Confira a lista completa em Security > Workflow no ProjectWise.' -ForegroundColor Green
}

$loginCriado = $false
try {
    . (Join-Path (Split-Path -LiteralPath $caminhoScriptARTESP) 'Funcoes\PreservarSegurancaARTESP.ps1')
    Import-Module PWPS_DAB -DisableNameChecking -ErrorAction Stop
    if (-not (Testar-SessaoPW)) {
        $login = @{ UseGui = $true; DoNotCreateWorkingDirectory = $true; ErrorAction = 'Stop' }
        if ($DatasourceName) { $login.DatasourceName = $DatasourceName }
        if (-not (New-PWLogin @login)) { throw 'Login cancelado ou nao concluido.' }
        $loginCriado = $true
    }
    elseif ($DatasourceName) {
        throw 'Ja existe uma sessao: execute sem DatasourceName ou encerre a sessao antes de escolher outro datasource.'
    }
    Write-Host 'ARTESP: Read no projeto e Unidade; Read + File Read nos documentos. Cada aplicacao exige confirmacao.' -ForegroundColor Yellow
    Get-PWCurrentDatasource | Format-List | Out-Host
    $raizes = @(Get-PWFoldersImmediateChildren -Root -ErrorAction Stop)
    $engenharia = Localizar $raizes 'Engenharia'
    $concessao = Escolher (Filhos $engenharia) 'Selecione a concessao'
    $script:Destino = Join-Path $PastaSaida ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + [guid]::NewGuid().ToString('N').Substring(0,6))
    New-Item -ItemType Directory -Path $script:Destino -Force | Out-Null
    $concessao | Export-Clixml (Join-Path $Destino 'ConcessaoSelecionada.xml') -Depth 5
    Write-Host ('Concessao selecionada: {0} | Nome: {1} | ID: {2}' -f (Texto $concessao @('Description','Name')), (Nome-Pasta $concessao), (Id-Pasta $concessao)) -ForegroundColor Cyan
    $filhosConcessao = @(Filhos $concessao)
    if ($filhosConcessao.Count -eq 0) {
        # Consulta alternativa pelo caminho exato, sem interpretar $false como pasta.
        $caminhoProjetos = (Nome-Pasta $engenharia) + '\' + (Nome-Pasta $concessao) + '\Projetos'
        Write-Host "Nenhum filho retornado por ID. Consultando caminho exato: $caminhoProjetos"
        $filhosConcessao = @(Get-PWFolders -FolderPath $caminhoProjetos -JustOne -ErrorAction Stop | Where-Object { $null -ne $_ -and $_ -isnot [bool] })
        if ($filhosConcessao.Count -ne 1 -or (Nome-Pasta $filhosConcessao[0]) -ne 'Projetos') {
            throw "Pasta Projetos nao localizada na concessao selecionada. Confira a concessao, o datasource e o acesso. Diagnostico: $Destino"
        }
    }
    $projetos = Localizar $filhosConcessao 'Projetos'
    $projeto = Escolher (Filhos $projetos) 'Selecione o projeto (filhos diretos de Projetos)'
    Conferir-ARTESPPasta $projeto 'Projeto'
    $unidade = Localizar (Filhos $projeto) '2 - Unidade'
    Conferir-ARTESPPasta $unidade 'Unidade'
    # Navegacao permite Executivo > disciplina, sem presumir niveis fixos.
    $disciplina = $unidade
    while ($true) {
        Write-Host ("Pasta atual: " + (Nome-Pasta $disciplina))
        $resposta = Read-Host 'Usar esta pasta como escopo? S confirma; Enter abre subpastas; Q cancela'
        if ($resposta -eq 'S') { break }
        if ($resposta -eq 'Q') { throw 'Operacao cancelada pelo usuario.' }
        $disciplina = Escolher (Filhos $disciplina) 'Selecione a etapa ou disciplina; Q cancela'
    }
    Exportar-Seguranca $projeto 'SegurancaProjeto'
    Exportar-Seguranca $unidade 'SegurancaUnidade'
    Exportar-Seguranca $disciplina 'SegurancaDisciplina'
    Write-Host 'Lendo documentos atuais apenas da pasta selecionada (sem subpastas nem versoes historicas)...'
    $documentos = @(Get-PWDocumentsBySearch -FolderID (Id-Pasta $disciplina) -JustThisFolder -PopulatePath -ErrorAction Stop)
    $linhas = @(); $indice = 0
    foreach ($doc in $documentos) {
        $indice++
        $linhas += [pscustomobject]@{
            Indice = $indice
            Documento = Texto $doc @('Name','DocumentName','FileName')
            Arquivo = Texto $doc @('FileName')
            Workflow = Texto $doc @('Workflow','WorkflowName')
            Estado = Texto $doc @('WorkflowState','StateName','State')
            DocumentGUID = Texto $doc @('DocumentGUID','DocumentGuid','GUID')
            DocumentID = Texto $doc @('DocumentID','ID')
            FolderID = Id-Pasta $disciplina
            Versao = Texto $doc @('Version','VersionString')
        }
    }
    if ($linhas.Count -eq 0) { throw 'Nenhum documento encontrado na pasta selecionada.' }
    $linhas | Export-Csv (Join-Path $Destino 'Inventario.csv') -NoTypeInformation -Encoding UTF8
    $documentos | Export-Clixml (Join-Path $Destino 'ObjetosDocumentos.xml') -Depth 5
    Write-Host ("Total lido na pasta: {0} documentos." -f $linhas.Count) -ForegroundColor Cyan
    $estados = @($linhas | Group-Object Workflow,Estado | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Rotulo = ('{0} / {1} ({2} documentos)' -f $_.Group[0].Workflow, $_.Group[0].Estado, $_.Count)
            Workflow = $_.Group[0].Workflow; Estado = $_.Group[0].Estado; Quantidade = $_.Count
        }
    })
    $estados | Export-Csv (Join-Path $Destino 'EstadosDisponiveis.csv') -NoTypeInformation -Encoding UTF8
    do {
        $selecionados = @(Escolher $estados 'Selecione workflow/estado(s) pelos numeros do menu, separados por virgula' -Varios)
        Write-Host 'Estados selecionados:' -ForegroundColor Cyan
        $selecionados | Format-Table Workflow,Estado,Quantidade -AutoSize | Out-Host
        $totalEsperado = ($selecionados | Measure-Object Quantidade -Sum).Sum
        Write-Host ("Total esperado para esta selecao: {0} documentos." -f $totalEsperado)
        $confirmacaoEstados = Read-Host 'S confirma os estados; Q cancela; Enter refaz a selecao'
        if ($confirmacaoEstados -eq 'Q') { throw 'Operacao cancelada pelo usuario.' }
    } while ($confirmacaoEstados -ne 'S')
    $selecionados | Export-Csv (Join-Path $Destino 'EstadosSelecionados.csv') -NoTypeInformation -Encoding UTF8
    $candidatos = @($linhas | Where-Object {
        $linha = $_
        @($selecionados | Where-Object { $_.Workflow -eq $linha.Workflow -and $_.Estado -eq $linha.Estado }).Count -gt 0
    } | ForEach-Object {
        $_ | Add-Member NoteProperty Rotulo ('{0} | {1} | {2}' -f $_.Documento,$_.Estado,$_.DocumentGUID) -PassThru
    })
    if ($candidatos.Count -ne $totalEsperado) { throw 'Contagem do filtro diverge da soma dos estados selecionados.' }
    $candidatos | Export-Csv (Join-Path $Destino 'DocumentosFiltrados.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Documentos encontrados nos estados selecionados: {0}" -f $candidatos.Count) -ForegroundColor Cyan
    $candidatos | Format-Table Indice,Documento,Workflow,Estado -AutoSize | Out-Host
    if ((Read-Host 'Selecionar todos os documentos acima? S confirma; Enter permite escolher') -eq 'S') {
        $escolhidos = $candidatos
    } else { $escolhidos = @(Escolher $candidatos 'Selecione os documentos' -Varios) }
    $escolhidos | Format-Table Documento,Workflow,Estado -AutoSize | Out-Host
    $confirmacaoPlano = (Read-Host 'Confirmar geracao do plano para estes documentos? Digite SIM').Trim()
    if ($confirmacaoPlano -ine 'SIM') {
        Write-Host 'Plano documental cancelado. As etapas do projeto e Unidade podem ter sido aplicadas; consulte os relatorios.' -ForegroundColor Yellow
        return
    }
    $escolhidos | Select-Object *, @{n='Resultado';e={'Selecionado - aplicacao documental ainda nao iniciada'}},
        @{n='DireitosPropostos';e={'ARTESP: Read + File Read; preservar ACL completa'}} |
        Export-Csv (Join-Path $Destino 'PlanoDocumentos.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{
        Grupo = 'ARTESP'; Projeto = Nome-Pasta $projeto; ProjetoID = Id-Pasta $projeto
        UnidadeID = Id-Pasta $unidade; Disciplina = Nome-Pasta $disciplina
        DisciplinaID = Id-Pasta $disciplina; TotalSelecionado = @($escolhidos).Count
        Modo = 'Projeto e Unidade conferidos/aplicados; consulte ResultadoDocumentos.csv para o lote'; Subpastas = $false; VersoesHistoricas = $false
    } | ConvertTo-Json | Set-Content (Join-Path $Destino 'Resumo.json') -Encoding UTF8
    Write-Host "Plano concluido. Relatorios em: $Destino" -ForegroundColor Green
    $modoDocumental=(Read-Host 'Digite LOTE para planejar/aplicar os selecionados; TESTE para um documento; Enter encerra').Trim()
    if ($modoDocumental -ieq 'TESTE') {
        Testar-AcessoDocumento @($escolhidos)
    } elseif ($modoDocumental -ieq 'LOTE') {
        Executar-LoteARTESP -Linhas @($escolhidos) -Destino (Join-Path $Destino 'Lote')
    }
}
catch {
    if ($script:Destino -and (Test-Path -LiteralPath $script:Destino)) {
        $_ | Out-String | Set-Content (Join-Path $script:Destino 'Erro.txt') -Encoding UTF8
    }
    throw
}
finally {
    if ($loginCriado) { Undo-PWLogin -ErrorAction Continue | Out-Null }
}
