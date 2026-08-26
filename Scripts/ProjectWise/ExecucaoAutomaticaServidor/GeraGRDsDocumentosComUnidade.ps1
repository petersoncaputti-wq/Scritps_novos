[CmdletBinding()]
param(
    [string]$DatasourceName = '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01',
    [string]$UserName = 'admin',
    [string]$Password = $env:ECORODOVIAS_PW_PASSWORD,
    [switch]$UseGuiLogin,
    [string[]]$States = @('Emitido pela Engenharia'),
    [string]$UnidadeFiltro,
    [string]$ProjetoFiltro,
    [string[]]$ExcluirProjetos = @(),
    [string]$CaminhoModeloExcel,
    [string]$DiretorioTemporario,
    [string]$DiretorioSaidaExcel,
    [string]$LogDirectory,
    [switch]$SomenteTestarLogin,
    [switch]$Executar,
    [string]$ConfirmarExecucao,
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
$utf8Encoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = $utf8Encoding
try {
    [Console]::OutputEncoding = $utf8Encoding
}
catch {
    # ISE e Agendador de Tarefas podem não disponibilizar um console válido.
}

if (-not ('SimulationQueryProgress' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Threading;

public static class SimulationQueryProgress
{
    private static Thread worker;
    private static volatile bool running;
    private static DateTime startedAt;
    private static long documentsReceived;

    public static void Start()
    {
        Stop();
        startedAt = DateTime.Now;
        Interlocked.Exchange(ref documentsReceived, 0);
        running = true;
        worker = new Thread(() =>
        {
            while (running)
            {
                Thread.Sleep(5000);
                if (!running) break;
                TimeSpan elapsed = DateTime.Now - startedAt;
                try
                {
                    Console.WriteLine(
                        "[{0:yyyy-MM-dd HH:mm:ss.fff}] [ANDAMENTO] {1} documento(s) localizado(s) em {2:hh\\:mm\\:ss}.",
                        DateTime.Now,
                        Interlocked.Read(ref documentsReceived),
                        elapsed
                    );
                }
                catch
                {
                    // Execuções sem console continuam normalmente e mantêm o log principal.
                }
            }
        });
        worker.IsBackground = true;
        worker.Start();
    }

    public static void DocumentReceived()
    {
        Interlocked.Increment(ref documentsReceived);
    }

    public static void Stop()
    {
        running = false;
        if (worker != null && worker.IsAlive)
        {
            worker.Join(1000);
        }
        worker = null;
    }
}
'@
}

$diretorioBase = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    (Get-Location).Path
}
else {
    $PSScriptRoot
}

$diretorioAplicacao = if ([string]::IsNullOrWhiteSpace($env:ProgramData)) {
    Join-Path -Path $diretorioBase -ChildPath 'DadosGRD'
}
else {
    Join-Path -Path $env:ProgramData -ChildPath 'Ecorodovias\GRD'
}

if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
    $LogDirectory = Join-Path -Path $diretorioAplicacao -ChildPath 'Logs'
}

if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}

$timestampLog = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $LogDirectory "GeraGRDsDocumentosComUnidade_$timestampLog.log"
$transcriptIniciado = $false

if (-not [string]::IsNullOrWhiteSpace($DiretorioSaidaExcel) -and
    -not [string]::IsNullOrWhiteSpace($DiretorioTemporario)) {
    throw 'Informe somente -DiretorioTemporario. -DiretorioSaidaExcel foi mantido apenas para compatibilidade.'
}
if ([string]::IsNullOrWhiteSpace($DiretorioTemporario)) {
    $DiretorioTemporario = if ([string]::IsNullOrWhiteSpace($DiretorioSaidaExcel)) {
        Join-Path -Path $diretorioAplicacao -ChildPath 'Staging'
    }
    else {
        $DiretorioSaidaExcel
    }
}
if ([string]::IsNullOrWhiteSpace($CaminhoModeloExcel)) {
    $CaminhoModeloExcel = Join-Path -Path $diretorioAplicacao -ChildPath 'Modelos\ModeloGRD.xlsx'
}
$diretorioExecucaoExcel = Join-Path -Path $DiretorioTemporario -ChildPath $timestampLog

function Write-SimulationLog {
    param(
        [ValidateSet('INFO', 'AVISO', 'ERRO')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory)]
        [string]$Message
    )

    $linha = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    Write-Host $linha
}

function Get-ProximoStateSimulado {
    param([string]$StateAtual)

    switch ($StateAtual) {
        'Em Envio para Projetista' {
            return 'Em revisao pelo Projetista'
        }
        'Enviado para Unidade com Ressalvas' {
            return 'Enviado para Unidade e Projetista com Ressalvas'
        }
        default {
            return $StateAtual
        }
    }
}

function Get-EnvironmentSimulado {
    param([string]$PoderConcedente)

    switch ($PoderConcedente) {
        'ARTESP' { return 'dmsGRDRetornoARTESP' }
        'ANTT'   { return 'dmsGRDRetornoANTT' }
        default  { return '' }
    }
}

function Get-ProximaGRDSimulada {
    param(
        [string]$PrefixoGRD,
        $PastaSaida
    )

    $prefixoSQL = $PrefixoGRD.Replace("'", "''")
    $query = "select Max(o_projectname) as 'GRD' from dms_proj where o_projectname like '$prefixoSQL%' and o_parentno = $($PastaSaida.ProjectID)"
    $maiorGRD = (Select-PWSQLDataTable -SQLSelectStatement $query).GRD

    if (-not $maiorGRD) {
        return "$PrefixoGRD-0000001"
    }

    $sequencial = $maiorGRD.Split('-')[-1]
    return "$PrefixoGRD-$(([int]$sequencial + 1).ToString('D7'))"
}

function ConvertTo-NomeArquivoSeguro {
    param([Parameter(Mandatory)][string]$Nome)

    return ($Nome -replace '[<>:"/\\|?*]', '_').Trim()
}

function Remove-DiretorioSeVazio {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return
    }

    if (@(Get-ChildItem -LiteralPath $Path -Force).Count -eq 0) {
        Remove-Item -LiteralPath $Path -Force
        Write-SimulationLog -Message "Diretório temporário vazio removido: $Path"
    }
}

function New-ExcelGRDSimulado {
    param(
        [Parameter(Mandatory)]$Documentos,
        [Parameter(Mandatory)][string]$Projeto,
        [Parameter(Mandatory)][string]$PoderConcedente,
        [Parameter(Mandatory)][string]$NumeroGRD
    )

    if (-not (Test-Path -LiteralPath $CaminhoModeloExcel)) {
        throw "Modelo Excel não encontrado: $CaminhoModeloExcel"
    }

    Import-Module ImportExcel -ErrorAction Stop

    $pastaPoderConcedente = Join-Path -Path $diretorioExecucaoExcel -ChildPath $PoderConcedente
    if (-not (Test-Path -LiteralPath $pastaPoderConcedente)) {
        New-Item -ItemType Directory -Path $pastaPoderConcedente -Force | Out-Null
    }

    $nomeArquivo = "$(ConvertTo-NomeArquivoSeguro -Nome $NumeroGRD).xlsx"
    $caminhoExcel = Join-Path -Path $pastaPoderConcedente -ChildPath $nomeArquivo
    Copy-Item -LiteralPath $CaminhoModeloExcel -Destination $caminhoExcel -Force

    $pacote = $null
    try {
        $pacote = Open-ExcelPackage -Path $caminhoExcel
        $planilha = $pacote.Workbook.Worksheets['Relatorio']
        if (-not $planilha) {
            throw "A aba 'Relatorio' não existe no modelo Excel."
        }

        $ultimaLinhaModelo = $planilha.Dimension.End.Row
        if ($ultimaLinhaModelo -ge 5) {
            $planilha.DeleteRow(5, $ultimaLinhaModelo - 4)
        }

        $planilha.Cells['C1'].Value = $Projeto
        $planilha.Cells['C2'].Value = 'Ultimas Emissões de Documentos'

        $linha = 5
        foreach ($documento in @($Documentos | Sort-Object CreateDate, Name, Version)) {
            if ($linha -gt 5) {
                $planilha.Cells['A5:I5'].Copy($planilha.Cells["A$linha`:I$linha"])
            }

            $fileName = if ($documento.FileName) { [string]$documento.FileName } else { [string]$documento.Name }
            $valores = @(
                [string]$documento.Attributes.NumeroPoderConcedente
                [string]$documento.Attributes.Revisao
                [string]$documento.Attributes.Versao
                [IO.Path]::GetExtension($fileName).TrimStart('.').ToLowerInvariant()
                [string]$documento.Attributes.Disciplina
                [string]$documento.Attributes.FaseProjetoPai
                [string]$documento.WorkflowState
                [string]$documento.Attributes.GestorEngenharia
                $documento.CreateDate
            )

            for ($coluna = 1; $coluna -le 9; $coluna++) {
                $planilha.Cells[$linha, $coluna].Value = $valores[$coluna - 1]
            }
            $linha++
        }

        $ultimaLinha = $linha - 1
        $planilha.Cells["A4:I$ultimaLinha"].AutoFilter = $true
        $planilha.View.FreezePanes(5, 1)
        $planilha.Cells["I5:I$ultimaLinha"].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm:ss'

        Close-ExcelPackage $pacote
        $pacote = $null
        return $caminhoExcel
    }
    finally {
        if ($pacote) {
            Close-ExcelPackage $pacote -NoSave
        }
    }
}

$loginEfetuado = $false

try {
    Start-Transcript -LiteralPath $logPath -Force | Out-Null
    $transcriptIniciado = $true

    Write-SimulationLog -Message 'Início do processamento. O modo de operação será validado antes de qualquer gravação.'
    Write-SimulationLog -Message "Datasource: $DatasourceName"
    Write-SimulationLog -Message "Usuário: $UserName"
    Write-SimulationLog -Message "Consulta direta por estado(s): $($States -join ', ')"
    Write-SimulationLog -Message 'Versões anteriores: incluídas (-GetVersionsToo).'
    Write-SimulationLog -Message "Filtro de Unidade: $(if ($UnidadeFiltro) { $UnidadeFiltro } else { 'todos' })"
    Write-SimulationLog -Message "Filtro de Projeto: $(if ($ProjetoFiltro) { $ProjetoFiltro } else { 'todos' })"
    Write-SimulationLog -Message "Projetos excluídos: $(if ($ExcluirProjetos.Count) { $ExcluirProjetos -join ', ' } else { 'nenhum' })"
    Write-SimulationLog -Message 'Limite local: nenhum; todos os resultados da pesquisa serão analisados.'
    Write-SimulationLog -Message "Modelo Excel: $CaminhoModeloExcel"
    Write-SimulationLog -Message "Diretório temporário desta execução: $diretorioExecucaoExcel"
    Write-SimulationLog -Message "Arquivo de log: $logPath"
    if ($Executar -and $Simular) {
        throw 'Os parâmetros -Executar e -Simular não podem ser usados juntos.'
    }
    $modoExecucao = -not $Simular
    Write-SimulationLog -Message "Modo: $(if ($modoExecucao) { 'EXECUÇÃO REAL' } else { 'SIMULAÇÃO' })."
    if ($modoExecucao -and -not $Executar) {
        Write-SimulationLog -Level AVISO -Message 'Execução real iniciada pelo comportamento padrão do script.'
    }

    Write-SimulationLog -Message 'Validando pré-requisitos locais antes do login.'
    if (-not (Test-Path -LiteralPath $CaminhoModeloExcel -PathType Leaf)) {
        $diretorioModelo = [IO.Path]::GetDirectoryName($CaminhoModeloExcel)
        throw "Modelo Excel não encontrado: $CaminhoModeloExcel. Crie o diretório '$diretorioModelo' e copie o modelo para esse local antes de executar novamente."
    }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "O módulo PowerShell 'ImportExcel' não está instalado ou não está disponível para a conta que executa o script."
    }
    Write-SimulationLog -Message 'Pré-requisitos locais validados: modelo Excel e módulo ImportExcel disponíveis.'

    Write-SimulationLog -Message 'Iniciando login lógico no ProjectWise.'
    if ($UseGuiLogin -or [string]::IsNullOrWhiteSpace($Password)) {
        $login = New-PWLogin -DatasourceName $DatasourceName -UseGui -DoNotCreateWorkingDirectory
    }
    else {
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $login = New-PWLogin -DatasourceName $DatasourceName -UserName $UserName -Password $securePassword -DoNotCreateWorkingDirectory
    }
    if (-not $login) {
        throw 'Não foi possível efetuar login no ProjectWise.'
    }
    $loginEfetuado = $true
    Write-SimulationLog -Message 'Login no ProjectWise concluído.'

    if ($SomenteTestarLogin) {
        Write-SimulationLog -Message 'Teste de login concluído. A pesquisa não será executada.'
        return
    }

    # Consulta completa da pesquisa salva. Nenhuma limitação local é aplicada.
    Write-SimulationLog -Message 'Iniciando consulta direta no ProjectWise.'
    $cronometroConsulta = [System.Diagnostics.Stopwatch]::StartNew()
    [SimulationQueryProgress]::Start()
    try {
        $documentosConsultados = @(
            Get-PWDocumentsBySearchExtended `
                -States $States `
                -GetVersionsToo `
                -ColumnsToReturn @(
                    'Unidade'
                    'Projeto'
                    'PoderConcedente'
                    'ComentarioParaProjetista'
                    'NumeroGRDSaida'
                    'NumeroPoderConcedente'
                    'Revisao'
                    'Versao'
                    'Disciplina'
                    'FaseProjetoPai'
                    'GestorEngenharia'
                ) |
                ForEach-Object {
                    [SimulationQueryProgress]::DocumentReceived()
                    $_
                }
        )
    }
    finally {
        [SimulationQueryProgress]::Stop()
        $cronometroConsulta.Stop()
    }
    Write-SimulationLog -Message "Tempo total da consulta: $($cronometroConsulta.Elapsed.ToString('hh\:mm\:ss'))."
    Write-SimulationLog -Message "Consulta concluída. Itens recebidos pelo simulador: $($documentosConsultados.Count)."
    $documentos = @(
        $documentosConsultados |
            Where-Object {
                (-not $UnidadeFiltro -or $_.Attributes.Unidade -eq $UnidadeFiltro) -and
                (-not $ProjetoFiltro -or $_.Attributes.Projeto -eq $ProjetoFiltro) -and
                ($ExcluirProjetos -notcontains [string]$_.Attributes.Projeto)
            }
    )
    Write-SimulationLog -Message "Documentos após os filtros locais: $($documentos.Count)."

    if ($documentos.Count -eq 0) {
        Write-SimulationLog -Level AVISO -Message 'Nenhum documento permaneceu após a consulta e os filtros locais.'
        return
    }

    Write-SimulationLog -Message 'Agrupando documentos por PoderConcedente, Unidade e Projeto.'
    $grupos = $documentos | Group-Object -Property @{
        Expression = {
            "$($_.Attributes.PoderConcedente)|$($_.Attributes.Unidade)|$($_.Attributes.Projeto)"
        }
    }
    Write-SimulationLog -Message "Total de grupos encontrados: $(@($grupos).Count)."

    $resultado = foreach ($grupo in $grupos) {
        $documentoReferencia = $grupo.Group[0]
        $poderConcedente = $documentoReferencia.Attributes[0].PoderConcedente
        $unidade = $documentoReferencia.Attributes[0].Unidade
        $nomeProjeto = $documentoReferencia.Attributes[0].Projeto
        $prefixoGRD = "GRD-COM-$unidade-$nomeProjeto"
        Write-SimulationLog -Message "Analisando grupo '$prefixoGRD' ($poderConcedente) com $($grupo.Count) documento(s)."
        $projeto = Get-PWRichProjectForDocument -InputDocument $documentoReferencia
        $caminhoDocumentosGerais = "$($projeto.FullPath)\0 - Documentos Gerais"
        $caminhoDocumentosUnidade = "$caminhoDocumentosGerais\Documentos com unidade"
        $pastaDocumentosGerais = Get-PWFolders -FolderPath $caminhoDocumentosGerais -JustOne -WarningAction SilentlyContinue
        $pastaDocumentosUnidade = Get-PWFolders -FolderPath $caminhoDocumentosUnidade -JustOne -WarningAction SilentlyContinue

        if (-not $poderConcedente -or -not $unidade -or -not $nomeProjeto) {
            $motivo = if (-not $poderConcedente) {
                'PoderConcedente não preenchido'
            }
            else {
                'Unidade ou Projeto não preenchido'
            }
            Write-SimulationLog -Level AVISO -Message "Grupo '$prefixoGRD' ignorado. $motivo"
            [pscustomobject]@{
                Grupo                 = $prefixoGRD
                PoderConcedente       = $poderConcedente
                GRDPrevista           = ''
                Documentos            = $grupo.Count
                ExcelLocal            = ''
                DestinoPrevisto       = ''
                Resultado             = "Ignorado: $motivo"
            }
            continue
        }

        $pastasACriar = @()
        if (-not $pastaDocumentosGerais) {
            $pastasACriar += $caminhoDocumentosGerais
        }
        if (-not $pastaDocumentosUnidade) {
            $pastasACriar += $caminhoDocumentosUnidade
        }

        if ($pastaDocumentosUnidade) {
            $grdPrevista = Get-ProximaGRDSimulada -PrefixoGRD $prefixoGRD -PastaSaida $pastaDocumentosUnidade
        }
        else {
            $grdPrevista = "$prefixoGRD-0000001"
        }

        $destinoPrevisto = "$caminhoDocumentosUnidade\$grdPrevista"
        $pastasACriar += $destinoPrevisto

        if ($pastasACriar.Count -gt 1) {
            $verboPastas = if ($modoExecucao) { 'Precisam ser criadas' } else { 'Seriam criadas' }
            Write-SimulationLog -Level AVISO -Message "Estrutura ausente para '$prefixoGRD'. ${verboPastas}: $($pastasACriar -join '; ')"
        }
        Write-SimulationLog -Message "Grupo '$prefixoGRD': GRD prevista '$grdPrevista'; destino '$destinoPrevisto'. Nenhuma pasta receberá ambiente."

        $excelLocal = New-ExcelGRDSimulado `
            -Documentos $grupo.Group `
            -Projeto $nomeProjeto `
            -PoderConcedente $poderConcedente `
            -NumeroGRD $grdPrevista
        Write-SimulationLog -Message "Excel local gerado para '$prefixoGRD': $excelLocal"

        $documentoCriado = $null
        if ($modoExecucao) {
            foreach ($caminhoPasta in $pastasACriar) {
                $pastaExistente = Get-PWFolders -FolderPath $caminhoPasta -JustOne -WarningAction SilentlyContinue
                if ($pastaExistente) {
                    Write-SimulationLog -Message "Pasta já existente, mantida sem alterações: $caminhoPasta (ProjectID=$($pastaExistente.ProjectID))."
                    continue
                }

                Write-SimulationLog -Message "Criando pasta sem ambiente: $caminhoPasta"
                $pastaCriada = New-PWFolder -FolderPath $caminhoPasta
                if (-not $pastaCriada) {
                    throw "O ProjectWise não confirmou a criação da pasta: $caminhoPasta"
                }
                Write-SimulationLog -Message "Pasta criada: $caminhoPasta (ProjectID=$($pastaCriada.ProjectID))."
            }

            $pastaDestino = Get-PWFolders -FolderPath $destinoPrevisto -JustOne -WarningAction SilentlyContinue
            if (-not $pastaDestino) {
                throw "A pasta final não foi localizada após a criação: $destinoPrevisto"
            }

            $nomeDocumentoExcel = [IO.Path]::GetFileName($excelLocal)
            $documentoExistente = @(
                Get-PWDocumentsBySearch -FolderPath $destinoPrevisto -DocumentName $nomeDocumentoExcel -JustThisFolder -WarningAction SilentlyContinue
            )
            if ($documentoExistente.Count -gt 0) {
                throw "Já existe um documento chamado '$nomeDocumentoExcel' na pasta final. Nada será sobrescrito."
            }

            Write-SimulationLog -Message "Importando Excel na pasta final: $nomeDocumentoExcel"
            $documentoCriado = New-PWDocument `
                -InputFolders $pastaDestino `
                -FilePath $excelLocal `
                -DocumentName $nomeDocumentoExcel `
                -Description "Lista de documentos da $grdPrevista"
            if (-not $documentoCriado) {
                throw "O ProjectWise não confirmou a importação do Excel '$nomeDocumentoExcel'."
            }
            Write-SimulationLog -Message "Excel importado com sucesso: DocumentID=$($documentoCriado.DocumentID); caminho=$destinoPrevisto\$nomeDocumentoExcel"

            Write-SimulationLog -Message "Confirmando o upload por uma nova consulta ao ProjectWise: $nomeDocumentoExcel"
            $documentosConfirmados = @(
                Get-PWDocumentsBySearch `
                    -FolderPath $destinoPrevisto `
                    -DocumentName $nomeDocumentoExcel `
                    -JustThisFolder `
                    -PopulatePath `
                    -WarningAction SilentlyContinue |
                    Where-Object {
                        $_.Name -eq $nomeDocumentoExcel -or
                        $_.FileName -eq $nomeDocumentoExcel
                    }
            )
            if ($documentosConfirmados.Count -ne 1) {
                throw "O upload não pôde ser confirmado de forma inequívoca. Foram encontrados $($documentosConfirmados.Count) documento(s) com o nome '$nomeDocumentoExcel'. O arquivo local será preservado em '$excelLocal'."
            }

            $documentoConfirmado = $documentosConfirmados[0]
            Write-SimulationLog -Message "Upload confirmado: DocumentID=$($documentoConfirmado.DocumentID); FullPath=$($documentoConfirmado.FullPath)"

            Remove-Item -LiteralPath $excelLocal -Force
            if (Test-Path -LiteralPath $excelLocal) {
                throw "O upload foi confirmado, mas não foi possível excluir o arquivo temporário: $excelLocal"
            }
            Write-SimulationLog -Message "Arquivo temporário excluído após confirmação do upload: $excelLocal"
            Remove-DiretorioSeVazio -Path ([IO.Path]::GetDirectoryName($excelLocal))
        }

        [pscustomobject]@{
            Grupo                 = $prefixoGRD
            PoderConcedente       = $poderConcedente
            GRDPrevista           = $grdPrevista
            Documentos            = $grupo.Count
            ExcelLocal            = $excelLocal
            DestinoPrevisto       = $destinoPrevisto
            PastasACriar          = $pastasACriar -join '; '
            Resultado             = if ($modoExecucao) {
                "Criado no ProjectWise; DocumentID=$($documentoCriado.DocumentID)"
            }
            else {
                'Simulado; nenhuma pasta terá ambiente'
            }
        }
    }

    Write-Host ''
    if ($modoExecucao) {
        Write-SimulationLog -Message 'EXECUÇÃO REAL CONCLUÍDA — estrutura e planilha criadas no ProjectWise.'
    }
    else {
        Write-SimulationLog -Message 'SIMULAÇÃO CONCLUÍDA — nenhuma pasta ou documento foi alterado.'
    }
    if ($modoExecucao) {
        Remove-DiretorioSeVazio -Path $diretorioExecucaoExcel
    }
    Write-SimulationLog -Message "Documentos considerados: $($documentos.Count) (pesquisa completa)."
    $resultado | Format-Table Grupo, PoderConcedente, GRDPrevista, Documentos, ExcelLocal -AutoSize

    Write-Host ''
    Write-Host 'Resumo:'
    $resultado | Format-Table Grupo, PoderConcedente, GRDPrevista, Documentos, DestinoPrevisto, PastasACriar, Resultado -AutoSize
}
catch {
    Write-SimulationLog -Level ERRO -Message $_.Exception.Message
    Write-SimulationLog -Level ERRO -Message ($_ | Format-List * -Force | Out-String)
    throw
}
finally {
    if ($loginEfetuado) {
        Write-SimulationLog -Message 'Encerrando login no ProjectWise.'
        try {
            Undo-PWLogin | Out-Null
            Write-SimulationLog -Message 'Logout concluído.'
        }
        catch {
            Write-SimulationLog -Level ERRO -Message "Falha durante o logout: $($_.Exception.Message)"
        }
    }

    if ($transcriptIniciado) {
        Write-SimulationLog -Message "Fim da execução. Log: $logPath"
        Stop-Transcript | Out-Null
    }
}
