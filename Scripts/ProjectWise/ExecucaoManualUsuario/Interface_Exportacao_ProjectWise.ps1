[CmdletBinding()]
param([string]$DatasourceName = '', [string]$Destino, [switch]$SomenteListar)

# UI STA e um unico trabalhador MTA: a sessao PW nunca cruza threads.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $argumentos = @('-NoProfile', '-STA', '-File', $PSCommandPath)
    if ($DatasourceName) { $argumentos += @('-DatasourceName', $DatasourceName) }
    if ($Destino) { $argumentos += @('-Destino', $Destino) }
    if ($SomenteListar) { $argumentos += '-SomenteListar' }
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" @argumentos
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$estado = [hashtable]::Synchronized(@{ Cancelar = $false; Encerrar = $false })
$pedidos = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
$eventos = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'

$trabalhador = {
    param($estado, $pedidos, $eventos, $arquivoBase)
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    function Emitir($tipo, $dados) { $eventos.Enqueue([pscustomobject]@{ Tipo = $tipo; Dados = $dados }) }
    $login = $false
    try {
        # Reutiliza somente as funcoes do script de console, sem executar seu fluxo.
        $tokens = $null; $erros = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($arquivoBase, [ref]$tokens, [ref]$erros)
        if ($erros.Count) { throw 'O script base possui erros de sintaxe.' }
        foreach ($declaracao in $ast.EndBlock.Statements) {
            if ($declaracao -is [Management.Automation.Language.FunctionDefinitionAst]) {
                . ([scriptblock]::Create($declaracao.Extent.Text))
            }
        }
        $concessoes = @(); $projetos = @(); $areas = @()
        while (-not $estado.Encerrar) {
            $pedido = $null
            if (-not $pedidos.TryDequeue([ref]$pedido)) { Start-Sleep -Milliseconds 100; continue }
            try {
                switch ($pedido.Acao) {
                    'Conectar' {
                        Emitir 'Mensagem' 'Interface v1.5 | Preparando PWPS_DAB...'
                        # O inicializador do PWPS_DAB 24.0.2 acessa
                        # $Host.Runspace.ThreadOptions. O host padrao de uma
                        # runspace criada por API nao expoe essa propriedade.
                        # Fornece a runspace MTA atual sem alterar o modulo.
                        $runspaceAtual = [runspace]::DefaultRunspace
                        if ($null -eq $runspaceAtual) { throw 'Runspace atual indisponivel para inicializar o PWPS_DAB.' }
                        if ($null -eq $Host.PSObject.Properties['Runspace'] -or $null -eq $Host.Runspace) {
                            $Host | Add-Member -MemberType NoteProperty -Name Runspace -Value $runspaceAtual -Force
                        }
                        if ($null -eq $Host.Runspace.PSObject.Properties['ThreadOptions']) {
                            throw 'O host nao disponibilizou Runspace.ThreadOptions. Feche a interface e execute novamente em Windows PowerShell 5.1.'
                        }
                        Emitir 'Mensagem' "Host: $($Host.Name) | Apartment: $($runspaceAtual.ApartmentState) | ThreadOptions: $($Host.Runspace.ThreadOptions)"
                        # O inicializador tambem consulta .Length em uma
                        # variavel de ambiente que pode estar ausente.
                        Set-StrictMode -Off
                        Import-Module PWPS_DAB -ErrorAction Stop
                        Emitir 'Mensagem' 'Modulo importado. Abrindo login...'
                        $p = @{ UseGui = $true; DoNotCreateWorkingDirectory = $true; ErrorAction = 'Stop' }
                        if ($pedido.Datasource) { $p.DatasourceName = $pedido.Datasource }
                        if (-not $login) {
                            if (-not (New-PWLogin @p)) { throw 'Login cancelado ou nao concluido.' }
                            $login = $true
                        }
                        $engenharia = Localizar-Pasta (Obter-FilhosRaiz) @('Engenharia', 'Engineering') 'Engenharia'
                        $concessoes = @(Obter-FilhosPasta (Obter-IdPasta $engenharia) | Sort-Object { Obter-RotuloPasta $_ })
                        Emitir 'Concessoes' @($concessoes | ForEach-Object { Obter-RotuloPasta $_ })
                    }
                    'Projetos' {
                        $areas = @()
                        $concessao = $concessoes[$pedido.Indice]
                        $nomeConcessao = Obter-RotuloPasta $concessao
                        $pasta = Localizar-Pasta (Obter-FilhosPasta (Obter-IdPasta $concessao)) @('Projetos','Projeto','Projects','Project') 'Projetos'
                        $projetos = @(Obter-FilhosPasta (Obter-IdPasta $pasta) | Sort-Object { Obter-RotuloPasta $_ })
                        Emitir 'Projetos' @($projetos | ForEach-Object { Obter-RotuloPasta $_ })
                    }
                    'Localizar' {
                        $areas = @()
                        $lista = New-Object 'System.Collections.Generic.List[object]'
                        foreach ($indice in $pedido.Indices) {
                            if ($estado.Cancelar) { break }
                            $projeto = $projetos[$indice]
                            $nome = Obter-RotuloPasta $projeto
                            Emitir 'Mensagem' "Localizando areas: $nome"
                            try {
                                foreach ($area in (Localizar-AreasTransferenciaNoProjeto $projeto)) {
                                    $lista.Add([pscustomobject]@{
                                        Concessao = $nomeConcessao; Projeto = $nome
                                        ProjetoID = Obter-IdPasta $projeto; AreaID = Obter-IdPasta $area
                                        CaminhoPW = Obter-CaminhoPW $area
                                    })
                                }
                            } catch { Emitir 'Mensagem' "Falha na pesquisa de ${nome}: $($_.Exception.Message)" }
                        }
                        $areas = $lista.ToArray()
                        Emitir 'Areas' $areas
                        Emitir 'Mensagem' "Pesquisa encerrada: $($areas.Count) area(s). Consulte os detalhes para eventuais falhas."
                    }
                    'Exportar' {
                        if (-not $areas.Count) { throw 'Localize as areas antes de exportar.' }
                        $destino = [IO.Path]::GetFullPath($pedido.Destino)
                        $null = [IO.Directory]::CreateDirectory($destino)
                        $destinoConcessao = Join-Path $destino (Converter-NomePastaLocal $nomeConcessao)
                        $null = [IO.Directory]::CreateDirectory($destinoConcessao)
                        $sufixo = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
                        $relatorio = Join-Path $destino "Relatorio_Areas_Transferencia_$sufixo.csv"
                        $areas | Export-Csv -LiteralPath $relatorio -NoTypeInformation -Encoding UTF8
                        Emitir 'Mensagem' "Inventario: $relatorio"
                        if ($pedido.SomenteListar) { Emitir 'Mensagem' 'SomenteListar: inventario salvo, sem exportacao.'; break }
                        $log = Join-Path $destino "Log_Exportacao_$sufixo.csv"
                        $logDocumentos = Join-Path $destino "Log_Documentos_$sufixo.csv"
                        $falhas = 0; $processadas = 0
                        for ($i = 0; $i -lt $areas.Count; $i++) {
                            if ($estado.Cancelar) { break }
                            $area = $areas[$i]
                            Emitir 'Linha' @{ Indice = $i; Status = 'Exportando'; Retornos = ''; Duracao = '' }
                            Emitir 'Mensagem' "Exportando $($i + 1)/$($areas.Count): $($area.Projeto) | $($area.CaminhoPW)"
                            $inicio = Get-Date
                            $relogio = [Diagnostics.Stopwatch]::StartNew()
                            $retornos = 0; $erro = ''; $status = 'Finalizada*'
                            $contagem = ''; $totalDocumentos = 0; $documentosProcessados = 0; $falhasDocumentos = 0
                            try {
                                if (-not $area.CaminhoPW) { throw 'Caminho ProjectWise vazio.' }
                                if ($pedido.Detalhado) {
                                    Emitir 'Arquivos' @{ Valor = 0; Texto = 'Contando documentos da area atual...' }
                                    $documentos = @(Get-PWDocumentsBySearch -FolderID ([int]$area.AreaID) -ErrorAction Stop |
                                        Where-Object { -not [string]::IsNullOrWhiteSpace((Obter-Texto $_ @('FileName'))) })
                                    $totalDocumentos = $documentos.Count
                                    $contagem = "0/$totalDocumentos"
                                    foreach ($documento in $documentos) {
                                        if ($estado.Cancelar) { break }
                                        $nomeArquivo = Obter-Texto $documento @('FileName', 'Name')
                                        Emitir 'Arquivos' @{ Valor = [int](100 * $documentosProcessados / $totalDocumentos); Texto = "$documentosProcessados/$totalDocumentos processados | Atual: $nomeArquivo | Falhas: $falhasDocumentos" }
                                        $erroDocumento = ''; $estadoDocumento = 'Sem erro*'
                                        try {
                                            $resultado = @(Export-PWDocuments -InputDocuments @($documento) -OutputFolder $destinoConcessao -UseUnderscores -ErrorAction Stop)
                                            $retornos += $resultado.Count
                                            if ($resultado.Count -eq 0) { $estadoDocumento = 'Sem retorno'; }
                                        } catch {
                                            $erroDocumento = $_.Exception.Message
                                            $estadoDocumento = 'Falha'; $falhasDocumentos++
                                            Emitir 'Mensagem' "Falha em ${nomeArquivo}: $erroDocumento"
                                        }
                                        $documentosProcessados++
                                        $contagem = "$documentosProcessados/$totalDocumentos"
                                        [pscustomobject]@{
                                            AreaID = $area.AreaID; Projeto = $area.Projeto
                                            DocumentoID = Obter-Texto $documento @('DocumentGUID','DocumentID','ID')
                                            Arquivo = $nomeArquivo; Status = $estadoDocumento; Erro = $erroDocumento; Fim = Get-Date
                                        } | Export-Csv -LiteralPath $logDocumentos -NoTypeInformation -Encoding UTF8 -Append
                                        Emitir 'Linha' @{ Indice = $i; Status = 'Exportando'; Retornos = $contagem; Duracao = $relogio.Elapsed.ToString('hh\:mm\:ss') }
                                        Emitir 'Arquivos' @{ Valor = [int](100 * $documentosProcessados / $totalDocumentos); Texto = "$contagem processados | Ultimo: $nomeArquivo | Falhas: $falhasDocumentos" }
                                    }
                                    if ($falhasDocumentos -gt 0) { $status = 'Com falhas'; $falhas++; $erro = "$falhasDocumentos documento(s) com falha. Consulte $logDocumentos" }
                                    if ($documentosProcessados -lt $totalDocumentos) { $status = 'Interrompida' }
                                    if ($totalDocumentos -eq 0) { $status = 'Sem arquivos'; Emitir 'Arquivos' @{ Valor = 0; Texto = 'Nenhum documento com arquivo nesta area.' } }
                                } else {
                                    Emitir 'Arquivos' @{ Valor = 0; Texto = 'Modo em lote: contagem disponivel somente ao terminar a area.' }
                                    $resultado = @(Export-PWDocuments -ProjectWiseFolder $area.CaminhoPW -OutputFolder $destinoConcessao -UseUnderscores -ErrorAction Stop)
                                    $retornos = $resultado.Count
                                    $contagem = [string]$retornos
                                }
                            } catch { $erro = $_.Exception.Message; $status = 'Falha'; $falhas++; Emitir 'Mensagem' "Falha: $erro" }
                            $relogio.Stop()
                            [pscustomobject]@{
                                Concessao = $area.Concessao; Projeto = $area.Projeto; AreaID = $area.AreaID
                                CaminhoPW = $area.CaminhoPW; DestinoLocal = $destinoConcessao
                                Status = $status; DocumentosRetornados = $retornos
                                DocumentosProcessados = $documentosProcessados; TotalDocumentos = $totalDocumentos; FalhasDocumentos = $falhasDocumentos
                                Inicio = $inicio; Fim = Get-Date; DuracaoSegundos = $relogio.Elapsed.TotalSeconds; Erro = $erro
                            } | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8 -Append
                            if ($status -ne 'Interrompida') { $processadas++ }
                            Emitir 'Linha' @{ Indice = $i; Status = $status; Retornos = $contagem; Duracao = $relogio.Elapsed.ToString('hh\:mm\:ss') }
                            Emitir 'Progresso' @{ Valor = [int](100 * $processadas / $areas.Count); Texto = "$processadas/$($areas.Count) areas processadas | $falhas falha(s)" }
                        }
                        Emitir 'Mensagem' "Exportacao encerrada: $processadas/$($areas.Count) areas processadas; $falhas falha(s). Log: $log"
                        if ($pedido.Detalhado) { Emitir 'Mensagem' "Detalhes por documento: $logDocumentos" }
                        if ($estado.Cancelar) { Emitir 'Mensagem' 'Parada solicitada. As areas restantes nao foram exportadas.' }
                    }
                }
            } catch {
                Emitir 'Erro' $_.Exception.Message
                Emitir 'Mensagem' "Origem do erro: $($_.InvocationInfo.PositionMessage)"
                Emitir 'Mensagem' "Pilha: $($_.ScriptStackTrace)"
            }
            finally { Emitir 'Livre' $null }
        }
    } catch { Emitir 'Erro' $_.Exception.Message }
    finally { if ($login) { try { Undo-PWLogin -ErrorAction Stop | Out-Null } catch { Emitir 'Erro' $_.Exception.Message } } }
}

$form = New-Object Windows.Forms.Form
$form.Text = 'ProjectWise | Exportar Areas de Transferencia | v1.5'
$form.Font = New-Object Drawing.Font('Segoe UI', 10)
$form.AutoScaleMode = 'Dpi'
$form.AutoScroll = $true
$form.MinimumSize = New-Object Drawing.Size(400, 320)
$form.StartPosition = 'CenterScreen'
$areaTela = [Windows.Forms.Screen]::FromPoint([Windows.Forms.Cursor]::Position).WorkingArea
$form.Size = New-Object Drawing.Size([Math]::Min(1120, $areaTela.Width), [Math]::Min(900, $areaTela.Height))

$layout = New-Object Windows.Forms.TableLayoutPanel
$layout.ColumnCount = 1; $layout.RowCount = 0
$layout.Dock = 'Top'; $layout.AutoSize = $true; $layout.AutoSizeMode = 'GrowAndShrink'
$layout.Padding = New-Object Windows.Forms.Padding(12)
$null = $layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
$form.Controls.Add($layout)
function Controle($tipo, $texto) {
    $c = New-Object "Windows.Forms.$tipo"
    $c.Text = $texto
    $c.Margin = New-Object Windows.Forms.Padding(3, 4, 3, 4)
    return $c
}
function Adicionar-Linha($controle) {
    $linha = $layout.RowCount
    $layout.RowCount++
    $null = $layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    $layout.Controls.Add($controle, 0, $linha)
    $controle.Dock = 'Top'
}
function Barra-Botoes($controles) {
    $painel = New-Object Windows.Forms.FlowLayoutPanel
    $painel.AutoSize = $true; $painel.AutoSizeMode = 'GrowAndShrink'
    $painel.WrapContents = $true
    foreach ($c in $controles) {
        if ($c -is [Windows.Forms.Button] -or $c -is [Windows.Forms.CheckBox]) {
            $c.AutoSize = $true
            $c.Padding = New-Object Windows.Forms.Padding(6)
        }
        $painel.Controls.Add($c)
    }
    Adicionar-Linha $painel
}
$titulo = Controle Label 'EXPORTAR AREAS DE TRANSFERENCIA'
$titulo.AutoSize = $true
Adicionar-Linha $titulo
$conectar = Controle Button 'Conectar ao ProjectWise'
$concessao = Controle ComboBox ''
$concessao.DropDownStyle = 'DropDownList'
$concessao.DropDownWidth = 600
$carregar = Controle Button 'Carregar projetos'
Barra-Botoes @($conectar, $carregar)
Adicionar-Linha $concessao
$rotuloProjetos = Controle Label 'Projetos (marque os que deseja pesquisar)'
$rotuloProjetos.AutoSize = $true
Adicionar-Linha $rotuloProjetos
$projetosUI = Controle CheckedListBox ''
$projetosUI.CheckOnClick = $true
$projetosUI.IntegralHeight = $false
$projetosUI.HorizontalScrollbar = $true
Adicionar-Linha $projetosUI
$todos = Controle Button 'Marcar todos'
$nenhum = Controle Button 'Desmarcar todos'
$localizar = Controle Button 'Localizar areas'
Barra-Botoes @($todos, $nenhum, $localizar)
$rotuloDestino = Controle Label 'Destino:'
$rotuloDestino.AutoSize = $true
Adicionar-Linha $rotuloDestino
$destinoUI = Controle TextBox $Destino
Adicionar-Linha $destinoUI
$escolher = Controle Button 'Escolher destino...'
$abrir = Controle Button 'Abrir destino'
Barra-Botoes @($escolher, $abrir)
$iniciar = Controle Button $(if ($SomenteListar) { 'Salvar inventario' } else { 'Iniciar exportacao' })
$cancelar = Controle Button 'Cancelar'
$detalhado = Controle CheckBox 'Acompanhar por arquivo (mais lento)'
$detalhado.Checked = $true
Barra-Botoes @($iniciar, $cancelar, $detalhado)
$grid = Controle DataGridView ''
$grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false; $grid.AutoSizeColumnsMode = 'Fill'
$grid.DefaultCellStyle.WrapMode = 'True'
$grid.AutoSizeRowsMode = 'AllCells'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.ScrollBars = 'Both'
foreach ($coluna in @('Projeto','Area ID','Estado','Processados / total','Duracao')) { $null = $grid.Columns.Add($coluna, $coluna) }
$grid.Columns[0].FillWeight = 300
foreach ($coluna in $grid.Columns) { $coluna.MinimumWidth = 95 }
$grid.Columns[0].MinimumWidth = 230
Adicionar-Linha $grid
$barra = Controle ProgressBar ''
$barra.Height = 20
Adicionar-Linha $barra
$situacao = Controle Label 'Conecte ao ProjectWise para comecar.'
$situacao.AutoSize = $true
Adicionar-Linha $situacao
$tempo = Controle Label 'Tempo: 00:00:00'
$tempo.AutoSize = $true
Adicionar-Linha $tempo
$barraArquivos = Controle ProgressBar ''
$barraArquivos.Height = 20
Adicionar-Linha $barraArquivos
$situacaoArquivos = Controle Label 'Progresso dos arquivos da area atual.'
$situacaoArquivos.AutoSize = $true
Adicionar-Linha $situacaoArquivos
$nota = Controle Label '* Processados incluem tentativas com falha e arquivos ja atualizados. Sem erro nao significa copia verificada.'
$nota.AutoSize = $true
Adicionar-Linha $nota
$detalhes = Controle TextBox ''
$detalhes.Multiline = $true; $detalhes.ReadOnly = $true; $detalhes.ScrollBars = 'Vertical'; $detalhes.WordWrap = $true
Adicionar-Linha $detalhes
$detalhes.AppendText("Interface v1.5 | Arquivo: $PSCommandPath`r`n")
function Ajustar-Layout {
    $largura = [Math]::Max(100, $layout.ClientSize.Width - $layout.Padding.Horizontal - 12)
    foreach ($rotulo in @($titulo, $rotuloProjetos, $rotuloDestino, $situacao, $tempo, $situacaoArquivos, $nota)) {
        $rotulo.MaximumSize = New-Object Drawing.Size($largura, 0)
    }
    $escala = $form.Font.Height / 18.0
    $projetosUI.Height = [Math]::Max([int](90 * $escala), [int]($form.ClientSize.Height * 0.13))
    $grid.Height = [Math]::Max([int](130 * $escala), [int]($form.ClientSize.Height * 0.20))
    $detalhes.Height = [Math]::Max([int](90 * $escala), [int]($form.ClientSize.Height * 0.12))
}
$form.Add_ClientSizeChanged({ Ajustar-Layout })
$layout.Add_SizeChanged({ Ajustar-Layout })
Ajustar-Layout

$ui = @{ Ocupado = $false; Conectado = $false; Areas = 0; Relogio = [Diagnostics.Stopwatch]::new(); Fechar = $false }
function Atualizar-Botoes {
    $livre = -not $ui.Ocupado -and -not $estado.Encerrar
    $conectar.Enabled = $livre -and -not $ui.Conectado
    $concessao.Enabled = $livre -and $ui.Conectado
    $carregar.Enabled = $livre -and $concessao.SelectedIndex -ge 0
    $projetosUI.Enabled = $livre
    $todos.Enabled = $livre; $nenhum.Enabled = $livre
    $localizar.Enabled = $livre -and $projetosUI.CheckedIndices.Count -gt 0
    $iniciar.Enabled = $livre -and $ui.Areas -gt 0
    $destinoUI.Enabled = $livre; $escolher.Enabled = $livre
    $detalhado.Enabled = $livre
    $cancelar.Enabled = $ui.Ocupado -and -not $estado.Cancelar -and -not $estado.Encerrar
}
function Enviar($pedido) {
    $ui.Acao = $pedido.Acao
    $cancelar.Text = 'Cancelar'
    $estado.Cancelar = $false; $ui.Ocupado = $true; $ui.Relogio.Restart()
    $barra.Value = 0; $situacao.Text = 'Processando: ' + $pedido.Acao
    Atualizar-Botoes
    $pedidos.Enqueue($pedido)
}
$conectar.Add_Click({ Enviar @{ Acao = 'Conectar'; Datasource = $DatasourceName } })
$concessao.Add_SelectedIndexChanged({
    $projetosUI.Items.Clear(); $grid.Rows.Clear(); $ui.Areas = 0; Atualizar-Botoes
})
$carregar.Add_Click({
    $projetosUI.Items.Clear(); $grid.Rows.Clear(); $ui.Areas = 0
    Enviar @{ Acao = 'Projetos'; Indice = $concessao.SelectedIndex }
})
$projetosUI.Add_ItemCheck({ $ui.Areas = 0; $grid.Rows.Clear() })
$todos.Add_Click({ for ($i = 0; $i -lt $projetosUI.Items.Count; $i++) { $projetosUI.SetItemChecked($i, $true) } })
$nenhum.Add_Click({ for ($i = 0; $i -lt $projetosUI.Items.Count; $i++) { $projetosUI.SetItemChecked($i, $false) } })
$localizar.Add_Click({
    $ui.Areas = 0; $grid.Rows.Clear()
    Enviar @{ Acao = 'Localizar'; Indices = @($projetosUI.CheckedIndices | ForEach-Object { [int]$_ }) }
})
$escolher.Add_Click({
    $dialogo = New-Object Windows.Forms.FolderBrowserDialog
    try { if ($dialogo.ShowDialog($form) -eq 'OK') { $destinoUI.Text = $dialogo.SelectedPath } }
    finally { $dialogo.Dispose() }
})
$iniciar.Add_Click({
    if ([string]::IsNullOrWhiteSpace($destinoUI.Text)) {
        [Windows.Forms.MessageBox]::Show($form, 'Escolha uma pasta de destino.') | Out-Null
        return
    }
    foreach ($linha in $grid.Rows) { $linha.Cells[2].Value = 'Aguardando'; $linha.Cells[3].Value = ''; $linha.Cells[4].Value = '' }
    $grid.Columns[3].HeaderText = if ($detalhado.Checked) { 'Processados / total' } else { 'Retornos PW' }
    $barraArquivos.Value = 0
    Enviar @{ Acao = 'Exportar'; Destino = $destinoUI.Text; SomenteListar = [bool]$SomenteListar; Detalhado = $detalhado.Checked }
})
$cancelar.Add_Click({
    $estado.Cancelar = $true
    $cancelar.Text = 'Cancelando...'
    $situacao.Text = 'Cancelamento solicitado. Aguardando a chamada atual do ProjectWise terminar...'
    $detalhes.AppendText("Cancelamento solicitado. Os arquivos ja exportados serao mantidos.`r`n")
    Atualizar-Botoes
})
$abrir.Add_Click({
    if (Test-Path -LiteralPath $destinoUI.Text -PathType Container) { Invoke-Item -LiteralPath $destinoUI.Text }
})
$form.Add_FormClosing({
    param($sender, $e)
    if (-not $ui.Fechar) {
        $e.Cancel = $true
        $estado.Cancelar = $true; $estado.Encerrar = $true
        $ui.Ocupado = $true
        $situacao.Text = 'Encerrando: aguardando a etapa atual e o logout do ProjectWise...'
        Atualizar-Botoes
    }
})

$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = 'MTA'; $runspace.ThreadOptions = 'ReuseThread'; $runspace.Open()
$ps = [PowerShell]::Create(); $ps.Runspace = $runspace
$null = $ps.AddScript($trabalhador.ToString()).AddArgument($estado).AddArgument($pedidos).AddArgument($eventos).AddArgument((Join-Path $PSScriptRoot '12 - Exportar_Areas_de_Transferencia_ProjectWise.ps1'))
$execucao = $ps.BeginInvoke()
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    $evento = $null
    while ($eventos.TryDequeue([ref]$evento)) {
        $d = $evento.Dados
        switch ($evento.Tipo) {
            'Concessoes' { $ui.Conectado = $true; foreach ($nome in $d) { $null = $concessao.Items.Add($nome) }; if ($concessao.Items.Count) { $concessao.SelectedIndex = 0 }; $situacao.Text = 'Selecione a concessao e carregue os projetos.' }
            'Projetos' { foreach ($nome in $d) { $null = $projetosUI.Items.Add($nome) }; $situacao.Text = "$($projetosUI.Items.Count) projeto(s). Marque e localize as areas." }
            'Areas' { $ui.Areas = @($d).Count; foreach ($area in $d) { $null = $grid.Rows.Add([object[]]@($area.Projeto, $area.AreaID, 'Aguardando', '', '')) } }
            'Linha' { $linha = $grid.Rows[$d.Indice]; $linha.Cells[2].Value = $d.Status; $linha.Cells[3].Value = $d.Retornos; $linha.Cells[4].Value = $d.Duracao }
            'Progresso' { $barra.Value = $d.Valor; $situacao.Text = $d.Texto }
            'Arquivos' { $barraArquivos.Value = $d.Valor; $situacaoArquivos.Text = $d.Texto }
            'Mensagem' { $situacao.Text = [string]$d; $detalhes.AppendText("$(Get-Date -Format 'HH:mm:ss') | $d`r`n") }
            'Erro' { $situacao.Text = 'Falha: ' + $d; $detalhes.AppendText("ERRO: $d`r`n") }
            'Livre' {
                $ui.Ocupado = $false; $ui.Relogio.Stop()
                $cancelar.Text = 'Cancelar'
                if ($estado.Cancelar -and -not $estado.Encerrar) {
                    $situacao.Text = 'Operacao encerrada apos pedido de cancelamento. Resultados ja obtidos foram mantidos.'
                    if ($ui.Acao -eq 'Exportar') {
                        foreach ($linha in $grid.Rows) {
                            if ($linha.Cells[2].Value -eq 'Aguardando') { $linha.Cells[2].Value = 'Cancelada' }
                        }
                    }
                    $detalhes.AppendText("$($situacao.Text)`r`n")
                }
            }
        }
    }
    $tempo.Text = 'Tempo: ' + $ui.Relogio.Elapsed.ToString('hh\:mm\:ss')
    Atualizar-Botoes
    if ($execucao.IsCompleted -and $estado.Encerrar) { $ui.Fechar = $true; $form.Close() }
    elseif ($execucao.IsCompleted) { $situacao.Text = 'O trabalhador encerrou. Feche e reabra a interface.'; $ui.Ocupado = $true; Atualizar-Botoes; $cancelar.Enabled = $false }
})
try {
    Atualizar-Botoes
    $timer.Start()
    $null = $form.ShowDialog()
    $ps.EndInvoke($execucao)
} finally {
    $timer.Stop(); $timer.Dispose(); $ps.Dispose(); $runspace.Dispose(); $form.Dispose()
}
