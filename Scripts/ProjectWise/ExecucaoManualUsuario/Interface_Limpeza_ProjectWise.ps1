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
    param($estado, $pedidos, $eventos, $arquivoBase, $apenasListar)
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
        function Write-Etapa { param([string]$Mensagem) Emitir 'Mensagem' $Mensagem }
        $concessoes = @(); $projetos = @(); $areas = @(); $alvos = @(); $analiseValida = $false
        while (-not $estado.Encerrar) {
            $pedido = $null
            if (-not $pedidos.TryDequeue([ref]$pedido)) { Start-Sleep -Milliseconds 100; continue }
            try {
                switch ($pedido.Acao) {
                    'Conectar' {
                        Emitir 'Mensagem' 'Interface Limpeza v1.2 | Preparando PWPS_DAB...'
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
                        $areas = @(); $alvos = @(); $analiseValida = $false
                        $concessao = $concessoes[$pedido.Indice]
                        $nomeConcessao = Obter-RotuloPasta $concessao
                        $pasta = Localizar-Pasta (Obter-FilhosPasta (Obter-IdPasta $concessao)) @('Projetos','Projeto','Projects','Project') 'Projetos'
                        $projetos = @(Obter-FilhosPasta (Obter-IdPasta $pasta) | Sort-Object { Obter-RotuloPasta $_ })
                        Emitir 'Projetos' @($projetos | ForEach-Object { Obter-RotuloPasta $_ })
                    }
                    'Localizar' {
                        $alvos = @(); $analiseValida = $false
                        $lista = New-Object 'System.Collections.Generic.List[object]'
                        $vistos = New-Object 'System.Collections.Generic.HashSet[int]'
                        $analisados = 0; $pastasVazias = 0
                        foreach ($indice in $pedido.Indices) {
                            if ($estado.Cancelar) { break }
                            $projeto = $projetos[$indice]
                            foreach ($area in (Localizar-AreasTransferenciaNoProjeto $projeto)) {
                                $idArea = Obter-IdPastaValido $area
                                foreach ($pasta in (Obter-FilhosPasta $idArea | Where-Object { Testar-PastaLimpeza $_ })) {
                                    if ($estado.Cancelar) { break }
                                    $idPasta = Obter-IdPastaValido $pasta
                                    if ($idPasta -eq $idArea) { throw 'Alvo corresponde a pasta protegida.' }
                                    if (-not $vistos.Add($idPasta)) { continue }
                                    Emitir 'Mensagem' "Analisando $(Obter-RotuloPasta $projeto) / $(Obter-NomePasta $pasta)..."
                                    $inventario = Obter-InventarioTamanho $idPasta
                                    $filhos = @(Obter-FilhosPasta $idPasta)
                                    if ($inventario.Count -eq 0 -and $filhos.Count -eq 0) {
                                        $pastasVazias++
                                        Emitir 'Mensagem' "Ja vazia: $(Obter-RotuloPasta $projeto) / $(Obter-NomePasta $pasta) (ID $idPasta). Fora da fila de limpeza."
                                        continue
                                    }
                                    $medida = Medir-Exclusao -Antes $inventario -Depois @{}
                                    $lista.Add([pscustomobject]@{
                                        Projeto = Obter-RotuloPasta $projeto; AreaID = $idArea; PastaID = $idPasta
                                        Pasta = Obter-NomePasta $pasta; CaminhoPW = Obter-CaminhoPW $pasta
                                        DocumentosComVersoes = $inventario.Count; SubpastasDiretas = $filhos.Count
                                        MiB = $medida.MiBExcluidos
                                    })
                                }
                            }
                            $analisados++
                            Emitir 'Progresso' @{ Valor = [int](100 * $analisados / $pedido.Indices.Count); Texto = "$analisados/$($pedido.Indices.Count) projetos analisados" }
                        }
                        if ($estado.Cancelar) { Emitir 'Mensagem' 'Analise cancelada. Analise novamente antes de limpar.'; break }
                        $alvos = $lista.ToArray(); $analiseValida = $true
                        Emitir 'Areas' $alvos
                        Emitir 'Mensagem' "Analise concluida: $($alvos.Count) pasta(s) com conteudo para limpar; $pastasVazias pasta(s) ja vazia(s), fora da fila."
                    }
                    'Limpar' {
                        if ($apenasListar) { throw 'SomenteListar: exclusao bloqueada.' }
                        if (-not $analiseValida -or -not $alvos.Count -or -not $pedido.Confirmado) { throw 'Analise e confirmacao necessarias antes da limpeza.' }
                        $analiseValida = $false
                        $pastaLog = [IO.Path]::GetFullPath($pedido.Destino)
                        $null = [IO.Directory]::CreateDirectory($pastaLog)
                        $sufixo = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
                        $arquivoLog = Join-Path $pastaLog "Log_Limpeza_$sufixo.csv"
                        $alvos | Export-Csv -LiteralPath (Join-Path $pastaLog "Alvos_Limpeza_$sufixo.csv") -NoTypeInformation -Encoding UTF8
                        # Garante acesso de gravacao antes de qualquer exclusao.
                        'Projeto,PastaID,Status' | Set-Content -LiteralPath $arquivoLog -Encoding UTF8
                        $log = New-Object 'System.Collections.Generic.List[object]'
                        $falhas = 0; $processadas = 0
                        [decimal]$totalBytes = 0; $totalRemovidos = 0
                        for ($i = 0; $i -lt $alvos.Count; $i++) {
                            if ($estado.Cancelar) { break }
                            $alvo = $alvos[$i]
                            $interromperFila = $false
                            $status = 'Concluido'; $erro = ''; $antes = $null; $medicao = $null
                            $statusMedicao = 'Nao confirmado'; $erroMedicao = ''
                            $linhaErro = $null; $pilhaErro = ''; $identificadorErro = ''
                            $diagnostico = [pscustomobject]@{ Etapa = 'Inventario anterior'; ItensRetornados = $null; Item = '' }
                            $inicio = Get-Date; $relogio = [Diagnostics.Stopwatch]::StartNew()
                            Emitir 'Linha' @{ Indice = $i; Status = 'Limpando'; Retornos = ''; Duracao = ''; MiB = '' }
                            try {
                                Emitir 'Mensagem' "Limpando $($alvo.Projeto) / $($alvo.Pasta) (ID $($alvo.PastaID))..."
                                $antes = Obter-InventarioTamanho $alvo.PastaID
                                if ($estado.Cancelar) { $status = 'Cancelada' }
                                else { Limpar-ConteudoPasta -Alvo $alvo -Diagnostico $diagnostico }
                            } catch {
                                $interromperFila = $true
                                $status = 'Erro'; $erro = $_.Exception.Message; $falhas++
                                $linhaErro = $_.InvocationInfo.ScriptLineNumber; $pilhaErro = $_.ScriptStackTrace
                                $identificadorErro = $_.FullyQualifiedErrorId
                                Emitir 'Mensagem' "Falha em $($alvo.PastaID), etapa $($diagnostico.Etapa): $erro"
                            }
                            if ($null -ne $antes) {
                                try {
                                    Confirmar-PastaProtegida $alvo
                                    $depois = Obter-InventarioTamanho $alvo.PastaID
                                    $medicao = Medir-Exclusao -Antes $antes -Depois $depois
                                    $statusMedicao = 'Confirmado por inventario'
                                    $totalBytes += $medicao.BytesExcluidos
                                    $totalRemovidos += $medicao.DocumentosVersoesExcluidos
                                } catch {
                                    $erroMedicao = $_.Exception.Message; $interromperFila = $true
                                    if ($status -ne 'Erro') { $status = 'Erro'; $falhas++ }
                                    Emitir 'Mensagem' "Medicao indisponivel: $erroMedicao"
                                }
                            } else { $erroMedicao = 'Inventario anterior indisponivel. Exclusao nao iniciada.' }
                            $relogio.Stop()
                            $log.Add([pscustomobject]@{
                                Projeto = $alvo.Projeto; AreaID = $alvo.AreaID; PastaID = $alvo.PastaID
                                Pasta = $alvo.Pasta; CaminhoPW = $alvo.CaminhoPW; Status = $status; Erro = $erro
                                Etapa = $diagnostico.Etapa; ItensSubpastasRetornados = $diagnostico.ItensRetornados
                                ItemDiagnostico = $diagnostico.Item; LinhaErro = $linhaErro
                                IdentificadorErro = $identificadorErro; PilhaErro = $pilhaErro
                                Inicio = $inicio; Fim = Get-Date; DuracaoSegundos = $relogio.Elapsed.TotalSeconds
                                BytesAntes = if ($medicao) { $medicao.BytesAntes } else { $null }
                                BytesDepois = if ($medicao) { $medicao.BytesDepois } else { $null }
                                BytesExcluidos = if ($medicao) { $medicao.BytesExcluidos } else { $null }
                                MiBExcluidos = if ($medicao) { $medicao.MiBExcluidos } else { $null }
                                GiBExcluidos = if ($medicao) { $medicao.GiBExcluidos } else { $null }
                                DocumentosVersoesExcluidos = if ($medicao) { $medicao.DocumentosVersoesExcluidos } else { $null }
                                StatusMedicao = $statusMedicao; ErroMedicao = $erroMedicao
                            })
                            $log.ToArray() | Export-Csv -LiteralPath $arquivoLog -NoTypeInformation -Encoding UTF8
                            if ($status -ne 'Cancelada') { $processadas++ }
                            Emitir 'Linha' @{ Indice = $i; Status = $status; Duracao = $relogio.Elapsed.ToString('hh\:mm\:ss')
                                Retornos = $(if ($medicao) { $medicao.DocumentosVersoesExcluidos } else { 'Nao confirmado' })
                                MiB = $(if ($medicao) { $medicao.MiBExcluidos } else { 'Nao confirmado' }) }
                            Emitir 'Progresso' @{ Valor = [int](100 * $processadas / $alvos.Count); Texto = "$processadas/$($alvos.Count) pastas processadas | $falhas falha(s)" }
                            Emitir 'Resumo' "Medido por inventario: $totalRemovidos documentos/versoes removidos | $([math]::Round($totalBytes / 1MB, 2)) MiB logicos"
                            if ($interromperFila) {
                                for ($j = $i + 1; $j -lt $alvos.Count; $j++) {
                                    $pendente = $alvos[$j]
                                    $registroPendente = [ordered]@{}
                                    foreach ($propriedade in $log[0].PSObject.Properties) { $registroPendente[$propriedade.Name] = $null }
                                    $registroPendente.Projeto = $pendente.Projeto; $registroPendente.AreaID = $pendente.AreaID
                                    $registroPendente.PastaID = $pendente.PastaID; $registroPendente.Pasta = $pendente.Pasta
                                    $registroPendente.CaminhoPW = $pendente.CaminhoPW; $registroPendente.Status = 'Nao executado'
                                    $registroPendente.Erro = "Fila interrompida apos falha na pasta $($alvo.PastaID)."
                                    $registroPendente.StatusMedicao = 'Nao realizado'
                                    $log.Add([pscustomobject]$registroPendente)
                                    Emitir 'Linha' @{ Indice = $j; Status = 'Nao executado'; Retornos = ''; MiB = ''; Duracao = '' }
                                }
                                $log.ToArray() | Export-Csv -LiteralPath $arquivoLog -NoTypeInformation -Encoding UTF8
                                Emitir 'Erro' 'Fila interrompida na primeira falha. Consulte o diagnostico e realize nova analise antes de limpar novamente.'
                                break
                            }
                        }
                        Emitir 'Mensagem' "Limpeza encerrada: $processadas/$($alvos.Count) pastas processadas; $falhas falha(s). Log: $arquivoLog"
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
$form.Text = 'ProjectWise | Limpar Areas de Transferencia | Limpeza v1.2'
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
$titulo = Controle Label 'LIMPAR AREAS DE TRANSFERENCIA'
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
$localizar = Controle Button 'Analisar limpeza'
Barra-Botoes @($todos, $nenhum, $localizar)
$rotuloDestino = Controle Label 'Pasta dos relatorios:'
$rotuloDestino.AutoSize = $true
Adicionar-Linha $rotuloDestino
if (-not $Destino) { $Destino = Join-Path $PSScriptRoot 'Logs' }
$destinoUI = Controle TextBox $Destino
Adicionar-Linha $destinoUI
$escolher = Controle Button 'Escolher pasta dos relatorios...'
$abrir = Controle Button 'Abrir relatorios'
Barra-Botoes @($escolher, $abrir)
$iniciar = Controle Button 'Limpar conteudo listado'
$cancelar = Controle Button 'Cancelar'
$detalhado = Controle CheckBox 'Somente analisar (sem excluir)'
$detalhado.Checked = [bool]$SomenteListar
Barra-Botoes @($iniciar, $cancelar, $detalhado)
$grid = Controle DataGridView ''
$grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false; $grid.AutoSizeColumnsMode = 'Fill'
$grid.DefaultCellStyle.WrapMode = 'True'
$grid.AutoSizeRowsMode = 'AllCells'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.ScrollBars = 'Both'
foreach ($coluna in @('Projeto','Pasta','Estado','Docs/versoes previstos','Subpastas diretas','Removidos','MiB removidos','Duracao','Caminho PW')) { $null = $grid.Columns.Add($coluna, $coluna) }
$grid.Columns[0].FillWeight = 300
foreach ($coluna in $grid.Columns) { $coluna.MinimumWidth = 95; $coluna.SortMode = 'NotSortable' }
$grid.Columns[0].MinimumWidth = 230
$grid.Columns[8].MinimumWidth = 280
$grid.Columns[8].FillWeight = 300
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
$barraArquivos.Height = 1; $barraArquivos.Visible = $false
Adicionar-Linha $barraArquivos
$situacaoArquivos = Controle Label 'A medicao sera atualizada apos cada pasta, inclusive em falhas parciais.'
$situacaoArquivos.AutoSize = $true
Adicionar-Linha $situacaoArquivos
$nota = Controle Label 'Entrada, Saida e Importacao sao preservadas. A limpeza exclui documentos, versoes e subpastas. MiB sao tamanhos logicos; alteracoes concorrentes podem afetar a medicao.'
$nota.AutoSize = $true
Adicionar-Linha $nota
$detalhes = Controle TextBox ''
$detalhes.Multiline = $true; $detalhes.ReadOnly = $true; $detalhes.ScrollBars = 'Vertical'; $detalhes.WordWrap = $true
Adicionar-Linha $detalhes
$detalhes.AppendText("Interface Limpeza v1.2 | Arquivo: $PSCommandPath`r`n")
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
    $iniciar.Enabled = $livre -and $ui.Areas -gt 0 -and -not $detalhado.Checked -and -not $SomenteListar
    $destinoUI.Enabled = $livre; $escolher.Enabled = $livre
    $detalhado.Enabled = $livre -and -not $SomenteListar
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
        [Windows.Forms.MessageBox]::Show($form, 'Escolha a pasta dos relatorios.') | Out-Null
        return
    }
    $quantidadeDocs = 0
    foreach ($linha in $grid.Rows) { $quantidadeDocs += [int]$linha.Cells[3].Value }
    $aviso = "Excluir o conteudo das $($grid.Rows.Count) pastas listadas?`r`n`r`nInventario: $quantidadeDocs documentos/versoes, alem das subpastas.`r`nEntrada, Saida e Importacao serao preservadas.`r`nA limpeza abrange o conteudo presente no momento da execucao.`r`n`r`nConfira os caminhos na tabela antes de confirmar."
    if ([Windows.Forms.MessageBox]::Show($form, $aviso, 'Confirmar limpeza no ProjectWise', 'YesNo', 'Warning', 'Button2') -ne 'Yes') { return }
    foreach ($linha in $grid.Rows) { $linha.Cells[2].Value = 'Aguardando' }
    Enviar @{ Acao = 'Limpar'; Destino = $destinoUI.Text; Confirmado = $true }
})
$cancelar.Add_Click({
    $estado.Cancelar = $true
    $cancelar.Text = 'Cancelando...'
    $situacao.Text = 'Cancelamento solicitado. Aguardando a chamada atual do ProjectWise terminar...'
    $detalhes.AppendText("Cancelamento solicitado. A limpeza atual terminara antes da parada; exclusoes realizadas nao serao desfeitas.`r`n")
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
$null = $ps.AddScript($trabalhador.ToString()).AddArgument($estado).AddArgument($pedidos).AddArgument($eventos).AddArgument((Join-Path $PSScriptRoot '15 - Limpar_Areas_de_Transferencia_ProjectWise.ps1')).AddArgument([bool]$SomenteListar)
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
            'Areas' { $ui.Areas = @($d).Count; foreach ($area in $d) { $null = $grid.Rows.Add([object[]]@($area.Projeto, $area.Pasta, 'Analisada', $area.DocumentosComVersoes, $area.SubpastasDiretas, '', '', '', $area.CaminhoPW)) } }
            'Linha' { $linha = $grid.Rows[$d.Indice]; $linha.Cells[2].Value = $d.Status; $linha.Cells[5].Value = $d.Retornos; $linha.Cells[6].Value = $d.MiB; $linha.Cells[7].Value = $d.Duracao }
            'Progresso' { $barra.Value = $d.Valor; $situacao.Text = $d.Texto }
            'Resumo' { $situacaoArquivos.Text = [string]$d }
            'Mensagem' { $situacao.Text = [string]$d; $detalhes.AppendText("$(Get-Date -Format 'HH:mm:ss') | $d`r`n") }
            'Erro' { $situacao.Text = 'Falha: ' + $d; $detalhes.AppendText("ERRO: $d`r`n") }
            'Livre' {
                $ui.Ocupado = $false; $ui.Relogio.Stop()
                $cancelar.Text = 'Cancelar'
                if ($ui.Acao -eq 'Limpar') { $ui.Areas = 0 }
                if ($estado.Cancelar -and -not $estado.Encerrar) {
                    $situacao.Text = 'Operacao encerrada apos pedido de cancelamento. Resultados ja obtidos foram mantidos.'
                    if ($ui.Acao -eq 'Limpar') {
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
