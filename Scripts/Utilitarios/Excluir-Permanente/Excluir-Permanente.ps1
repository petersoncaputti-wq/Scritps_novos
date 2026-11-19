<#
.SYNOPSIS
    Exclui permanentemente o CONTEUDO de uma pasta (equivalente a Shift+Del: nao vai para a Lixeira).
    A pasta selecionada e sempre mantida; somente arquivos e subpastas dentro dela sao apagados.

.EXAMPLE
    .\Excluir-Permanente.ps1
    Abre a tela para selecionar, analisar e excluir o conteudo da pasta.

.EXAMPLE
    .\Excluir-Permanente.ps1 -Pasta 'D:\Temp\Exportacao'
    Abre a tela ja com a pasta selecionada e analisada.

.EXAMPLE
    .\Excluir-Permanente.ps1 -Pasta 'D:\Temp\Exportacao' -SemConfirmar
    Sem tela (automacoes): exclui o conteudo direto e mantem a pasta.

.NOTES
    Salvo em UTF-8 com BOM para o Windows PowerShell 5.1 exibir os acentos corretamente.
    Analise e exclusao rodam em segundo plano para a janela nao travar em pastas grandes.
#>
param(
    [string]$Pasta,
    [switch]$SemConfirmar
)
$ErrorActionPreference = 'Stop'
# Quantidade maxima de linhas exibidas na lista (a analise e a exclusao consideram tudo).
$LimiteLista = 3000
$CaminhoScript = $PSCommandPath

function Test-Administrador {
    $identidade = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal $identidade).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ComoAdministrador([string]$pasta) {
    # Reabre este script elevado (UAC). Retorna $false se o usuario recusar.
    $argumentos = '-NoProfile -STA -ExecutionPolicy Bypass -File "{0}"' -f $CaminhoScript
    if ($pasta) { $argumentos += ' -Pasta "{0}"' -f $pasta }
    try {
        Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList $argumentos | Out-Null
        return $true
    } catch {
        return $false
    }
}

function ConvertTo-CaminhoLongo([string]$caminho) {
    # Prefixo \\?\ permite trabalhar com caminhos acima de 260 caracteres.
    if ($caminho.StartsWith('\\?\')) { return $caminho }
    if ($caminho.StartsWith('\\')) { return '\\?\UNC\' + $caminho.Substring(2) }
    return '\\?\' + $caminho
}

function ConvertFrom-CaminhoLongo([string]$caminho) {
    if ($caminho.StartsWith('\\?\UNC\')) { return '\\' + $caminho.Substring(8) }
    if ($caminho.StartsWith('\\?\')) { return $caminho.Substring(4) }
    return $caminho
}

function Get-MensagemErro($registro) {
    $excecao = $registro.Exception
    if ($excecao.InnerException) { $excecao = $excecao.InnerException }
    return $excecao.Message
}

function Format-Tamanho([double]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N2} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N1} MB' -f ($bytes / 1MB) }
    if ($bytes -ge 1KB) { return '{0:N0} KB' -f ($bytes / 1KB) }
    return "$bytes bytes"
}

function Get-MotivoBloqueio([string]$caminho) {
    # Retorna o motivo se a pasta nao puder ter o conteudo excluido; $null se permitido.
    if (-not $caminho) { return 'Nenhuma pasta selecionada.' }
    if (-not (Test-Path -LiteralPath $caminho -PathType Container)) { return "Pasta não encontrada: $caminho" }
    $normalizado = [IO.Path]::GetFullPath($caminho).TrimEnd('\')
    if ($normalizado -match '^[A-Za-z]:$' -or $normalizado -match '^\\\\[^\\]+\\[^\\]+$') {
        return 'Por segurança, não é permitido limpar a raiz de uma unidade ou compartilhamento.'
    }
    $protegidas = @(
        $env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData,
        $env:USERPROFILE, (Split-Path $env:USERPROFILE -Parent), $env:APPDATA, $env:LOCALAPPDATA,
        [Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'),
        $env:OneDrive, $env:OneDriveCommercial
    ) | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') }
    if ($protegidas -contains $normalizado) { return "Por segurança, esta pasta do sistema/usuário não pode ser limpa: $normalizado" }
    $pastaScript = $PSScriptRoot.TrimEnd('\')
    if ($pastaScript -eq $normalizado -or $pastaScript.StartsWith($normalizado + '\')) {
        return 'A pasta selecionada contém este próprio script. Escolha outra pasta.'
    }
    return $null
}

function Test-PastaOneDrive([string]$caminho) {
    foreach ($raiz in @($env:OneDrive, $env:OneDriveCommercial)) {
        if ($raiz -and $caminho.StartsWith($raiz.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# ---------- Funcoes de trabalho (rodam em segundo plano; $sync e compartilhado com a tela) ----------

function Measure-Arvore([string]$pasta, $sync) {
    # Conta arquivos/subpastas/bytes abaixo de $pasta. Nao entra em links (junction/symlink).
    $resultado = @{ Arquivos = 0; Pastas = 0; Bytes = [int64]0; SemAcesso = 0 }
    $pilha = New-Object 'System.Collections.Generic.Stack[string]'
    $pilha.Push($pasta)
    while ($pilha.Count -gt 0) {
        if ($sync.Cancelar) { break }
        $atual = $pilha.Pop()
        try { $entradas = (New-Object IO.DirectoryInfo $atual).GetFileSystemInfos() } catch { $resultado.SemAcesso++; continue }
        foreach ($entrada in $entradas) {
            if ($entrada -is [IO.DirectoryInfo]) {
                $resultado.Pastas++
                if (-not ($entrada.Attributes -band [IO.FileAttributes]::ReparsePoint)) { $pilha.Push($entrada.FullName) }
            } else {
                $resultado.Arquivos++
                $resultado.Bytes += $entrada.Length
            }
        }
    }
    return $resultado
}

function Get-AnaliseConteudo([string]$raiz, $sync) {
    $linhasPastas = New-Object 'System.Collections.Generic.List[object]'
    $linhasArquivos = New-Object 'System.Collections.Generic.List[object]'
    $total = @{ Itens = 0; Arquivos = 0; Pastas = 0; Bytes = [int64]0; SemAcesso = 0; Cancelado = $false; Linhas = $null }
    foreach ($entrada in (New-Object IO.DirectoryInfo $raiz).EnumerateFileSystemInfos()) {
        if ($sync.Cancelar) { break }
        $total.Itens++
        $sync.Processados = $total.Itens
        if ($entrada -is [IO.DirectoryInfo]) {
            $total.Pastas++
            $sync.Atual = $entrada.Name
            if ($entrada.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                $linhasPastas.Add([pscustomobject]@{ Nome = $entrada.Name; Tipo = 'Atalho de pasta'; Bytes = [int64]0; Arquivos = $null; Pastas = $null })
            } else {
                $medida = Measure-Arvore $entrada.FullName $sync
                $total.Arquivos += $medida.Arquivos; $total.Pastas += $medida.Pastas
                $total.Bytes += $medida.Bytes; $total.SemAcesso += $medida.SemAcesso
                $linhasPastas.Add([pscustomobject]@{ Nome = $entrada.Name; Tipo = 'Pasta'; Bytes = $medida.Bytes; Arquivos = $medida.Arquivos; Pastas = $medida.Pastas })
            }
        } else {
            $total.Arquivos++
            $total.Bytes += $entrada.Length
            $linhasArquivos.Add([pscustomobject]@{ Nome = $entrada.Name; Tipo = 'Arquivo'; Bytes = $entrada.Length; Arquivos = $null; Pastas = $null })
        }
    }
    if ($sync.Cancelar) { $total.Cancelado = $true; return $total }
    $sync.Atual = 'organizando a lista'
    $ordenadas = New-Object 'System.Collections.Generic.List[object]'
    foreach ($linha in ($linhasPastas | Sort-Object Nome)) { $ordenadas.Add($linha) }
    foreach ($linha in ($linhasArquivos | Sort-Object Nome)) { $ordenadas.Add($linha) }
    $total.Linhas = $ordenadas
    return $total
}

function Remove-ConteudoPasta([string]$raiz, $sync) {
    # Exclui tudo DENTRO de $raiz, sem Lixeira. A propria $raiz nunca e excluida.
    $falhas = New-Object 'System.Collections.Generic.List[string]'
    $caminhosComFalha = New-Object 'System.Collections.Generic.List[string]'
    $pastas = New-Object 'System.Collections.Generic.List[string]'
    $pilha = New-Object 'System.Collections.Generic.Stack[string]'
    $arquivosExcluidos = 0
    $pastasExcluidas = 0
    $semPermissao = 0
    $pilha.Push($raiz)
    while ($pilha.Count -gt 0) {
        if ($sync.Cancelar) { break }
        $atual = $pilha.Pop()
        $pastas.Add($atual)
        try { $entradas = (New-Object IO.DirectoryInfo $atual).GetFileSystemInfos() } catch {
            $falhas.Add((ConvertFrom-CaminhoLongo $atual) + ' -> ' + (Get-MensagemErro $_))
            $caminhosComFalha.Add($atual)
            continue
        }
        foreach ($entrada in $entradas) {
            if ($sync.Cancelar) { break }
            try {
                if ($entrada -is [IO.DirectoryInfo]) {
                    if ($entrada.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                        # Remove apenas o atalho; o conteudo do destino nao e tocado.
                        [IO.Directory]::Delete($entrada.FullName, $false)
                        $pastasExcluidas++
                    } else {
                        $pilha.Push($entrada.FullName)
                    }
                } else {
                    $sync.Atual = $entrada.Name
                    if ($entrada.Attributes -band [IO.FileAttributes]::ReadOnly) { $entrada.Attributes = [IO.FileAttributes]::Normal }
                    $entrada.Delete()
                    $arquivosExcluidos++
                    $sync.Processados = $arquivosExcluidos
                }
            } catch {
                $excecao = $_.Exception
                if ($excecao.InnerException) { $excecao = $excecao.InnerException }
                if ($excecao -is [UnauthorizedAccessException]) { $semPermissao++ }
                $falhas.Add((ConvertFrom-CaminhoLongo $entrada.FullName) + ' -> ' + $excecao.Message)
                $caminhosComFalha.Add($entrada.FullName)
            }
        }
    }
    $cancelado = [bool]$sync.Cancelar
    if (-not $cancelado) {
        # Subpastas de baixo para cima; indice 0 e a pasta raiz, que e mantida.
        for ($i = $pastas.Count - 1; $i -ge 1; $i--) {
            $pastaAtual = $pastas[$i]
            $prefixo = $pastaAtual + '\'
            if ($caminhosComFalha | Where-Object { $_.StartsWith($prefixo) -or $_ -eq $pastaAtual }) { continue }
            try {
                $info = New-Object IO.DirectoryInfo $pastaAtual
                if ($info.Attributes -band [IO.FileAttributes]::ReadOnly) { $info.Attributes = [IO.FileAttributes]::Directory }
                [IO.Directory]::Delete($pastaAtual, $false)
                $pastasExcluidas++
            } catch {
                $falhas.Add((ConvertFrom-CaminhoLongo $pastaAtual) + ' -> ' + (Get-MensagemErro $_))
                $caminhosComFalha.Add($pastaAtual)
            }
        }
    }
    return @{ Arquivos = $arquivosExcluidos; Pastas = $pastasExcluidas; Falhas = $falhas; Cancelado = $cancelado; SemPermissao = $semPermissao }
}

# ---------- Tela ----------

function Show-Interface([string]$pastaInicial) {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $estado = @{ Pasta = $null; TotalArquivos = 0; TotalPastas = 0; TotalItens = 0; Ocupado = $false; FecharAoTerminar = $false; Tipo = $null }
    $sync = [hashtable]::Synchronized(@{ Cancelar = $false; Processados = 0; Atual = ''; Resultado = $null; Erro = $null })

    # Ambiente de segundo plano com as funcoes de trabalho.
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    foreach ($nome in @('ConvertFrom-CaminhoLongo', 'Get-MensagemErro', 'Measure-Arvore', 'Get-AnaliseConteudo', 'Remove-ConteudoPasta')) {
        $definicao = (Get-Item "function:$nome").Definition
        $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $nome, $definicao))
    }
    $scriptTarefa = {
        param($sync, $tipo, $raiz)
        $ErrorActionPreference = 'Stop'
        try {
            if ($tipo -eq 'Analise') { $sync.Resultado = Get-AnaliseConteudo $raiz $sync }
            else { $sync.Resultado = Remove-ConteudoPasta $raiz $sync }
        } catch {
            $sync.Erro = Get-MensagemErro $_
        }
    }.ToString()

    $form = New-Object System.Windows.Forms.Form
    $ehAdministrador = Test-Administrador
    $form.Text = 'Excluir conteúdo de pasta (permanente)'
    if ($ehAdministrador) { $form.Text += ' — Administrador' }
    $form.ClientSize = New-Object System.Drawing.Size(844, 570)
    $form.MinimumSize = New-Object System.Drawing.Size(740, 480)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $lblPasta = New-Object System.Windows.Forms.Label
    $lblPasta.Text = 'Pasta a analisar:'
    $lblPasta.Location = New-Object System.Drawing.Point(12, 12)
    $lblPasta.AutoSize = $true

    $txtPasta = New-Object System.Windows.Forms.TextBox
    $txtPasta.Location = New-Object System.Drawing.Point(12, 34)
    $txtPasta.Size = New-Object System.Drawing.Size(674, 25)
    $txtPasta.ReadOnly = $true
    $txtPasta.Anchor = 'Top, Left, Right'

    $btnSelecionar = New-Object System.Windows.Forms.Button
    $btnSelecionar.Text = 'Selecionar pasta...'
    $btnSelecionar.Location = New-Object System.Drawing.Point(694, 32)
    $btnSelecionar.Size = New-Object System.Drawing.Size(138, 28)
    $btnSelecionar.Anchor = 'Top, Right'

    $lblAviso = New-Object System.Windows.Forms.Label
    $lblAviso.Text = 'A pasta selecionada NÃO será apagada. Somente o conteúdo dentro dela (arquivos e subpastas) será excluído permanentemente, sem passar pela Lixeira.'
    $lblAviso.Location = New-Object System.Drawing.Point(12, 68)
    $lblAviso.Size = New-Object System.Drawing.Size(820, 36)
    $lblAviso.Anchor = 'Top, Left, Right'
    $lblAviso.ForeColor = [System.Drawing.Color]::DarkRed
    $lblAviso.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    $lblResumo = New-Object System.Windows.Forms.Label
    $lblResumo.Text = 'Selecione uma pasta para analisar o conteúdo.'
    $lblResumo.Location = New-Object System.Drawing.Point(12, 108)
    $lblResumo.Size = New-Object System.Drawing.Size(820, 20)
    $lblResumo.Anchor = 'Top, Left, Right'

    $lista = New-Object System.Windows.Forms.ListView
    $lista.Location = New-Object System.Drawing.Point(12, 132)
    $lista.Size = New-Object System.Drawing.Size(820, 330)
    $lista.Anchor = 'Top, Bottom, Left, Right'
    $lista.View = 'Details'
    $lista.FullRowSelect = $true
    $lista.GridLines = $true
    [void]$lista.Columns.Add('Nome', 360)
    [void]$lista.Columns.Add('Tipo', 110)
    [void]$lista.Columns.Add('Tamanho', 110, 'Right')
    [void]$lista.Columns.Add('Arquivos', 100, 'Right')
    [void]$lista.Columns.Add('Subpastas', 100, 'Right')

    $progresso = New-Object System.Windows.Forms.ProgressBar
    $progresso.Location = New-Object System.Drawing.Point(12, 472)
    $progresso.Size = New-Object System.Drawing.Size(820, 18)
    $progresso.Anchor = 'Bottom, Left, Right'
    $progresso.MarqueeAnimationSpeed = 30

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 494)
    $lblStatus.Size = New-Object System.Drawing.Size(820, 20)
    $lblStatus.Anchor = 'Bottom, Left, Right'
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.AutoEllipsis = $true

    $btnAnalisar = New-Object System.Windows.Forms.Button
    $btnAnalisar.Text = 'Analisar novamente'
    $btnAnalisar.Location = New-Object System.Drawing.Point(12, 526)
    $btnAnalisar.Size = New-Object System.Drawing.Size(140, 32)
    $btnAnalisar.Anchor = 'Bottom, Left'
    $btnAnalisar.Enabled = $false

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = 'Cancelar'
    $btnCancelar.Location = New-Object System.Drawing.Point(160, 526)
    $btnCancelar.Size = New-Object System.Drawing.Size(110, 32)
    $btnCancelar.Anchor = 'Bottom, Left'
    $btnCancelar.Enabled = $false

    $btnAdmin = New-Object System.Windows.Forms.Button
    $btnAdmin.Text = 'Executar como administrador'
    $btnAdmin.Location = New-Object System.Drawing.Point(278, 526)
    $btnAdmin.Size = New-Object System.Drawing.Size(190, 32)
    $btnAdmin.Anchor = 'Bottom, Left'
    $btnAdmin.Visible = -not $ehAdministrador

    $btnExcluir = New-Object System.Windows.Forms.Button
    $btnExcluir.Text = 'Excluir conteúdo'
    $btnExcluir.Location = New-Object System.Drawing.Point(566, 526)
    $btnExcluir.Size = New-Object System.Drawing.Size(150, 32)
    $btnExcluir.Anchor = 'Bottom, Right'
    $btnExcluir.BackColor = [System.Drawing.Color]::Firebrick
    $btnExcluir.ForeColor = [System.Drawing.Color]::White
    $btnExcluir.FlatStyle = 'Flat'
    $btnExcluir.Enabled = $false

    $btnFechar = New-Object System.Windows.Forms.Button
    $btnFechar.Text = 'Fechar'
    $btnFechar.Location = New-Object System.Drawing.Point(724, 526)
    $btnFechar.Size = New-Object System.Drawing.Size(108, 32)
    $btnFechar.Anchor = 'Bottom, Right'

    $form.Controls.AddRange(@($lblPasta, $txtPasta, $btnSelecionar, $lblAviso, $lblResumo, $lista, $progresso, $lblStatus, $btnAnalisar, $btnCancelar, $btnAdmin, $btnExcluir, $btnFechar))

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 300

    $definirOcupado = {
        param([bool]$ocupado)
        $estado.Ocupado = $ocupado
        $btnSelecionar.Enabled = -not $ocupado
        $btnAnalisar.Enabled = (-not $ocupado) -and [bool]$estado.Pasta
        $btnFechar.Enabled = -not $ocupado
        $btnCancelar.Enabled = $ocupado
        $btnAdmin.Enabled = -not $ocupado
        $btnExcluir.Enabled = (-not $ocupado) -and $estado.TotalItens -gt 0
        if ($ocupado) { $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting } else { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }

    $iniciarTarefa = {
        param([string]$tipo)
        $sync.Cancelar = $false; $sync.Processados = 0; $sync.Atual = ''; $sync.Resultado = $null; $sync.Erro = $null
        $estado.Tipo = $tipo
        $estado.Inicio = Get-Date
        & $definirOcupado $true
        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
        $runspace.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript($scriptTarefa).AddArgument($sync).AddArgument($tipo).AddArgument((ConvertTo-CaminhoLongo $estado.Pasta))
        $estado.Runspace = $runspace
        $estado.Ps = $ps
        $estado.Handle = $ps.BeginInvoke()
        $timer.Start()
    }

    $mostrarAnalise = {
        $res = $sync.Resultado
        $lista.BeginUpdate()
        $lista.Items.Clear()
        $progresso.Style = 'Blocks'
        $progresso.Value = 0
        $estado.TotalItens = 0
        if ($sync.Erro -or -not $res) {
            $lblResumo.Text = 'Não foi possível ler a pasta: ' + $sync.Erro
            $lblStatus.Text = ''
        } elseif ($res.Cancelado) {
            $lblResumo.Text = 'Análise cancelada. Clique em "Analisar novamente" para refazer.'
            $lblStatus.Text = ''
        } else {
            $exibir = [Math]::Min($res.Linhas.Count, $LimiteLista)
            $itens = New-Object 'System.Windows.Forms.ListViewItem[]' $exibir
            for ($i = 0; $i -lt $exibir; $i++) {
                $linha = $res.Linhas[$i]
                $item = New-Object System.Windows.Forms.ListViewItem([string]$linha.Nome)
                [void]$item.SubItems.Add([string]$linha.Tipo)
                [void]$item.SubItems.Add((Format-Tamanho $linha.Bytes))
                if ($linha.Tipo -eq 'Pasta') {
                    [void]$item.SubItems.Add(([int64]$linha.Arquivos).ToString('N0'))
                    [void]$item.SubItems.Add(([int64]$linha.Pastas).ToString('N0'))
                } else {
                    [void]$item.SubItems.Add('')
                    [void]$item.SubItems.Add('')
                }
                $itens[$i] = $item
            }
            $lista.Items.AddRange($itens)
            $estado.TotalItens = $res.Itens
            $estado.TotalArquivos = $res.Arquivos
            $estado.TotalPastas = $res.Pastas
            if ($res.Itens -eq 0) {
                $lblResumo.Text = 'A pasta está vazia. Não há nada para excluir.'
            } else {
                $lblResumo.Text = 'Conteúdo: {0:N0} arquivo(s), {1:N0} subpasta(s), {2}' -f $res.Arquivos, $res.Pastas, (Format-Tamanho $res.Bytes)
                if ($res.SemAcesso -gt 0) { $lblResumo.Text += " — $($res.SemAcesso) pasta(s) sem permissão de leitura" }
                if ($res.Linhas.Count -gt $LimiteLista) { $lblResumo.Text += ' (lista mostra os primeiros {0:N0} de {1:N0} itens)' -f $LimiteLista, $res.Linhas.Count }
            }
            $lblStatus.Text = 'Análise concluída em {0:hh\:mm\:ss}.' -f ((Get-Date) - $estado.Inicio)
            if (Test-PastaOneDrive $estado.Pasta) { $lblStatus.Text += ' Atenção: pasta sincronizada com o OneDrive; a exclusão também ocorre na nuvem.' }
        }
        $lista.EndUpdate()
        & $definirOcupado $false
    }

    $concluirExclusao = {
        $res = $sync.Resultado
        & $definirOcupado $false
        if ($sync.Erro -or -not $res) {
            $lblStatus.Text = 'Erro durante a exclusão.'
            if (-not $estado.FecharAoTerminar) {
                [void][System.Windows.Forms.MessageBox]::Show($form, 'Erro durante a exclusão: ' + $sync.Erro, 'Erro', 'OK', 'Error')
            }
            return
        }
        $resumo = '{0:N0} arquivo(s) e {1:N0} pasta(s) excluídos em {2:hh\:mm\:ss}. A pasta selecionada foi mantida.' -f $res.Arquivos, $res.Pastas, ((Get-Date) - $estado.Inicio)
        if ($res.Cancelado) { $resumo = 'Exclusão cancelada. ' + $resumo }
        $lblStatus.Text = $resumo
        if ($estado.FecharAoTerminar) { return }
        if ($res.SemPermissao -gt 0 -and -not $ehAdministrador) {
            $texto = $resumo + "`n`n{0:N0} item(ns) não foram excluídos por falta de permissão (acesso negado).`n`nDeseja reabrir como administrador e tentar novamente?" -f $res.SemPermissao
            $resposta = [System.Windows.Forms.MessageBox]::Show($form, $texto, 'Permissão de administrador necessária', 'YesNo', 'Warning')
            if ($resposta -eq [System.Windows.Forms.DialogResult]::Yes) {
                if (Start-ComoAdministrador $estado.Pasta) { $form.Close(); return }
                [void][System.Windows.Forms.MessageBox]::Show($form, 'A elevação foi cancelada ou negada. Peça a um administrador do servidor para executar a ferramenta.', 'Sem elevação', 'OK', 'Information')
            }
            & $analisar
            return
        }
        if ($res.Falhas.Count -gt 0) {
            $exibir = @($res.Falhas | Select-Object -First 15)
            $texto = $resumo + "`n`nNão foi possível excluir $($res.Falhas.Count) item(ns) (em uso ou sem permissão):`n`n" + ($exibir -join "`n")
            if ($res.Falhas.Count -gt 15) { $texto += "`n... e mais $($res.Falhas.Count - 15)." }
            [void][System.Windows.Forms.MessageBox]::Show($form, $texto, 'Exclusão concluída com pendências', 'OK', 'Warning')
        } else {
            [void][System.Windows.Forms.MessageBox]::Show($form, $resumo, 'Exclusão concluída', 'OK', 'Information')
        }
        & $analisar
    }

    $timer.Add_Tick({
        if (-not $estado.Handle.IsCompleted) {
            $tempo = '{0:hh\:mm\:ss}' -f ((Get-Date) - $estado.Inicio)
            if ($sync.Cancelar) {
                $lblStatus.Text = 'Cancelando, aguarde...'
            } elseif ($estado.Tipo -eq 'Analise') {
                $lblStatus.Text = 'Analisando... {0:N0} item(ns) lido(s) — {1} — {2}' -f $sync.Processados, $tempo, $sync.Atual
            } else {
                $feitos = [int]$sync.Processados
                $progresso.Value = [Math]::Min($feitos, $progresso.Maximum)
                $percentual = [Math]::Floor(100 * $feitos / [Math]::Max(1, $estado.TotalArquivos))
                $lblStatus.Text = 'Excluindo... {0:N0} de {1:N0} arquivo(s) ({2}%) — {3} — {4}' -f $feitos, $estado.TotalArquivos, $percentual, $tempo, $sync.Atual
            }
            return
        }
        $timer.Stop()
        try { [void]$estado.Ps.EndInvoke($estado.Handle) } catch { if (-not $sync.Erro) { $sync.Erro = Get-MensagemErro $_ } }
        $estado.Ps.Dispose()
        $estado.Runspace.Dispose()
        $estado.Handle = $null
        if ($estado.Tipo -eq 'Analise') { & $mostrarAnalise } else { & $concluirExclusao }
        if ($estado.FecharAoTerminar -and -not $estado.Ocupado) { $form.Close() }
    })

    $analisar = {
        $motivo = Get-MotivoBloqueio $estado.Pasta
        $estado.TotalItens = 0
        if ($motivo) {
            $lista.Items.Clear()
            $lblResumo.Text = $motivo
            & $definirOcupado $false
            [void][System.Windows.Forms.MessageBox]::Show($form, $motivo, 'Pasta não permitida', 'OK', 'Warning')
            return
        }
        $lista.Items.Clear()
        $lblResumo.Text = 'Analisando o conteúdo da pasta...'
        $progresso.Style = 'Marquee'
        & $iniciarTarefa 'Analise'
    }

    $btnSelecionar.Add_Click({
        $dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialogo.Description = 'Selecione a pasta cujo CONTEÚDO será analisado (a pasta em si será mantida)'
        $dialogo.ShowNewFolderButton = $false
        if ($estado.Pasta) { $dialogo.SelectedPath = $estado.Pasta }
        if ($dialogo.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $estado.Pasta = $dialogo.SelectedPath
            $txtPasta.Text = $estado.Pasta
            & $analisar
        }
    })

    $btnAnalisar.Add_Click({ & $analisar })
    $btnCancelar.Add_Click({ $sync.Cancelar = $true; $lblStatus.Text = 'Cancelando, aguarde...' })
    $btnFechar.Add_Click({ $form.Close() })
    $btnAdmin.Add_Click({
        if (Start-ComoAdministrador $estado.Pasta) { $form.Close() }
        else { [void][System.Windows.Forms.MessageBox]::Show($form, 'A elevação foi cancelada ou negada.', 'Sem elevação', 'OK', 'Information') }
    })

    $btnExcluir.Add_Click({
        $raiz = $estado.Pasta
        $motivo = Get-MotivoBloqueio $raiz
        if ($motivo) { [void][System.Windows.Forms.MessageBox]::Show($form, $motivo, 'Pasta não permitida', 'OK', 'Warning'); return }
        $mensagem = "Excluir PERMANENTEMENTE {0:N0} arquivo(s) e {1:N0} subpasta(s) de dentro de:`n`n{2}`n`nA pasta será mantida; apenas o conteúdo será apagado.`nOs itens NÃO irão para a Lixeira e não poderão ser recuperados." -f $estado.TotalArquivos, $estado.TotalPastas, $raiz
        if (Test-PastaOneDrive $raiz) { $mensagem += "`n`nAtenção: pasta sincronizada com o OneDrive; a exclusão também ocorre na nuvem." }
        $mensagem += "`n`nDeseja continuar?"
        $resposta = [System.Windows.Forms.MessageBox]::Show($form, $mensagem, 'Confirmar exclusão permanente', 'YesNo', 'Warning', 'Button2')
        if ($resposta -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $progresso.Style = 'Blocks'
        $progresso.Maximum = [Math]::Max(1, $estado.TotalArquivos)
        $progresso.Value = 0
        $lblStatus.Text = 'Excluindo...'
        & $iniciarTarefa 'Exclusao'
    })

    $form.Add_FormClosing({
        if ($estado.Ocupado) {
            # Fechar durante uma tarefa: cancela com seguranca e fecha quando o segundo plano parar.
            $_.Cancel = $true
            $estado.FecharAoTerminar = $true
            $sync.Cancelar = $true
            $lblStatus.Text = 'Cancelando; a janela fechará em seguida...'
        }
    })

    $form.Add_Shown({
        if ($pastaInicial) {
            $estado.Pasta = $pastaInicial
            $txtPasta.Text = $pastaInicial
            & $analisar
        }
    })

    [void]$form.ShowDialog()
    $timer.Dispose()
    $form.Dispose()
}

if ($SemConfirmar) {
    # Modo sem tela, para automacoes.
    if (-not $Pasta) { Write-Host 'Erro: informe -Pasta junto com -SemConfirmar.' -ForegroundColor Red; exit 1 }
    $motivo = Get-MotivoBloqueio $Pasta
    if ($motivo) { Write-Host "Erro: $motivo" -ForegroundColor Red; exit 1 }
    $raiz = [IO.Path]::GetFullPath($Pasta).TrimEnd('\')
    $syncLocal = @{ Cancelar = $false; Processados = 0; Atual = '' }
    $resultado = Remove-ConteudoPasta (ConvertTo-CaminhoLongo $raiz) $syncLocal
    Write-Host ('{0} arquivo(s) e {1} pasta(s) excluídos de {2}. A pasta foi mantida.' -f $resultado.Arquivos, $resultado.Pastas, $raiz) -ForegroundColor Green
    if ($resultado.Falhas.Count -gt 0) {
        Write-Host "Não foi possível excluir $($resultado.Falhas.Count) item(ns):" -ForegroundColor Yellow
        $resultado.Falhas | ForEach-Object { Write-Host "  $_" }
        exit 2
    }
    exit 0
}

$pastaInicial = $null
if ($Pasta -and (Test-Path -LiteralPath $Pasta -PathType Container)) { $pastaInicial = [IO.Path]::GetFullPath($Pasta).TrimEnd('\') }
Show-Interface $pastaInicial
