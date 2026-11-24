function New-IncluirUsuariosWebPanel {
    <#
      Painel compartilhado pelas rotinas que rodam um script de console
      (03, 04, 06, 07, 08, 09 e 10).

      O script Python/PowerShell continua o mesmo: este painel le a saida,
      reconhece as perguntas e mostra o controle adequado (lista, opcoes,
      Sim/Nao...). A resposta enviada ao script e o mesmo texto que seria
      digitado no console. Pergunta nao reconhecida cai numa caixa de texto.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Parent,
        [Parameter(Mandatory)][string]$CentralRoot,
        [Parameter(Mandatory)][scriptblock]$OnBack,
        [string]$PanelTitle = '03 - Incluir usuários em projetos PW Web',
        [string]$PythonRelativePath = '..\ExecucaoManualUsuario\PWDM_Gerenciamento_Participantes_V2\incluir_usuarios_projetos_roles_pwdm_v2.py',
        [string]$LogPrefix = 'PW_Web_IncluirUsuarios',
        [string]$LogDirectory = '03-IncluirUsuariosWeb',
        [ValidateSet('Python','PowerShell')][string]$Runtime = 'Python',
        [string]$HelpText = '',
        [string]$StartConfirmationText = '',
        [string]$RuntimeArguments = '',
        [string]$SecondaryButtonText = '',
        [string]$SecondaryRuntimeArguments = '',
        [string]$SecondaryConfirmationText = '',
        [string[]]$Etapas = @(
            'Escolha a concessão e os projetos no ProjectWise.',
            'Informe o usuário ou a planilha de usuários.',
            'Escolha a role do PW Web.',
            'Confira o resumo e confirme.',
            'Acompanhe a inclusão em cada projeto.'
        )
    )

    if (-not ('CentralScripts.StreamPump' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Collections.Concurrent;
using System.IO;
using System.Threading.Tasks;
namespace CentralScripts {
    public static class StreamPump {
        public static Task Start(StreamReader reader, ConcurrentQueue<string> queue, string prefix) {
            return Task.Run(() => {
                int value;
                bool first = true;
                while ((value = reader.Read()) >= 0) {
                    string text = new string((char)value, 1);
                    queue.Enqueue((first ? prefix : "") + text);
                    first = false;
                }
            });
        }
        public static Task StartFilteredError(StreamReader reader, ConcurrentQueue<string> queue) {
            return Task.Run(() => {
                string text = reader.ReadToEnd();
                if (string.IsNullOrWhiteSpace(text)) return;
                bool isPowerShellProgress = text.StartsWith("#< CLIXML")
                    && text.Contains("S=\"progress\"")
                    && !text.Contains("S=\"Error\"");
                if (!isPowerShellProgress) queue.Enqueue("[ERRO] " + text + "\n");
            });
        }
    }
}
'@
    }

    # ---------------------------------------------------------
    # Tema (Tema.ps1 da Central). Tudo e copiado para variaveis locais
    # porque os eventos usam closures, onde $script: nao alcanca o tema.
    # ---------------------------------------------------------
    function Get-PainelCor([int]$R, [int]$G, [int]$B) { [System.Drawing.Color]::FromArgb($R, $G, $B) }
    $T = $script:Tema
    $cor = @{
        Verde        = if ($T) { $T.VerdeEscuro }  else { Get-PainelCor 5 107 40 }
        VerdeHover   = if ($T) { $T.VerdeHover }   else { Get-PainelCor 4 88 33 }
        VerdeMedio   = if ($T) { $T.VerdeMedio }   else { Get-PainelCor 29 160 19 }
        VerdeSuave   = if ($T) { $T.VerdeSuave }   else { Get-PainelCor 228 243 226 }
        Fundo        = if ($T) { $T.Fundo }        else { Get-PainelCor 243 246 244 }
        Camada       = if ($T) { $T.Camada }       else { [System.Drawing.Color]::White }
        CamadaAlt    = if ($T) { $T.CamadaAlt }    else { Get-PainelCor 248 250 248 }
        Contorno     = if ($T) { $T.Contorno }     else { Get-PainelCor 221 229 223 }
        Texto        = if ($T) { $T.Texto }        else { Get-PainelCor 23 37 28 }
        TextoSuave   = if ($T) { $T.TextoSuave }   else { Get-PainelCor 92 107 98 }
        Critico      = if ($T) { $T.Critico }      else { Get-PainelCor 196 43 28 }
        CriticoSuave = if ($T) { $T.CriticoSuave } else { Get-PainelCor 253 231 229 }
        Atencao      = if ($T) { $T.Atencao }      else { Get-PainelCor 168 70 12 }
        AtencaoSuave = if ($T) { $T.AtencaoSuave } else { Get-PainelCor 254 236 222 }
        Ok           = Get-PainelCor 30 127 58
        Console      = Get-PainelCor 24 28 34
    }
    $temFonte = [bool](Get-Command Get-TemaFonte -ErrorAction SilentlyContinue)
    function Get-PainelFonte([string]$Papel, [single]$Tamanho, [string]$Reserva, [System.Drawing.FontStyle]$Estilo = 'Regular') {
        if ($temFonte) { return Get-TemaFonte -Papel $Papel -Tamanho $Tamanho }
        return New-Object System.Drawing.Font($Reserva, $Tamanho, $Estilo)
    }
    $fonte = @{
        Titulo     = Get-PainelFonte Titulo 15 'Segoe UI Semibold'
        Secao      = Get-PainelFonte Destaque 11.5 'Segoe UI Semibold'
        Texto      = Get-PainelFonte Texto 9.5 'Segoe UI'
        TextoForte = Get-PainelFonte TextoForte 9.5 'Segoe UI Semibold'
        Legenda    = Get-PainelFonte Legenda 8.5 'Segoe UI'
        Numero     = Get-PainelFonte Destaque 10 'Segoe UI Semibold'
        Mono       = New-Object System.Drawing.Font('Consolas', 9)
    }

    # ---------------------------------------------------------
    # Estado
    # ---------------------------------------------------------
    $state = [PSCustomObject]@{
        Process = $null
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
        LogPath = $null
        Finished = $false
        OutputTask = $null
        ErrorTask = $null
        PromptBuffer = ''
        DialogOpen = $false
        LastOutputUtc = [DateTime]::UtcNow
        NextRuntimeArguments = $RuntimeArguments
        NextConfirmationText = $StartConfirmationText
    }
    $logsDir = Join-Path (Join-Path $CentralRoot 'Logs') $LogDirectory
    if (-not (Test-Path -LiteralPath $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }

    # $ctx e compartilhado por todas as closures (referencia unica).
    $ctx = @{
        Cor = $cor; Fonte = $fonte; State = $state
        Etapas = $Etapas; HelpText = $HelpText; PedeConfirmacao = -not [string]::IsNullOrWhiteSpace($StartConfirmationText)
        LogsDir = $logsDir; LogPrefix = $LogPrefix
        LinhaPendente = ''; Blocos = New-Object System.Collections.Generic.List[object]
        UltimaEraLista = $false; UltimoTitulo = ''; Recentes = New-Object System.Collections.Generic.List[string]
        UltimaResposta = $null; Etapa = -1; EtapaTotal = 0; EtapaTexto = ''
        Ok = 0; Erros = 0; Avisos = 0; UltimosErros = New-Object System.Collections.Generic.List[string]
        Cancelado = $false; FeedSujo = $false; FeedVazio = $true
        Marcados = New-Object 'System.Collections.Generic.HashSet[string]'
        ListaItens = @(); ListaVisivel = @(); Reconstruindo = $false
        Ruido = '^(Security Protocol now set|PWPS_DAB Version|This is a \d+-bit process|\.Net Framework version|AVISO: Os nomes de alguns comandos importados|WARNING: The names of some imported commands|menos detect|par[aâ]metro Verbose|include unapproved verbs|Verbose parameter|Uma janela/login do ProjectWise pode ser aberta)'
    }

    # ---------------------------------------------------------
    # Estrutura da tela
    # ---------------------------------------------------------
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.BackColor = $cor.Fundo
    $panel.Tag = $state
    $Parent.Controls.Add($panel)
    $panel.BringToFront()

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.Margin = New-Object System.Windows.Forms.Padding(0)
    $root.ColumnCount = 1
    $root.RowCount = 5
    [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 64)))    # cabecalho
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 0)))     # etapas
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))    # andamento + interacao
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 0)))     # detalhes tecnicos
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 60)))    # acoes
    $panel.Controls.Add($root)
    $ctx.Root = $root

    # Cabecalho
    $header = New-Object System.Windows.Forms.TableLayoutPanel
    $header.Dock = 'Fill'
    $header.Margin = New-Object System.Windows.Forms.Padding(0)
    $header.Padding = New-Object System.Windows.Forms.Padding(12, 0, 20, 0)
    $header.BackColor = $cor.Camada
    $header.ColumnCount = 3
    [void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$header.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $header.Add_Paint({
        param($c, $e)
        $p = New-Object System.Drawing.Pen($ctx.Cor.Contorno)
        $e.Graphics.DrawLine($p, 0, ($c.Height - 1), $c.Width, ($c.Height - 1))
        $p.Dispose()
    }.GetNewClosure())
    $root.Controls.Add($header, 0, 0)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = '←  Voltar'
    $backButton.FlatStyle = 'Flat'
    $backButton.FlatAppearance.BorderSize = 0
    $backButton.FlatAppearance.MouseOverBackColor = $cor.VerdeSuave
    $backButton.BackColor = $cor.Camada
    $backButton.ForeColor = $cor.Verde
    $backButton.Font = $fonte.TextoForte
    $backButton.Size = New-Object System.Drawing.Size(92, 36)
    $backButton.Anchor = 'Left'
    $backButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $header.Controls.Add($backButton, 0, 0)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = $PanelTitle
    $title.Font = $fonte.Titulo
    $title.ForeColor = $cor.Verde
    $title.AutoSize = $true
    $title.Anchor = 'Left'
    $title.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    $header.Controls.Add($title, 1, 0)

    $chip = New-Object System.Windows.Forms.Label
    $chip.AutoSize = $true
    $chip.Anchor = 'Right'
    $chip.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $chip.Font = $fonte.TextoForte
    $header.Controls.Add($chip, 2, 0)
    $ctx.Chip = $chip

    # Barra de etapas (aparece quando o script informa [n/N])
    $stepper = New-Object System.Windows.Forms.Panel
    $stepper.Dock = 'Fill'
    $stepper.Margin = New-Object System.Windows.Forms.Padding(0)
    $stepper.BackColor = $cor.Fundo
    $prop = $stepper.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic,Instance')
    if ($prop) { $prop.SetValue($stepper, $true, $null) }
    $stepper.Add_Paint({
        param($c, $e)
        $g = $e.Graphics
        $g.SmoothingMode = 'AntiAlias'
        $x = 20; $w = $c.Width - 40
        $total = [Math]::Max(1, $ctx.EtapaTotal)
        $atual = [Math]::Max(0, [Math]::Min($ctx.Etapa, $total))
        $texto = "Etapa $atual de $total"
        if ($ctx.EtapaTexto) { $texto += "  ·  $($ctx.EtapaTexto)" }
        [System.Windows.Forms.TextRenderer]::DrawText($g, $texto, $ctx.Fonte.TextoForte, (New-Object System.Drawing.Point($x, 6)), $ctx.Cor.Texto)
        $y = 30
        $trilho = New-Object System.Drawing.SolidBrush($ctx.Cor.Contorno)
        $g.FillRectangle($trilho, $x, $y, $w, 6); $trilho.Dispose()
        $preenchido = [int]($w * $atual / $total)
        if ($preenchido -gt 0) {
            $barra = New-Object System.Drawing.SolidBrush($ctx.Cor.VerdeMedio)
            $g.FillRectangle($barra, $x, $y, $preenchido, 6); $barra.Dispose()
        }
    }.GetNewClosure())
    $root.Controls.Add($stepper, 0, 1)
    $ctx.Stepper = $stepper

    # Corpo: andamento (esquerda) + interacao (direita)
    $body = New-Object System.Windows.Forms.TableLayoutPanel
    $body.Dock = 'Fill'
    $body.Margin = New-Object System.Windows.Forms.Padding(0)
    $body.Padding = New-Object System.Windows.Forms.Padding(20, 14, 20, 6)
    $body.ColumnCount = 2
    [void]$body.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$body.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 440)))
    [void]$body.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $root.Controls.Add($body, 0, 2)

    $feedHost = New-Object System.Windows.Forms.Panel
    $feedHost.Dock = 'Fill'
    $feedHost.Margin = New-Object System.Windows.Forms.Padding(0, 0, 14, 0)
    $feedHost.Padding = New-Object System.Windows.Forms.Padding(16, 12, 6, 8)
    $feedHost.BackColor = $cor.Camada
    $feedHost.Add_Paint({
        param($c, $e)
        $p = New-Object System.Drawing.Pen($ctx.Cor.Contorno)
        $e.Graphics.DrawRectangle($p, 0, 0, ($c.Width - 1), ($c.Height - 1))
        $p.Dispose()
    }.GetNewClosure())
    $body.Controls.Add($feedHost, 0, 0)

    $feed = New-Object System.Windows.Forms.RichTextBox
    $feed.Dock = 'Fill'
    $feed.ReadOnly = $true
    $feed.BorderStyle = 'None'
    $feed.BackColor = $cor.Camada
    $feed.ForeColor = $cor.Texto
    $feed.Font = $fonte.Texto
    $feed.DetectUrls = $false
    $feedHost.Controls.Add($feed)
    $ctx.FeedBox = $feed

    $side = New-Object System.Windows.Forms.Panel
    $side.Dock = 'Fill'
    $side.Margin = New-Object System.Windows.Forms.Padding(0)
    $side.Padding = New-Object System.Windows.Forms.Padding(18, 16, 18, 16)
    $side.BackColor = $cor.Camada
    $side.Add_Paint({
        param($c, $e)
        $corBorda = if ($ctx.State.DialogOpen) { $ctx.Cor.VerdeMedio } else { $ctx.Cor.Contorno }
        $p = New-Object System.Drawing.Pen($corBorda, $(if ($ctx.State.DialogOpen) { 2 } else { 1 }))
        $e.Graphics.DrawRectangle($p, 0, 0, ($c.Width - 1), ($c.Height - 1))
        $p.Dispose()
    }.GetNewClosure())
    $body.Controls.Add($side, 1, 0)
    $ctx.Lado = $side

    # Detalhes tecnicos (console original, recolhido)
    $output = New-Object System.Windows.Forms.RichTextBox
    $output.Dock = 'Fill'
    $output.ReadOnly = $true
    $output.BackColor = $cor.Console
    $output.ForeColor = [System.Drawing.Color]::Gainsboro
    $output.Font = $fonte.Mono
    $output.BorderStyle = 'None'
    $output.Margin = New-Object System.Windows.Forms.Padding(20, 0, 20, 6)
    $root.Controls.Add($output, 0, 3)

    # Acoes
    $footer = New-Object System.Windows.Forms.TableLayoutPanel
    $footer.Dock = 'Fill'
    $footer.Margin = New-Object System.Windows.Forms.Padding(0)
    $footer.Padding = New-Object System.Windows.Forms.Padding(20, 0, 20, 0)
    $footer.BackColor = $cor.Camada
    $footer.ColumnCount = 2
    [void]$footer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$footer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$footer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $footer.Add_Paint({
        param($c, $e)
        $p = New-Object System.Drawing.Pen($ctx.Cor.Contorno)
        $e.Graphics.DrawLine($p, 0, 0, $c.Width, 0)
        $p.Dispose()
    }.GetNewClosure())
    $root.Controls.Add($footer, 0, 4)

    $actions = New-Object System.Windows.Forms.FlowLayoutPanel
    $actions.AutoSize = $true
    $actions.WrapContents = $false
    $actions.Anchor = 'Left'
    $actions.Margin = New-Object System.Windows.Forms.Padding(0)
    $footer.Controls.Add($actions, 0, 0)
    $extras = New-Object System.Windows.Forms.FlowLayoutPanel
    $extras.AutoSize = $true
    $extras.WrapContents = $false
    $extras.Anchor = 'Right'
    $extras.Margin = New-Object System.Windows.Forms.Padding(0)
    $footer.Controls.Add($extras, 1, 0)

    # ---------------------------------------------------------
    # Auxiliares de interface
    # ---------------------------------------------------------
    $ctx.Estilizar = {
        param($b, [string]$Tipo)
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        $b.FlatStyle = 'Flat'
        $b.UseVisualStyleBackColor = $false
        $b.Cursor = [System.Windows.Forms.Cursors]::Hand
        $b.Font = $ctx.Fonte.Texto
        switch ($Tipo) {
            'Principal' {
                $b.BackColor = $ctx.Cor.Verde; $b.ForeColor = [System.Drawing.Color]::White
                $b.FlatAppearance.BorderColor = $ctx.Cor.Verde
                $b.FlatAppearance.MouseOverBackColor = $ctx.Cor.VerdeHover
                $b.Font = $ctx.Fonte.TextoForte
            }
            'Perigo' {
                $b.BackColor = $ctx.Cor.Critico; $b.ForeColor = [System.Drawing.Color]::White
                $b.FlatAppearance.BorderColor = $ctx.Cor.Critico
                $b.Font = $ctx.Fonte.TextoForte
            }
            'Opcao' {
                $b.BackColor = $ctx.Cor.Camada; $b.ForeColor = $ctx.Cor.Texto
                $b.FlatAppearance.BorderColor = $ctx.Cor.Contorno
                $b.FlatAppearance.MouseOverBackColor = $ctx.Cor.VerdeSuave
                $b.TextAlign = 'MiddleLeft'
                $b.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
            }
            default {
                $b.BackColor = $ctx.Cor.Camada; $b.ForeColor = $ctx.Cor.Texto
                $b.FlatAppearance.BorderColor = $ctx.Cor.Contorno
                $b.FlatAppearance.MouseOverBackColor = $ctx.Cor.CamadaAlt
            }
        }
        $b.Add_EnabledChanged({
            param($s, $e)
            if (-not $s.Tag) { $s.Tag = @{ Fundo = $s.BackColor; Texto = $s.ForeColor } }
            if ($s.Enabled) { $s.BackColor = $s.Tag.Fundo; $s.ForeColor = $s.Tag.Texto }
            else { $s.BackColor = $ctx.Cor.Contorno; $s.ForeColor = $ctx.Cor.TextoSuave }
        }.GetNewClosure())
    }.GetNewClosure()

    $ctx.NovoBotao = {
        param([string]$Texto, [string]$Tipo = 'Secundario', [int]$Largura = 0)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $Texto
        $b.Height = 38
        if ($Largura -gt 0) { $b.Width = $Largura } else { $b.Dock = 'Top' }
        $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
        & $ctx.Estilizar $b $Tipo
        return $b
    }.GetNewClosure()

    $ctx.DefinirChip = {
        param([string]$Texto, [string]$Tom)
        $ctx.Chip.Text = $Texto
        switch ($Tom) {
            'Verde'   { $ctx.Chip.BackColor = $ctx.Cor.VerdeSuave;   $ctx.Chip.ForeColor = $ctx.Cor.Verde }
            'Atencao' { $ctx.Chip.BackColor = $ctx.Cor.AtencaoSuave; $ctx.Chip.ForeColor = $ctx.Cor.Atencao }
            'Critico' { $ctx.Chip.BackColor = $ctx.Cor.CriticoSuave; $ctx.Chip.ForeColor = $ctx.Cor.Critico }
            default   { $ctx.Chip.BackColor = $ctx.Cor.Fundo;        $ctx.Chip.ForeColor = $ctx.Cor.TextoSuave }
        }
    }.GetNewClosure()

    # Painel lateral: cada estado monta um bloco vertical novo.
    $ctx.NovoBloco = {
        $ctx.Lado.SuspendLayout()
        foreach ($c in @($ctx.Lado.Controls)) { $ctx.Lado.Controls.Remove($c); $c.Dispose() }
        $tl = New-Object System.Windows.Forms.TableLayoutPanel
        $tl.Dock = 'Fill'
        $tl.ColumnCount = 1
        $tl.RowCount = 0
        $tl.BackColor = $ctx.Cor.Camada
        [void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
        $ctx.Lado.Controls.Add($tl)
        $ctx.Bloco = $tl
        $ctx.Lado.ResumeLayout()
        $ctx.Lado.Invalidate()
    }.GetNewClosure()

    $ctx.Linha = {
        param($Controle, [switch]$Preencher)
        $tl = $ctx.Bloco
        $estilo = if ($Preencher) { New-Object System.Windows.Forms.RowStyle('Percent', 100) } else { New-Object System.Windows.Forms.RowStyle('AutoSize') }
        [void]$tl.RowStyles.Add($estilo)
        $tl.RowCount = $tl.RowStyles.Count
        # Dock Top (e nao Fill): a ultima linha recebe a sobra de altura e o botao esticaria.
        $tl.Controls.Add($Controle, 0, ($tl.RowCount - 1))
    }.GetNewClosure()

    $ctx.Rotulo = {
        param([string]$Texto, $Fonte, $Cor, [int]$Abaixo = 6)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Texto
        $l.Font = if ($Fonte) { $Fonte } else { $ctx.Fonte.Texto }
        $l.ForeColor = if ($Cor) { $Cor } else { $ctx.Cor.Texto }
        $l.AutoSize = $true
        $l.MaximumSize = New-Object System.Drawing.Size(396, 0)
        $l.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, $Abaixo)
        & $ctx.Linha $l
        return $l
    }.GetNewClosure()

    # ---------------------------------------------------------
    # Andamento legivel
    # ---------------------------------------------------------
    $ctx.Feed = {
        param([string]$Texto, [string]$Tipo = 'Normal')
        $rtb = $ctx.FeedBox
        if ($ctx.FeedVazio) { $rtb.Clear(); $ctx.FeedVazio = $false }
        $prefixo = ''
        $f = $ctx.Fonte.Texto
        $c = $ctx.Cor.Texto
        switch ($Tipo) {
            'Etapa'    { $f = $ctx.Fonte.TextoForte; $c = $ctx.Cor.Verde; if ($rtb.TextLength -gt 0) { $prefixo = "`n" } }
            'Ok'       { $c = $ctx.Cor.Ok; $prefixo = '✓  ' }
            'Erro'     { $c = $ctx.Cor.Critico; $f = $ctx.Fonte.TextoForte; $prefixo = '✕  ' }
            'Aviso'    { $c = $ctx.Cor.Atencao; $prefixo = '⚠  ' }
            'Info'     { $c = $ctx.Cor.TextoSuave }
            'Resposta' { $c = $ctx.Cor.Verde; $f = $ctx.Fonte.TextoForte; $prefixo = '→  ' }
            'Sistema'  { $c = $ctx.Cor.TextoSuave; $f = $ctx.Fonte.Legenda }
        }
        $rtb.SelectionStart = $rtb.TextLength
        $rtb.SelectionLength = 0
        $rtb.SelectionColor = $c
        $rtb.SelectionFont = $f
        $rtb.AppendText($prefixo + $Texto + "`n")
        $ctx.FeedSujo = $true
    }.GetNewClosure()

    $ctx.FecharBlocoLista = {
        # Listas curtas aparecem inteiras no andamento; longas viram um resumo
        # (a lista completa fica no console e no painel de selecao).
        if ($ctx.UltimaEraLista -and $ctx.Blocos.Count -gt 0) {
            $bloco = $ctx.Blocos[$ctx.Blocos.Count - 1]
            # O titulo da lista ja saiu no andamento como linha normal; aqui so os itens.
            if ($bloco.Itens.Count -le 12) {
                foreach ($item in $bloco.Itens) { & $ctx.Feed ("    {0}. {1}" -f $item.Numero, $item.Nome) 'Info' }
            }
            else {
                & $ctx.Feed ("    ({0} itens; a lista completa está em Detalhes técnicos)" -f $bloco.Itens.Count) 'Info'
            }
        }
        $ctx.UltimaEraLista = $false
    }.GetNewClosure()

    $ctx.ProcessarLinha = {
        param([string]$Bruta)
        $t = $Bruta.Trim()
        if ($t -eq '') { & $ctx.FecharBlocoLista; return }

        # Itens de lista: "01. Nome", "001. Nome | ID PW: ...", "[1] Nome"
        $m = [regex]::Match($t, '^(?:\[(\d{1,4})\]|(\d{1,4})\.)\s+(.+)$')
        if ($m.Success) {
            $numero = if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[2].Value }
            $nome = $m.Groups[3].Value.Trim()
            $extra = ''
            $corte = $nome.IndexOf(' | ID PW:')
            if ($corte -ge 0) { $extra = $nome.Substring($corte + 3).Trim(); $nome = $nome.Substring(0, $corte).Trim() }
            if (-not $ctx.UltimaEraLista) {
                $ctx.Blocos.Add([PSCustomObject]@{ Titulo = $ctx.UltimoTitulo; Itens = New-Object System.Collections.Generic.List[object] })
                while ($ctx.Blocos.Count -gt 8) { $ctx.Blocos.RemoveAt(0) }
            }
            $ctx.Blocos[$ctx.Blocos.Count - 1].Itens.Add([PSCustomObject]@{ Numero = $numero; Nome = $nome; Extra = $extra })
            $ctx.UltimaEraLista = $true
            return
        }
        & $ctx.FecharBlocoLista

        if ($t -match $ctx.Ruido) { return }
        if ($null -ne $ctx.UltimaResposta -and $t -eq $ctx.UltimaResposta) { $ctx.UltimaResposta = $null; return }

        # A ultima linha antes de uma lista vira o titulo dela ("Modo de usuarios", "Concessoes disponiveis:").
        $ctx.UltimoTitulo = $t.TrimEnd(':').Trim()
        $ctx.Recentes.Add($t)
        while ($ctx.Recentes.Count -gt 40) { $ctx.Recentes.RemoveAt(0) }

        $etapa = [regex]::Match($t, '^\[(\d+)/(\d+)\]\s*(.*)$')
        if ($etapa.Success) {
            $ctx.Etapa = [int]$etapa.Groups[1].Value
            $ctx.EtapaTotal = [int]$etapa.Groups[2].Value
            $ctx.EtapaTexto = $etapa.Groups[3].Value.Trim().TrimEnd('.')
            if ($ctx.Root.RowStyles[1].Height -eq 0) { $ctx.Root.RowStyles[1].Height = 46 }
            $ctx.Stepper.Invalidate()
            & $ctx.Feed ("Etapa {0} de {1}  ·  {2}" -f $ctx.Etapa, $ctx.EtapaTotal, $ctx.EtapaTexto) 'Etapa'
            return
        }

        $tipo = 'Normal'
        if ($t -match '^\[ERRO\]|\|\s*ERRO\s*\||^ERRO\b|^Traceback|Exception:') {
            $tipo = 'Erro'; $ctx.Erros++
            $t = $t -replace '^\[ERRO\]\s*', ''
            $ctx.UltimosErros.Add($t)
            while ($ctx.UltimosErros.Count -gt 4) { $ctx.UltimosErros.RemoveAt(0) }
        }
        elseif ($t -match '\[OK\]') { $tipo = 'Ok'; $ctx.Ok++; $t = $t -replace '\[OK\]\s*', '' }
        elseif ($t -match '^\[AVISO\]|\|\s*AVISO\s*\||^AVISO\b') { $tipo = 'Aviso'; $ctx.Avisos++; $t = $t -replace '^\[AVISO\]\s*', '' }
        elseif ($t -match '^\[INFO\]') { $tipo = 'Info'; $t = $t -replace '^\[INFO\]\s*', '' }
        & $ctx.Feed $t $tipo
    }.GetNewClosure()

    # ---------------------------------------------------------
    # Estados do painel lateral
    # ---------------------------------------------------------
    $ctx.MostrarRepouso = {
        & $ctx.NovoBloco
        [void](& $ctx.Rotulo 'Como funciona' $ctx.Fonte.Secao $ctx.Cor.Texto 12)
        $i = 1
        foreach ($etapaTexto in @($ctx.Etapas)) {
            $linha = New-Object System.Windows.Forms.TableLayoutPanel
            $linha.AutoSize = $true
            $linha.ColumnCount = 2
            $linha.Dock = 'Top'
            $linha.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
            [void]$linha.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 30)))
            [void]$linha.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
            $n = New-Object System.Windows.Forms.Label
            $n.Text = [string]$i; $n.Font = $ctx.Fonte.Numero; $n.ForeColor = $ctx.Cor.Verde
            $n.BackColor = $ctx.Cor.VerdeSuave; $n.Size = New-Object System.Drawing.Size(22, 22); $n.TextAlign = 'MiddleCenter'
            $n.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
            $txt = New-Object System.Windows.Forms.Label
            $txt.Text = $etapaTexto; $txt.Font = $ctx.Fonte.Texto; $txt.ForeColor = $ctx.Cor.Texto
            $txt.AutoSize = $true; $txt.MaximumSize = New-Object System.Drawing.Size(360, 0)
            $txt.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
            $linha.Controls.Add($n, 0, 0); $linha.Controls.Add($txt, 1, 0)
            & $ctx.Linha $linha
            $i++
        }
        if ($ctx.HelpText) {
            [void](& $ctx.Rotulo ($ctx.HelpText -replace "`r?`n", ' ') $ctx.Fonte.Legenda $ctx.Cor.TextoSuave 8)
        }
        if ($ctx.PedeConfirmacao) {
            [void](& $ctx.Rotulo '⚠  Esta rotina pede confirmação antes de alterar dados.' $ctx.Fonte.Legenda $ctx.Cor.Atencao 8)
        }
        $ultimo = Get-ChildItem -LiteralPath $ctx.LogsDir -Filter "$($ctx.LogPrefix)_*.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $textoUltimo = if ($ultimo) { "Última execução: $($ultimo.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))" } else { 'Nenhuma execução registrada.' }
        [void](& $ctx.Rotulo $textoUltimo $ctx.Fonte.Legenda $ctx.Cor.TextoSuave 0)
    }.GetNewClosure()

    $ctx.MostrarExecucao = {
        & $ctx.NovoBloco
        [void](& $ctx.Rotulo 'Em execução' $ctx.Fonte.Secao $ctx.Cor.Verde 10)
        [void](& $ctx.Rotulo 'Nenhuma ação sua é necessária agora. Quando a rotina precisar de uma resposta, a pergunta aparece aqui.' $ctx.Fonte.Texto $ctx.Cor.TextoSuave 10)
        [void](& $ctx.Rotulo 'O andamento detalhado fica à esquerda; o console completo está em "Detalhes técnicos".' $ctx.Fonte.Legenda $ctx.Cor.TextoSuave 0)
    }.GetNewClosure()

    $ctx.MostrarFim = {
        param([int]$Codigo)
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        & $ctx.NovoBloco
        if ($ctx.Cancelado) { $titulo = 'Execução cancelada'; $tom = 'Atencao' }
        elseif ($Codigo -ne 0) { $titulo = "Finalizada com erro (código $Codigo)"; $tom = 'Critico' }
        elseif ($ctx.Erros -gt 0) { $titulo = 'Concluída com erros'; $tom = 'Critico' }
        else { $titulo = 'Concluída'; $tom = 'Verde' }
        $corTitulo = switch ($tom) { 'Verde' { $ctx.Cor.Verde } 'Critico' { $ctx.Cor.Critico } default { $ctx.Cor.Atencao } }
        & $ctx.DefinirChip $titulo $tom
        [void](& $ctx.Rotulo $titulo $ctx.Fonte.Secao $corTitulo 12)
        [void](& $ctx.Rotulo ("✓  {0} operação(ões) com sucesso" -f $ctx.Ok) $ctx.Fonte.Texto $ctx.Cor.Ok 4)
        [void](& $ctx.Rotulo ("✕  {0} erro(s)" -f $ctx.Erros) $ctx.Fonte.Texto $(if ($ctx.Erros) { $ctx.Cor.Critico } else { $ctx.Cor.TextoSuave }) 4)
        [void](& $ctx.Rotulo ("⚠  {0} aviso(s)" -f $ctx.Avisos) $ctx.Fonte.Texto $(if ($ctx.Avisos) { $ctx.Cor.Atencao } else { $ctx.Cor.TextoSuave }) 12)
        if ($ctx.UltimosErros.Count -gt 0) {
            [void](& $ctx.Rotulo 'Últimos erros' $ctx.Fonte.TextoForte $ctx.Cor.Texto 4)
            foreach ($erro in $ctx.UltimosErros) {
                $curto = if ($erro.Length -gt 160) { $erro.Substring(0, 157) + '...' } else { $erro }
                [void](& $ctx.Rotulo $curto $ctx.Fonte.Legenda $ctx.Cor.Critico 4)
            }
        }
        $abrir = & $ctx.NovoBotao 'Abrir log desta execução' 'Secundario'
        $abrir.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 8)
        $abrir.Add_Click({
            if ($ctx.State.LogPath -and (Test-Path -LiteralPath $ctx.State.LogPath)) {
                Start-Process explorer.exe -ArgumentList ('/select,"{0}"' -f $ctx.State.LogPath)
            }
        }.GetNewClosure())
        & $ctx.Linha $abrir
        $denovo = & $ctx.NovoBotao 'Executar novamente' 'Principal'
        $denovo.Add_Click({ $ctx.StartButton.PerformClick() }.GetNewClosure())
        & $ctx.Linha $denovo
    }.GetNewClosure()

    # ---------------------------------------------------------
    # Perguntas
    # ---------------------------------------------------------
    $ctx.Responder = {
        param([string]$Valor, [string]$Exibicao = '')
        $st = $ctx.State
        if (-not $st.Process -or $st.Process.HasExited -or -not $st.DialogOpen) { return }
        try {
            if ($st.Input) { $st.Input.WriteLine($Valor) }
            else { $st.Process.StandardInput.WriteLine($Valor); $st.Process.StandardInput.Flush() }
        }
        catch {
            & $ctx.Feed "Não foi possível enviar a resposta: $($_.Exception.Message)" 'Erro'
            return
        }
        $mostrar = if ($Exibicao) { $Exibicao } elseif ([string]::IsNullOrEmpty($Valor)) { 'Continuar' } else { $Valor }
        & $ctx.AppendOutput ("> " + $(if ([string]::IsNullOrEmpty($Valor)) { '[CONTINUAR]' } else { $Valor }))
        & $ctx.Feed $mostrar 'Resposta'
        $ctx.UltimaResposta = $Valor.Trim()
        # O texto da pergunta ("Escolha [1/2]: ") fica sem quebra de linha; descarta
        # para nao aparecer no andamento nem grudar na proxima linha.
        $ctx.LinhaPendente = ''
        $st.PromptBuffer = ''
        $st.DialogOpen = $false
        & $ctx.DefinirChip 'Em execução' 'Verde'
        & $ctx.MostrarExecucao
    }.GetNewClosure()

    $ctx.UltimoBloco = {
        param([string]$PadraoTitulo)
        for ($i = $ctx.Blocos.Count - 1; $i -ge 0; $i--) {
            $b = $ctx.Blocos[$i]
            if (-not $PadraoTitulo -or $b.Titulo -match $PadraoTitulo) { return $b }
        }
        return $null
    }.GetNewClosure()

    $ctx.Resumo = {
        # Linhas "- Chave: valor" logo antes da pergunta (ex.: confirmacao final).
        $linhas = New-Object System.Collections.Generic.List[string]
        for ($i = $ctx.Recentes.Count - 1; $i -ge 0; $i--) {
            $l = $ctx.Recentes[$i]
            if ($l -match '^- ') { if ($l -notmatch 'Responda S|digite exatamente') { $linhas.Insert(0, $l.Substring(2)) } continue }
            if ($linhas.Count -eq 0 -and $l -notmatch '\[S/N\]') { $linhas.Add($l) }
            break
        }
        return ,$linhas.ToArray()
    }.GetNewClosure()

    $ctx.Cabecalho = {
        param([string]$Titulo, [string]$Texto)
        & $ctx.NovoBloco
        [void](& $ctx.Rotulo 'Sua resposta é necessária' $ctx.Fonte.Legenda $ctx.Cor.Atencao 4)
        [void](& $ctx.Rotulo $Titulo $ctx.Fonte.Secao $ctx.Cor.Texto 6)
        if ($Texto) { [void](& $ctx.Rotulo $Texto $ctx.Fonte.Texto $ctx.Cor.TextoSuave 10) }
        & $ctx.DefinirChip 'Aguardando você' 'Atencao'
        $ctx.State.DialogOpen = $true
        $ctx.Lado.Invalidate()
    }.GetNewClosure()

    $ctx.PerguntaTexto = {
        param([string]$Titulo, [string]$Texto, [string]$Exemplo = '', [switch]$Email, [switch]$PermiteVazio)
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        & $ctx.Cabecalho $Titulo $Texto
        $caixa = New-Object System.Windows.Forms.TextBox
        $caixa.Font = $ctx.Fonte.Texto
        $caixa.Dock = 'Fill'
        $caixa.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        & $ctx.Linha $caixa
        $aviso = & $ctx.Rotulo '' $ctx.Fonte.Legenda $ctx.Cor.Critico 6
        $enviar = & $ctx.NovoBotao 'Enviar' 'Principal'
        $acao = {
            $valor = $caixa.Text.Trim()
            if ($Email -and $valor -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $aviso.Text = 'Informe um e-mail válido.'; return }
            & $ctx.Responder $valor
        }.GetNewClosure()
        $enviar.Add_Click($acao)
        $caixa.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; & $acao } }.GetNewClosure())
        & $ctx.Linha $enviar
        if ($Exemplo) { [void](& $ctx.Rotulo $Exemplo $ctx.Fonte.Legenda $ctx.Cor.TextoSuave 0) }
        $caixa.Focus() | Out-Null
    }.GetNewClosure()

    $ctx.PerguntaOpcoes = {
        param([string]$Titulo, [string]$Texto, $Bloco)
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        if (-not $Bloco -or $Bloco.Itens.Count -eq 0) { & $ctx.PerguntaTexto $Titulo $Texto; return }
        & $ctx.Cabecalho $Titulo $Texto
        foreach ($item in $Bloco.Itens) {
            $b = & $ctx.NovoBotao $item.Nome 'Opcao'
            $b.Height = 42
            $b.Tag = $null
            # Sem zero a esquerda: "Acao desejada [1/2/3]" so aceita "1", nao "01".
            # Todas as outras perguntas aceitam o numero sem zero.
            $b.AccessibleDescription = ([int]$item.Numero).ToString()
            $b.Add_Click({ param($s, $e) & $ctx.Responder $s.AccessibleDescription $s.Text }.GetNewClosure())
            & $ctx.Linha $b
        }
    }.GetNewClosure()

    $ctx.PerguntaSimNao = {
        param([string]$Pergunta)
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        $resumo = & $ctx.Resumo
        & $ctx.Cabecalho $Pergunta ''
        if ($resumo.Count -gt 0) {
            $quadro = New-Object System.Windows.Forms.Label
            $quadro.Text = ($resumo -join "`r`n")
            $quadro.Font = $ctx.Fonte.Texto
            $quadro.ForeColor = $ctx.Cor.Texto
            $quadro.BackColor = $ctx.Cor.CamadaAlt
            $quadro.AutoSize = $true
            $quadro.MaximumSize = New-Object System.Drawing.Size(396, 0)
            $quadro.MinimumSize = New-Object System.Drawing.Size(396, 0)
            $quadro.Padding = New-Object System.Windows.Forms.Padding(10)
            $quadro.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
            & $ctx.Linha $quadro
        }
        $sim = & $ctx.NovoBotao 'Sim' 'Principal'
        $sim.Add_Click({ & $ctx.Responder 'S' 'Sim' }.GetNewClosure())
        & $ctx.Linha $sim
        $nao = & $ctx.NovoBotao 'Não' 'Secundario'
        $nao.Add_Click({ & $ctx.Responder 'N' 'Não' }.GetNewClosure())
        & $ctx.Linha $nao
    }.GetNewClosure()

    $ctx.PerguntaPerigo = {
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        $resumo = & $ctx.Resumo
        & $ctx.Cabecalho 'Confirmar exclusão' 'Os participantes serão removidos dos projetos selecionados. Para confirmar, digite exatamente CONFIRMAR EXCLUSAO.'
        if ($resumo.Count -gt 0) { [void](& $ctx.Rotulo ($resumo -join "`r`n") $ctx.Fonte.Texto $ctx.Cor.Texto 10) }
        $caixa = New-Object System.Windows.Forms.TextBox
        $caixa.Font = $ctx.Fonte.Texto
        $caixa.Dock = 'Fill'
        $caixa.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
        & $ctx.Linha $caixa
        $confirmar = & $ctx.NovoBotao 'Excluir participantes' 'Perigo'
        $confirmar.Enabled = $false
        $caixa.Add_TextChanged({ $confirmar.Enabled = ($caixa.Text.Trim() -ceq 'CONFIRMAR EXCLUSAO') }.GetNewClosure())
        $confirmar.Add_Click({ & $ctx.Responder 'CONFIRMAR EXCLUSAO' 'Exclusão confirmada' }.GetNewClosure())
        & $ctx.Linha $confirmar
        $desistir = & $ctx.NovoBotao 'Desistir' 'Secundario'
        $desistir.Add_Click({ & $ctx.Responder '' 'Desistiu da exclusão' }.GetNewClosure())
        & $ctx.Linha $desistir
        $caixa.Focus() | Out-Null
    }.GetNewClosure()

    $ctx.PerguntaContinuar = {
        param([string]$Titulo, [string]$Texto, [string]$Botao = 'Continuar')
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        & $ctx.Cabecalho $Titulo $Texto
        $b = & $ctx.NovoBotao $Botao 'Principal'
        $b.Add_Click({ & $ctx.Responder '' $Botao }.GetNewClosure())
        & $ctx.Linha $b
    }.GetNewClosure()

    $ctx.FiltrarLista = {
        param([string]$Termo)
        $termo = $Termo.Trim().ToLowerInvariant()
        $ctx.ListaVisivel = @($ctx.ListaItens | Where-Object { -not $termo -or ($_.Nome.ToLowerInvariant().Contains($termo)) -or ($_.Numero -eq $termo) })
        $lista = $ctx.ListaControle
        $ctx.Reconstruindo = $true
        $lista.BeginUpdate()
        $lista.Items.Clear()
        foreach ($item in $ctx.ListaVisivel) {
            $idx = $lista.Items.Add(("{0}   {1}" -f $item.Numero, $item.Nome))
            if ($ctx.ListaMultipla -and $ctx.Marcados.Contains($item.Numero)) { $lista.SetItemChecked($idx, $true) }
        }
        $lista.EndUpdate()
        $ctx.Reconstruindo = $false
        & $ctx.AtualizarContagem
    }.GetNewClosure()

    $ctx.AtualizarContagem = {
        if ($ctx.ListaMultipla -and $ctx.Contagem) {
            $ctx.Contagem.Text = "{0} de {1} selecionado(s)" -f $ctx.Marcados.Count, $ctx.ListaItens.Count
            $ctx.ConfirmarLista.Enabled = $ctx.Marcados.Count -gt 0
            $ctx.ConfirmarLista.Text = "Confirmar seleção ({0})" -f $ctx.Marcados.Count
        }
    }.GetNewClosure()

    $ctx.PerguntaLista = {
        param([string]$Titulo, [string]$Texto, $Bloco, [switch]$Multipla, [string]$ValorTodos = '', [object[]]$Atalhos = @())
        $ctx = $ctx   # torna local: closures internas so capturam variaveis locais
        if (-not $Bloco -or $Bloco.Itens.Count -eq 0) { & $ctx.PerguntaTexto $Titulo $Texto; return }
        & $ctx.Cabecalho $Titulo $Texto
        # .ToArray(): @(List) dentro de closure falha no PowerShell 5.1 ("tipos de argumento nao correspondem").
        $ctx.ListaItens = $Bloco.Itens.ToArray()
        $ctx.ListaMultipla = [bool]$Multipla
        $ctx.ValorTodos = $ValorTodos
        $ctx.Marcados.Clear()

        $busca = New-Object System.Windows.Forms.TextBox
        $busca.Font = $ctx.Fonte.Texto
        $busca.Dock = 'Fill'
        $busca.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        & $ctx.Linha $busca

        if ($Multipla) {
            $botoes = New-Object System.Windows.Forms.FlowLayoutPanel
            $botoes.AutoSize = $true
            $botoes.WrapContents = $false
            $botoes.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
            $todos = & $ctx.NovoBotao 'Marcar exibidos' 'Secundario' 130
            $todos.Height = 30
            $todos.Add_Click({
                foreach ($item in $ctx.ListaVisivel) { [void]$ctx.Marcados.Add($item.Numero) }
                & $ctx.FiltrarLista $busca.Text
            }.GetNewClosure())
            $limpar = & $ctx.NovoBotao 'Limpar' 'Secundario' 80
            $limpar.Height = 30
            $limpar.Add_Click({ $ctx.Marcados.Clear(); & $ctx.FiltrarLista $busca.Text }.GetNewClosure())
            $botoes.Controls.Add($todos); $botoes.Controls.Add($limpar)
            & $ctx.Linha $botoes
            $lista = New-Object System.Windows.Forms.CheckedListBox
            $lista.CheckOnClick = $true
        }
        else {
            $lista = New-Object System.Windows.Forms.ListBox
        }
        $lista.Dock = 'Fill'
        $lista.Font = $ctx.Fonte.Texto
        $lista.BorderStyle = 'FixedSingle'
        $lista.IntegralHeight = $false
        $lista.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        $ctx.ListaControle = $lista
        & $ctx.Linha $lista -Preencher

        if ($Multipla) {
            $ctx.Contagem = & $ctx.Rotulo '' $ctx.Fonte.Legenda $ctx.Cor.TextoSuave 6
            $lista.Add_ItemCheck({
                param($s, $e)
                if ($ctx.Reconstruindo) { return }
                $numero = $ctx.ListaVisivel[$e.Index].Numero
                if ($e.NewValue -eq 'Checked') { [void]$ctx.Marcados.Add($numero) } else { [void]$ctx.Marcados.Remove($numero) }
                $ctx.Contagem.Text = "{0} de {1} selecionado(s)" -f $ctx.Marcados.Count, $ctx.ListaItens.Count
                $ctx.ConfirmarLista.Enabled = $ctx.Marcados.Count -gt 0
                $ctx.ConfirmarLista.Text = "Confirmar seleção ({0})" -f $ctx.Marcados.Count
            }.GetNewClosure())
        }

        foreach ($atalho in $Atalhos) {
            $a = & $ctx.NovoBotao $atalho.Texto 'Secundario'
            $a.AccessibleDescription = $atalho.Valor
            $a.Add_Click({ param($s, $e) & $ctx.Responder $s.AccessibleDescription $s.Text }.GetNewClosure())
            & $ctx.Linha $a
        }

        $confirmar = & $ctx.NovoBotao $(if ($Multipla) { 'Confirmar seleção (0)' } else { 'Selecionar' }) 'Principal'
        $ctx.ConfirmarLista = $confirmar
        $enviar = {
            if ($ctx.ListaMultipla) {
                if ($ctx.Marcados.Count -eq 0) { return }
                if ($ctx.ValorTodos -and $ctx.Marcados.Count -eq $ctx.ListaItens.Count) {
                    & $ctx.Responder $ctx.ValorTodos ("Todos os {0} itens" -f $ctx.ListaItens.Count)
                    return
                }
                $numeros = @($ctx.Marcados | ForEach-Object { [int]$_ } | Sort-Object)
                $nomes = @($ctx.ListaItens | Where-Object { $ctx.Marcados.Contains($_.Numero) } | Select-Object -First 3 | ForEach-Object { $_.Nome })
                $exibir = "{0} selecionado(s): {1}{2}" -f $numeros.Count, ($nomes -join '; '), $(if ($numeros.Count -gt 3) { '; ...' } else { '' })
                & $ctx.Responder ($numeros -join ',') $exibir
            }
            else {
                $idx = $ctx.ListaControle.SelectedIndex
                if ($idx -lt 0) { return }
                $item = $ctx.ListaVisivel[$idx]
                & $ctx.Responder ([int]$item.Numero).ToString() $item.Nome
            }
        }.GetNewClosure()
        $confirmar.Add_Click($enviar)
        if (-not $Multipla) { $lista.Add_DoubleClick($enviar) }
        & $ctx.Linha $confirmar

        $busca.Add_TextChanged({ & $ctx.FiltrarLista $busca.Text }.GetNewClosure())
        $busca.Add_KeyDown({
            param($s, $e)
            if ($e.KeyCode -eq 'Down' -and $ctx.ListaControle.Items.Count -gt 0) { $ctx.ListaControle.Focus(); $ctx.ListaControle.SelectedIndex = 0; $e.SuppressKeyPress = $true }
            elseif ($e.KeyCode -eq 'Enter' -and -not $ctx.ListaMultipla -and $ctx.ListaControle.Items.Count -eq 1) { $ctx.ListaControle.SelectedIndex = 0; & $enviar; $e.SuppressKeyPress = $true }
        }.GetNewClosure())
        if ($ctx.DefinirExemplo) { & $ctx.DefinirExemplo $busca 'Filtrar a lista' }
        & $ctx.FiltrarLista ''
        $busca.Focus() | Out-Null
    }.GetNewClosure()

    $ctx.DefinirExemplo = if (Get-Command Set-TemaTextoExemplo -ErrorAction SilentlyContinue) {
        { param($caixa, $texto) try { [void][CentralTema.Dwm]::SendMessage($caixa.Handle, 0x1501, [IntPtr]1, $texto) } catch { } }
    } else { $null }

    # Traduz o texto do console no controle adequado.
    $ctx.Perguntar = {
        param([string]$Prompt)
        $st = $ctx.State
        $buffer = $st.PromptBuffer
        if ($Prompt -match 'Escolha \[1/2\]' -or $Prompt -eq 'FORMA_SELECAO_PROJETOS') {
            $bloco = & $ctx.UltimoBloco ''
            $titulo = if ($bloco -and $bloco.Titulo) { $bloco.Titulo } elseif ($buffer -match 'Forma de selecao') { 'Forma de seleção dos projetos' } else { 'Escolha uma opção' }
            & $ctx.PerguntaOpcoes $titulo '' $bloco
        }
        elseif ($Prompt -eq 'SELECIONAR_CONCESSAO' -or $Prompt -eq 'SELECIONAR_CONCESSAO_BUSCAS') {
            & $ctx.PerguntaLista 'Selecionar concessão' 'Escolha a concessão. Digite para filtrar.' (& $ctx.UltimoBloco 'Concess')
        }
        elseif ($Prompt -eq 'SELECIONAR_PROJETOS') {
            # "Forma de selecao dos projetos" tambem cita projetos, mas e o bloco de opcoes.
            & $ctx.PerguntaLista 'Selecionar projetos' 'Marque os projetos. Use o filtro para encontrar pelo nome.' (& $ctx.UltimoBloco '^(?!.*Forma de selecao).*Projeto') -Multipla -ValorTodos 'todos'
        }
        elseif ($Prompt -eq 'SELECIONAR_PROJETOS_BUSCAS') {
            & $ctx.PerguntaLista 'Selecionar projetos' 'Marque os projetos que receberão as buscas salvas.' (& $ctx.UltimoBloco '^(?!.*Forma de selecao).*Projeto') -Multipla -ValorTodos 'T'
        }
        elseif ($Prompt -eq 'SELECIONAR_BUSCAS_SALVAS') {
            & $ctx.PerguntaLista 'Selecionar buscas salvas' 'Marque as buscas ou use um dos atalhos.' (& $ctx.UltimoBloco 'Buscas') -Multipla -ValorTodos 'T' -Atalhos @(
                [PSCustomObject]@{ Texto = 'Usar as buscas padrão'; Valor = 'P' }
            )
        }
        elseif ($Prompt -match 'E-mail do usuario') {
            & $ctx.PerguntaTexto 'E-mail do usuário' 'Informe o e-mail do usuário que será incluído.' -Email
        }
        elseif ($Prompt -match 'Acao desejada \[1/2/3\]') {
            & $ctx.PerguntaOpcoes 'O que fazer com esse usuário?' '' (& $ctx.UltimoBloco '')
        }
        elseif ($Prompt -match 'Titulo/cargo do participante') {
            & $ctx.PerguntaTexto 'Título do participante' 'Deixe em branco para usar "Participante".' -PermiteVazio
        }
        elseif ($Prompt -match 'Numero ou nome da role') {
            $bloco = & $ctx.UltimoBloco ''
            $atalhos = @()
            if ($bloco -and @($bloco.Itens | Where-Object { $_.Nome -match '\[padrao\]' }).Count -gt 0) {
                $atalhos = @([PSCustomObject]@{ Texto = 'Usar a role padrão (Team Member)'; Valor = '' })
            }
            & $ctx.PerguntaLista 'Selecionar role' 'Role do PW Web aplicada às linhas sem role informada.' $bloco -Atalhos $atalhos
        }
        elseif ($Prompt -match 'Confirmacao:') {
            & $ctx.PerguntaPerigo
        }
        elseif ($Prompt -match 'autenticado') {
            & $ctx.PerguntaContinuar 'Faça login no navegador' 'O navegador Bentley foi aberto. Conclua o login e depois clique em Continuar.' 'Já fiz o login, continuar'
        }
        elseif ($Prompt -match 'Pressione ENTER para fechar') {
            & $ctx.PerguntaContinuar 'Execução finalizada' 'A rotina terminou. Confira o andamento e feche o navegador.' 'Fechar navegador'
        }
        elseif ($Prompt -match '\[S/N\]:') {
            $pergunta = ($Prompt -replace '\s*\[S/N\]:\s*$', '').Trim()
            if (-not $pergunta) { $pergunta = 'Confirmar?' }
            & $ctx.PerguntaSimNao $pergunta
        }
        else {
            & $ctx.PerguntaTexto 'Continuar execução' 'Conclua a etapa indicada no andamento e envie a resposta (deixe vazio para apenas continuar).' -PermiteVazio
        }
    }.GetNewClosure()

    # ---------------------------------------------------------
    # Botoes de acao
    # ---------------------------------------------------------
    $startButton = & $ctx.NovoBotao 'Iniciar rotina' 'Principal' 150
    $startButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $actions.Controls.Add($startButton)
    $ctx.StartButton = $startButton
    $secondaryButton = & $ctx.NovoBotao $SecondaryButtonText 'Secundario' 130
    $secondaryButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $secondaryButton.Visible = -not [string]::IsNullOrWhiteSpace($SecondaryButtonText)
    $actions.Controls.Add($secondaryButton)
    $cancelButton = & $ctx.NovoBotao 'Cancelar execução' 'Secundario' 150
    $cancelButton.Margin = New-Object System.Windows.Forms.Padding(0)
    $cancelButton.Enabled = $false
    $actions.Controls.Add($cancelButton)

    $logsButton = & $ctx.NovoBotao 'Pasta de logs' 'Secundario' 120
    $logsButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $extras.Controls.Add($logsButton)
    $detailsButton = & $ctx.NovoBotao 'Detalhes técnicos  ▾' 'Secundario' 160
    $detailsButton.Margin = New-Object System.Windows.Forms.Padding(0)
    $extras.Controls.Add($detailsButton)

    $logsButton.Add_Click({ Start-Process explorer.exe -ArgumentList ('"{0}"' -f $ctx.LogsDir) }.GetNewClosure())
    $detailsButton.Add_Click({
        param($s, $e)
        $aberto = $ctx.Root.RowStyles[3].Height -gt 0
        $ctx.Root.RowStyles[3].Height = if ($aberto) { 0 } else { 220 }
        $s.Text = if ($aberto) { 'Detalhes técnicos  ▾' } else { 'Detalhes técnicos  ▴' }
    }.GetNewClosure())

    # ---------------------------------------------------------
    # Saida do processo
    # ---------------------------------------------------------
    $appendOutput = {
        param([string]$Line)
        if ($null -eq $Line) { return }
        $output.AppendText($Line + [Environment]::NewLine)
        if ($state.LogPath) {
            try { [IO.File]::AppendAllText($state.LogPath, $Line + [Environment]::NewLine, [Text.UTF8Encoding]::new($true)) } catch {}
        }
    }.GetNewClosure()
    $ctx.AppendOutput = $appendOutput

    $appendProcessOutput = {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return }
        $output.AppendText($Text)
        $state.PromptBuffer += $Text
        $state.LastOutputUtc = [DateTime]::UtcNow
        # Limite folgado: listas longas (ex.: 100+ projetos de uma concessao) precisam
        # manter o titulo ("Projetos encontrados:") no buffer para a pergunta ser reconhecida.
        if ($state.PromptBuffer.Length -gt 100000) {
            $state.PromptBuffer = $state.PromptBuffer.Substring($state.PromptBuffer.Length - 100000)
        }
        if ($state.LogPath) {
            try { [IO.File]::AppendAllText($state.LogPath, $Text, [Text.UTF8Encoding]::new($true)) } catch {}
        }
        # Andamento legivel: so linhas completas
        $ctx.LinhaPendente += $Text
        $partes = $ctx.LinhaPendente -split "`n"
        for ($i = 0; $i -lt $partes.Count - 1; $i++) { & $ctx.ProcessarLinha ($partes[$i].TrimEnd("`r")) }
        $ctx.LinhaPendente = $partes[$partes.Count - 1]
    }.GetNewClosure()

    $startButton.Add_Click({
        $argumentsForRun = [string]$state.NextRuntimeArguments
        $confirmationForRun = [string]$state.NextConfirmationText
        $state.NextRuntimeArguments = $RuntimeArguments
        $state.NextConfirmationText = $StartConfirmationText
        if (-not [string]::IsNullOrWhiteSpace($confirmationForRun)) {
            $confirmation = [Windows.Forms.MessageBox]::Show($panel.FindForm(), $confirmationForRun, 'Confirmar execução', 'YesNo', 'Warning')
            if ($confirmation -ne 'Yes') { & $ctx.Feed 'Execução cancelada antes de começar.' 'Sistema'; return }
        }
        $pythonScript = [IO.Path]::GetFullPath((Join-Path $CentralRoot $PythonRelativePath))
        if (-not (Test-Path -LiteralPath $pythonScript -PathType Leaf)) {
            [Windows.Forms.MessageBox]::Show("Script não encontrado:`r`n$pythonScript", 'Arquivo não encontrado', 'OK', 'Error') | Out-Null
            return
        }
        try {
            $state.LogPath = Join-Path $ctx.LogsDir ("{0}_{1}.log" -f $LogPrefix, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $state.Finished = $false
            $state.DialogOpen = $false
            $state.PromptBuffer = ''
            $ctx.LinhaPendente = ''; $ctx.Blocos.Clear(); $ctx.Recentes.Clear(); $ctx.UltimaEraLista = $false
            $ctx.Ok = 0; $ctx.Erros = 0; $ctx.Avisos = 0; $ctx.UltimosErros.Clear(); $ctx.Cancelado = $false
            $ctx.Etapa = -1; $ctx.EtapaTotal = 0; $ctx.EtapaTexto = ''
            $ctx.Root.RowStyles[1].Height = 0
            $output.Clear()
            $ctx.FeedBox.Clear(); $ctx.FeedVazio = $false
            & $appendOutput "Iniciando rotina: $PanelTitle..."
            $modo = if ([string]::IsNullOrWhiteSpace($argumentsForRun)) { '' } else { " ($($argumentsForRun.Trim()))" }
            & $ctx.Feed ("Iniciada às {0}{1}" -f (Get-Date -Format 'HH:mm'), $modo) 'Sistema'
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            if ($Runtime -eq 'PowerShell') {
                $command = "[Console]::OutputEncoding = [Text.UTF8Encoding]::new(`$false); & '$($pythonScript.Replace("'", "''"))' $argumentsForRun"
                $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand {0}' -f $encodedCommand
            }
            else {
                $pythonExecutable = if (Get-Command Get-CentralPythonExecutable -ErrorAction SilentlyContinue) {
                    Get-CentralPythonExecutable
                }
                else {
                    (Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source
                }
                if (-not $pythonExecutable) { throw 'Python não encontrado. Recrie o ambiente virtual .venv na raiz do projeto.' }
                $psi.FileName = $pythonExecutable
                $psi.Arguments = ('-u "{0}" {1}' -f $pythonScript, $argumentsForRun).Trim()
            }
            $psi.WorkingDirectory = Split-Path -Parent $pythonScript
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.RedirectStandardInput = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
            $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
            $psi.EnvironmentVariables['PYTHONIOENCODING'] = 'utf-8:replace'
            $psi.EnvironmentVariables['PYTHONUTF8'] = '1'
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            # O .NET cria a entrada do processo com Console.InputEncoding e ja grava o
            # BOM no Start(). Troca temporariamente por UTF-8 sem BOM.
            $codificacaoAnterior = $null
            try { $codificacaoAnterior = [Console]::InputEncoding; [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
            try { $iniciou = $process.Start() }
            finally { if ($codificacaoAnterior) { try { [Console]::InputEncoding = $codificacaoAnterior } catch { } } }
            if (-not $iniciou) { throw 'O processo não foi iniciado.' }
            $state.Process = $process
            # Escritor proprio, sem BOM: o padrao do .NET envia um BOM antes da
            # primeira resposta (o script recebia "﻿01" em vez de "01").
            $codificacao = if ($Runtime -eq 'Python') { New-Object System.Text.UTF8Encoding($false) } else { [Console]::InputEncoding }
            if ($codificacao.GetPreamble().Length -gt 0) { $codificacao = New-Object System.Text.UTF8Encoding($false) }
            $entrada = New-Object System.IO.StreamWriter($process.StandardInput.BaseStream, $codificacao)
            $entrada.AutoFlush = $true
            $state | Add-Member -NotePropertyName Input -NotePropertyValue $entrada -Force
            $state.OutputTask = [CentralScripts.StreamPump]::Start($process.StandardOutput, $state.Queue, '')
            $state.ErrorTask = [CentralScripts.StreamPump]::StartFilteredError($process.StandardError, $state.Queue)
            $startButton.Enabled = $false
            $secondaryButton.Enabled = $false
            $cancelButton.Enabled = $true
            & $ctx.DefinirChip 'Em execução' 'Verde'
            & $ctx.MostrarExecucao
        }
        catch {
            & $appendOutput "[ERRO] $($_.Exception.Message)"
            & $ctx.Feed "Falha ao iniciar: $($_.Exception.Message)" 'Erro'
            & $ctx.DefinirChip 'Falha ao iniciar' 'Critico'
        }
    }.GetNewClosure())

    $secondaryButton.Add_Click({
        if ($state.Process -and -not $state.Process.HasExited) { return }
        $state.NextRuntimeArguments = $SecondaryRuntimeArguments
        $state.NextConfirmationText = $SecondaryConfirmationText
        $startButton.PerformClick()
    }.GetNewClosure())

    $cancelButton.Add_Click({
        if ($state.Process -and -not $state.Process.HasExited) {
            if ([Windows.Forms.MessageBox]::Show('Deseja interromper a rotina em execução?', 'Cancelar execução', 'YesNo', 'Warning') -eq 'Yes') {
                $ctx.Cancelado = $true
                $state.Process.Kill()
                & $appendOutput '[AVISO] Cancelamento solicitado pelo operador.'
                & $ctx.Feed 'Cancelamento solicitado pelo operador.' 'Aviso'
            }
        }
    }.GetNewClosure())

    # ---------------------------------------------------------
    # Ciclo: le a fila, detecta perguntas e o fim do processo
    # ---------------------------------------------------------
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 150
    $timer.Add_Tick({
      try {
        $line = $null
        $batch = New-Object System.Text.StringBuilder
        while ($state.Queue.TryDequeue([ref]$line)) {
            [void]$batch.Append($line)
            $line = $null
        }
        if ($batch.Length -gt 0) { & $appendProcessOutput $batch.ToString() }
        if ($ctx.FeedSujo) {
            $ctx.FeedBox.SelectionStart = $ctx.FeedBox.TextLength
            $ctx.FeedBox.ScrollToCaret()
            $output.SelectionStart = $output.TextLength
            $output.ScrollToCaret()
            $ctx.FeedSujo = $false
        }
        if ($state.Process -and -not $state.Process.HasExited -and -not $state.DialogOpen) {
            $promptMatch = [regex]::Match(
                $state.PromptBuffer,
                '(?is)(Escolha \[1/2\]:?|E-mail do usuario:|Continuar com este e-mail\? \[S/N\]:|Continuar com estes usuarios\? \[S/N\]:|Acao desejada \[1/2/3\]:|Titulo/cargo do participante.*?:|Quando estiver autenticado.*?pressione ENTER no terminal\.\.\.|Numero ou nome da role:|Aplicar as operacoes no Bentley RBAC\? \[S/N\]:|Aplicar alteracoes\? \[S/N\]:|Confirmacao:|[^\r\n]+\[S/N\]:|Pressione ENTER para fechar (?:o navegador|esta janela)(?:\.\.\.)?)\s*$'
            )
            $chave = $null
            if ($promptMatch.Success) { $chave = $promptMatch.Value.Trim() }
            elseif (([DateTime]::UtcNow - $state.LastOutputUtc).TotalMilliseconds -ge 700) {
                if ($state.PromptBuffer -match '(?is)Digite numeros separados por virgula.*intervalos como 1-5.*ou.*todos') { $chave = 'SELECIONAR_PROJETOS' }
                elseif ($state.PromptBuffer -match '(?is)Projetos encontrados:.*\[\d+\].*') { $chave = 'SELECIONAR_PROJETOS_BUSCAS' }
                elseif ($state.PromptBuffer -match '(?is)Concessoes encontradas:.*\[\d+\].*') { $chave = 'SELECIONAR_CONCESSAO_BUSCAS' }
                elseif ($state.PromptBuffer -match '(?is)Buscas disponiveis para criar:.*\[\d+\].*') { $chave = 'SELECIONAR_BUSCAS_SALVAS' }
                elseif ($state.PromptBuffer -match '(?is)Forma de selecao dos projetos:.*01\..*02\.') { $chave = 'FORMA_SELECAO_PROJETOS' }
                elseif ($state.PromptBuffer -match '(?is)Concessoes disponiveis:.*\r?\n\s*\d{2}\.\s+') { $chave = 'SELECIONAR_CONCESSAO' }
            }
            if ($chave) {
                # Linha da pergunta ainda sem quebra: fecha a lista anterior antes de montar o controle.
                & $ctx.FecharBlocoLista
                try { & $ctx.Perguntar $chave }
                catch {
                    & $ctx.Feed "Não foi possível montar a pergunta ($($_.Exception.Message)). Use a caixa de texto." 'Aviso'
                    & $ctx.PerguntaTexto 'Continuar execução' $chave -PermiteVazio
                }
            }
        }
        if ($state.Process -and $state.Process.HasExited -and -not $state.Finished) {
            $state.Finished = $true
            $state.DialogOpen = $false
            if ($ctx.LinhaPendente) { & $ctx.ProcessarLinha $ctx.LinhaPendente; $ctx.LinhaPendente = '' }
            & $ctx.FecharBlocoLista
            $code = $state.Process.ExitCode
            & $appendOutput "Processo finalizado com código $code."
            if ($ctx.EtapaTotal -gt 0 -and $code -eq 0 -and -not $ctx.Cancelado) { $ctx.Etapa = $ctx.EtapaTotal; $ctx.Stepper.Invalidate() }
            & $ctx.Feed ("Finalizada às {0}." -f (Get-Date -Format 'HH:mm')) 'Sistema'
            $startButton.Enabled = $true; $secondaryButton.Enabled = $true; $cancelButton.Enabled = $false
            $state.Process.Dispose(); $state.Process = $null
            & $ctx.MostrarFim $code
            $ctx.FeedBox.SelectionStart = $ctx.FeedBox.TextLength
            $ctx.FeedBox.ScrollToCaret()
        }
      }
      catch {
        # Nunca deixa uma falha de tela derrubar a Central; registra no console.
        try { $output.AppendText("[PAINEL] $($_.Exception.Message)" + [Environment]::NewLine) } catch { }
      }
    }.GetNewClosure())
    $timer.Start()

    $returnToCentral = {
        if ($state.Process -and -not $state.Process.HasExited) {
            [Windows.Forms.MessageBox]::Show('Cancele ou conclua a execução antes de voltar à Central.', 'Rotina em execução', 'OK', 'Information') | Out-Null
            return
        }
        $timer.Stop(); $timer.Dispose(); & $OnBack
    }.GetNewClosure()
    $backButton.Add_Click($returnToCentral)

    # Estado inicial
    & $ctx.DefinirChip 'Pronto' 'Neutro'
    & $ctx.MostrarRepouso
    $feed.SelectionColor = $cor.TextoSuave
    $feed.AppendText('O andamento aparece aqui quando a rotina for iniciada.')
    $ctx.FeedVazio = $true

    return $panel
}
