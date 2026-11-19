# =========================================================
# TEMA VISUAL DA CENTRAL DE SCRIPTS
# Cores e tipografia conforme o Ecovias Brand Book (abril/2025):
#   1.12 Paleta primaria  - tres verdes EcoRodovias
#   1.13 Paleta secundaria - apoio, nunca mais que 20% da area
#   1.14 Tipografia       - Lexend (Verdana como fonte de sistema)
#   1.17 Elementos        - "janela para o mundo"
# =========================================================

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function New-TemaCor([int]$R, [int]$G, [int]$B) { [System.Drawing.Color]::FromArgb($R, $G, $B) }

$script:Tema = [ordered]@{
    # Paleta primaria (brand book 1.12)
    VerdeEscuro   = New-TemaCor 5 107 40      # #056B28 - barra de titulo, botao principal, titulos
    VerdeMedio    = New-TemaCor 29 160 19     # #1DA013 - selecao, foco, item ativo
    VerdeClaro    = New-TemaCor 114 210 22    # #72D216 - detalhes graficos
    # Tons derivados (tints de 20% e 10% previstos no brand book)
    VerdeSuave    = New-TemaCor 228 243 226   # fundo do icone, filtro ativo
    VerdeHover    = New-TemaCor 4 88 33       # botao principal com mouse em cima

    # Neutros com leve vies verde
    Fundo         = New-TemaCor 243 246 244
    Camada        = New-TemaCor 255 255 255
    CamadaAlt     = New-TemaCor 248 250 248
    Contorno      = New-TemaCor 221 229 223
    ContornoForte = New-TemaCor 176 196 182
    Texto         = New-TemaCor 23 37 28
    TextoSuave    = New-TemaCor 92 107 98

    # Semanticas (secundaria do brand book para atencao; vermelho so para critico)
    Critico       = New-TemaCor 196 43 28
    CriticoSuave  = New-TemaCor 253 231 229
    Atencao       = New-TemaCor 168 70 12     # laranja #F56612 escurecido para leitura
    AtencaoSuave  = New-TemaCor 254 236 222
    Info          = New-TemaCor 0 92 255      # azul #005CFF da paleta secundaria
}

# ---------------------------------------------------------
# FONTES
# Lexend fica em .\Fontes e e carregada so para este processo
# (nada e instalado no Windows). Sem o arquivo, usa Verdana,
# conforme brand book 1.15.
# ---------------------------------------------------------
if (-not ('CentralTema.FontesPrivadas' -as [type])) {
    Add-Type -Namespace CentralTema -Name FontesPrivadas -MemberDefinition @'
[DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
public static extern int AddFontResourceEx(string lpszFilename, uint fl, IntPtr pdv);
'@
}
if (-not ('CentralTema.Dwm' -as [type])) {
    Add-Type -Namespace CentralTema -Name Dwm -MemberDefinition @'
[DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
'@
}

$script:TemaFontes = New-Object System.Drawing.Text.PrivateFontCollection
$script:TemaFamilias = @{}

function Initialize-TemaFontes {
    param([Parameter(Mandatory)][string]$Raiz)

    $pasta = Join-Path $Raiz 'Fontes'
    foreach ($arquivo in @('Lexend-Regular.ttf', 'Lexend-SemiBold.ttf')) {
        $caminho = Join-Path $pasta $arquivo
        if (-not (Test-Path -LiteralPath $caminho)) { continue }
        try {
            # FR_PRIVATE (0x10): a fonte existe so para este processo. O GDI
            # (usado pelos Labels) precisa disso; o GDI+ usa a colecao privada.
            [void][CentralTema.FontesPrivadas]::AddFontResourceEx($caminho, 0x10, [IntPtr]::Zero)
            $script:TemaFontes.AddFontFile($caminho)
        }
        catch { }
    }
    foreach ($familia in $script:TemaFontes.Families) {
        $script:TemaFamilias[$familia.Name] = $familia
    }

    $instaladas = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
    $script:TemaFonteIcones = if ($instaladas -contains 'Segoe Fluent Icons') { 'Segoe Fluent Icons' } else { 'Segoe MDL2 Assets' }
    $script:TemaFonteTexto = if ($instaladas -contains 'Segoe UI Variable Text') { 'Segoe UI Variable Text' } else { 'Segoe UI' }
}

function Get-TemaFonte {
    <#
      Papel:
        Titulo / Subtitulo / Destaque -> Lexend SemiBold (marca)
        Marca                         -> Lexend Regular
        Texto / Legenda               -> Segoe UI (texto de interface do Windows)
        Icone                         -> Segoe Fluent Icons
    #>
    param(
        [ValidateSet('Titulo', 'Subtitulo', 'Destaque', 'Marca', 'Texto', 'TextoForte', 'Legenda', 'Icone')]
        [string]$Papel,
        [single]$Tamanho = 0
    )

    $padrao = @{ Titulo = 16; Subtitulo = 12; Destaque = 10.5; Marca = 10; Texto = 9.5; TextoForte = 9.5; Legenda = 8.5; Icone = 14 }
    if ($Tamanho -le 0) { $Tamanho = $padrao[$Papel] }

    switch ($Papel) {
        { $_ -in @('Titulo', 'Subtitulo', 'Destaque') } {
            if ($script:TemaFamilias.ContainsKey('Lexend SemiBold')) {
                return New-Object System.Drawing.Font($script:TemaFamilias['Lexend SemiBold'], $Tamanho, [System.Drawing.FontStyle]::Regular)
            }
            return New-Object System.Drawing.Font('Verdana', ($Tamanho - 1), [System.Drawing.FontStyle]::Bold)
        }
        'Marca' {
            if ($script:TemaFamilias.ContainsKey('Lexend')) {
                return New-Object System.Drawing.Font($script:TemaFamilias['Lexend'], $Tamanho, [System.Drawing.FontStyle]::Regular)
            }
            return New-Object System.Drawing.Font('Verdana', ($Tamanho - 1))
        }
        'Icone'      { return New-Object System.Drawing.Font($script:TemaFonteIcones, $Tamanho) }
        'TextoForte' { return New-Object System.Drawing.Font(($script:TemaFonteTexto -replace ' Text$', ' Text Semibold'), $Tamanho) }
        default      { return New-Object System.Drawing.Font($script:TemaFonteTexto, $Tamanho) }
    }
}

# ---------------------------------------------------------
# JANELA (Windows 11)
# ---------------------------------------------------------
function Set-TemaJanela {
    param([Parameter(Mandatory)][System.Windows.Forms.Form]$Form)

    # Valor calculado aqui fora: dentro da closure do evento $script:Tema nao existe.
    $c = $script:Tema.VerdeEscuro
    # [int] obrigatorio: Color.R/G/B sao [byte] e -shl em byte estoura para 0.
    $corBarra = ([int]$c.B -shl 16) -bor ([int]$c.G -shl 8) -bor [int]$c.R   # COLORREF 0x00BBGGRR
    $aplicar = {
        param($janela)
        try {
            $h = $janela.Handle
            $canto = 2                                   # DWMWA_WINDOW_CORNER_PREFERENCE = arredondado
            [void][CentralTema.Dwm]::DwmSetWindowAttribute($h, 33, [ref]$canto, 4)
            $cor = $corBarra
            [void][CentralTema.Dwm]::DwmSetWindowAttribute($h, 35, [ref]$cor, 4)        # DWMWA_CAPTION_COLOR
            $branco = 0x00FFFFFF
            [void][CentralTema.Dwm]::DwmSetWindowAttribute($h, 36, [ref]$branco, 4)     # DWMWA_TEXT_COLOR
            $borda = $corBarra
            [void][CentralTema.Dwm]::DwmSetWindowAttribute($h, 34, [ref]$borda, 4)      # DWMWA_BORDER_COLOR
        }
        catch { }
    }.GetNewClosure()
    if ($Form.IsHandleCreated) { & $aplicar $Form }
    $Form.Add_HandleCreated({ param($janela, $e) & $aplicar $janela }.GetNewClosure())
}

function Set-TemaTextoExemplo {
    # Texto de exemplo nativo (EM_SETCUEBANNER) em TextBox.
    param([System.Windows.Forms.TextBox]$Caixa, [string]$Texto)
    $definir = { param($c) [void][CentralTema.Dwm]::SendMessage($c.Handle, 0x1501, [IntPtr]1, $Texto) }.GetNewClosure()
    if ($Caixa.IsHandleCreated) { & $definir $Caixa }
    $Caixa.Add_HandleCreated({ param($c, $e) & $definir $c }.GetNewClosure())
}

function Enable-TemaBufferDuplo {
    param([System.Windows.Forms.Control]$Controle)
    $prop = $Controle.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic,Instance')
    if ($prop) { $prop.SetValue($Controle, $true, $null) }
}

# ---------------------------------------------------------
# DESENHO
# ---------------------------------------------------------
function New-TemaRetanguloArredondado {
    param([System.Drawing.RectangleF]$R, [single]$Raio)

    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [Math]::Min($Raio * 2, [Math]::Min($R.Width, $R.Height))
    if ($d -le 0) { $p.AddRectangle($R); return $p }
    $p.AddArc($R.X, $R.Y, $d, $d, 180, 90)
    $p.AddArc($R.Right - $d, $R.Y, $d, $d, 270, 90)
    $p.AddArc($R.Right - $d, $R.Bottom - $d, $d, $d, 0, 90)
    $p.AddArc($R.X, $R.Bottom - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Draw-TemaJanela {
    # Forma "janela para o mundo" (brand book 1.17): base reta, lado direito
    # reto, topo levemente inclinado e lado esquerdo em diagonal, cantos suaves.
    param(
        [System.Drawing.Graphics]$G,
        [System.Drawing.RectangleF]$R,
        [System.Drawing.Color]$Cor
    )
    $raio = [single]([Math]::Max(2, $R.Height * 0.14))
    $x = $R.X + $raio; $y = $R.Y + $raio; $w = $R.Width - 2 * $raio; $h = $R.Height - 2 * $raio
    $pts = [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF($x, ($y + $h))),
        (New-Object System.Drawing.PointF(($x + $w * 0.30), ($y + $h * 0.06))),
        (New-Object System.Drawing.PointF(($x + $w), $y)),
        (New-Object System.Drawing.PointF(($x + $w), ($y + $h)))
    )
    $pincel = New-Object System.Drawing.SolidBrush($Cor)
    $caneta = New-Object System.Drawing.Pen($Cor, ($raio * 2))
    $caneta.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    try {
        $G.FillPolygon($pincel, $pts)
        $G.DrawPolygon($caneta, $pts)
    }
    finally { $pincel.Dispose(); $caneta.Dispose() }
}

function Draw-TemaSelo {
    # Selo em formato de pilula. Retorna a largura ocupada.
    param(
        [System.Drawing.Graphics]$G,
        [string]$Texto,
        [System.Drawing.Font]$Fonte,
        [int]$X, [int]$Y,
        [System.Drawing.Color]$Fundo,
        [System.Drawing.Color]$Cor,
        [switch]$Ponto
    )
    $flags = [System.Windows.Forms.TextFormatFlags]'NoPadding,SingleLine'
    $tam = [System.Windows.Forms.TextRenderer]::MeasureText($G, $Texto, $Fonte, (New-Object System.Drawing.Size(1000, 100)), $flags)
    $extra = if ($Ponto) { 10 } else { 0 }
    $largura = $tam.Width + 16 + $extra
    $altura = $tam.Height + 6
    $r = New-Object System.Drawing.RectangleF($X, $Y, $largura, $altura)
    $path = New-TemaRetanguloArredondado -R $r -Raio ($altura / 2)
    $pincel = New-Object System.Drawing.SolidBrush($Fundo)
    try {
        $G.FillPath($pincel, $path)
        if ($Ponto) {
            $pp = New-Object System.Drawing.SolidBrush($Cor)
            $G.FillEllipse($pp, ($X + 8), ($Y + $altura / 2 - 3), 6, 6)
            $pp.Dispose()
        }
        [System.Windows.Forms.TextRenderer]::DrawText($G, $Texto, $Fonte, (New-Object System.Drawing.Point(($X + 8 + $extra), ($Y + 3))), $Cor, $flags)
    }
    finally { $pincel.Dispose(); $path.Dispose() }
    return [int]$largura
}

# ---------------------------------------------------------
# TEMA APLICADO A UMA TELA PRONTA
# Usado nas rotinas que embutem a janela do script original (02, 05) e na 01.
# So muda cores, bordas e a fonte do titulo: posicoes e tamanhos ficam iguais,
# para nao quebrar o layout absoluto dos scripts originais.
# ---------------------------------------------------------
$script:TemaCoresBotao = @{}

# Cores antigas usadas pelas telas -> cores do tema
function Convert-TemaCorAntiga {
    param([System.Drawing.Color]$Cor)
    $chave = '{0},{1},{2}' -f $Cor.R, $Cor.G, $Cor.B
    switch ($chave) {
        '22,54,92'   { return $script:Tema.VerdeEscuro }   # azul-marinho antigo
        '31,105,178' { return $script:Tema.VerdeEscuro }   # azul de botao antigo
        '70,78,88'   { return $script:Tema.TextoSuave }
        '75,82,90'   { return $script:Tema.TextoSuave }
        '170,80,20'  { return $script:Tema.Atencao }
    }
    return $null
}

$script:TemaPinturaGrupo = {
    param($grupo, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $fundoPai = if ($grupo.Parent) { $grupo.Parent.BackColor } else { $script:Tema.Fundo }
    $g.Clear($fundoPai)
    $r = New-Object System.Drawing.RectangleF(0.5, 0.5, ($grupo.Width - 1.5), ($grupo.Height - 1.5))
    $path = New-TemaRetanguloArredondado -R $r -Raio 8
    $pincel = New-Object System.Drawing.SolidBrush($script:Tema.Camada)
    $caneta = New-Object System.Drawing.Pen($script:Tema.Contorno)
    $g.FillPath($pincel, $path)
    $g.DrawPath($caneta, $path)
    $pincel.Dispose(); $caneta.Dispose(); $path.Dispose()
    if ($grupo.Text) {
        $flags = [System.Windows.Forms.TextFormatFlags]'NoPadding,SingleLine,EndEllipsis'
        $area = New-Object System.Drawing.Rectangle(10, 2, ($grupo.Width - 20), 20)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $grupo.Text, $script:TemaFonteGrupo, $area, $script:Tema.VerdeEscuro, $flags)
    }
}

$script:TemaBotaoHabilitado = {
    param($botao, $e)
    $cores = $script:TemaCoresBotao[$botao]
    if (-not $cores) { return }
    if ($botao.Enabled) {
        $botao.BackColor = $cores.Fundo
        $botao.FlatAppearance.BorderColor = $cores.Borda
    }
    else {
        $botao.BackColor = $script:Tema.CamadaAlt
        $botao.FlatAppearance.BorderColor = $script:Tema.Contorno
    }
}

$script:TemaControleAdicionado = {
    param($pai, $e)
    Set-TemaControles -Raiz $e.Control -SemTitulo
}

function Set-TemaBotaoGenerico {
    param([System.Windows.Forms.Button]$Botao)

    $texto = ($Botao.Text -replace '&', '').Trim()
    $corAtual = $Botao.BackColor
    $jaVerde = ($corAtual.R -eq 5 -and $corAtual.G -eq 107 -and $corAtual.B -eq 40) -or ($null -ne (Convert-TemaCorAntiga $corAtual) -and $Botao.ForeColor.ToArgb() -eq [System.Drawing.Color]::White.ToArgb())

    if ($texto -match '^(<|←)?\s*Voltar') {
        $Botao.Text = '←  Voltar'
        $tipo = 'Voltar'
    }
    elseif ($texto -match '^(Excluir|Desativar)') { $tipo = 'Perigo' }
    elseif ($jaVerde -or $texto -match '^(Buscar usu|Adicionar usu|Selecionar conces|Confirmar sele|Executar|Validar planilha|Iniciar|Conectar|Consultar)') { $tipo = 'Principal' }
    else { $tipo = 'Secundario' }

    $Botao.FlatStyle = 'Flat'
    $Botao.UseVisualStyleBackColor = $false
    $Botao.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Botao.FlatAppearance.BorderSize = 1
    switch ($tipo) {
        'Principal' {
            $fundo = $script:Tema.VerdeEscuro; $borda = $script:Tema.VerdeEscuro
            $Botao.ForeColor = [System.Drawing.Color]::White
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.VerdeHover
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeHover
        }
        'Perigo' {
            $fundo = $script:Tema.Camada; $borda = $script:Tema.Critico
            $Botao.ForeColor = $script:Tema.Critico
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.CriticoSuave
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.CriticoSuave
        }
        'Voltar' {
            $fundo = if ($Botao.Parent) { $Botao.Parent.BackColor } else { $script:Tema.Fundo }
            $borda = $fundo
            $Botao.FlatAppearance.BorderSize = 0
            $Botao.ForeColor = $script:Tema.VerdeEscuro
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.VerdeSuave
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeSuave
            $Botao.Font = New-Object System.Drawing.Font($Botao.Font.FontFamily, $Botao.Font.Size, [System.Drawing.FontStyle]::Bold)
        }
        default {
            $fundo = $script:Tema.Camada; $borda = $script:Tema.Contorno
            $Botao.ForeColor = $script:Tema.Texto
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.VerdeSuave
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeSuave
        }
    }
    $Botao.BackColor = $fundo
    $Botao.FlatAppearance.BorderColor = $borda
    if (-not $script:TemaCoresBotao.ContainsKey($Botao)) { $Botao.Add_EnabledChanged($script:TemaBotaoHabilitado) }
    $script:TemaCoresBotao[$Botao] = @{ Fundo = $fundo; Borda = $borda }
    & $script:TemaBotaoHabilitado $Botao $null
}

function Set-TemaControles {
    <#
      Percorre a tela e aplica o tema. -Titulo troca o texto do maior rotulo
      (o titulo da janela original) para padronizar com as demais rotinas.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Raiz,
        [string]$Titulo = '',
        [switch]$SemTitulo
    )
    if (-not $script:TemaFonteGrupo) { $script:TemaFonteGrupo = Get-TemaFonte -Papel TextoForte -Tamanho 9.5 }

    $padroes = @(
        [System.Drawing.SystemColors]::Control.ToArgb(),
        [System.Drawing.Color]::FromArgb(245, 247, 250).ToArgb(),
        [System.Drawing.Color]::FromArgb(240, 240, 240).ToArgb()
    )

    $pilha = New-Object System.Collections.Generic.Stack[System.Windows.Forms.Control]
    $pilha.Push($Raiz)
    $maiorRotulo = $null
    while ($pilha.Count -gt 0) {
        $c = $pilha.Pop()
        $corAntiga = Convert-TemaCorAntiga $c.BackColor
        if ($corAntiga) { $c.BackColor = $corAntiga }
        $foreAntiga = Convert-TemaCorAntiga $c.ForeColor
        if ($foreAntiga) { $c.ForeColor = $foreAntiga }

        if ($c -is [System.Windows.Forms.Form]) {
            $c.BackColor = $script:Tema.Fundo
        }
        elseif ($c -is [System.Windows.Forms.GroupBox]) {
            $c.BackColor = $script:Tema.Camada
            $c.ForeColor = $script:Tema.Texto
            Enable-TemaBufferDuplo -Controle $c
            $c.Remove_Paint($script:TemaPinturaGrupo)
            $c.Add_Paint($script:TemaPinturaGrupo)
        }
        elseif ($c -is [System.Windows.Forms.Button]) {
            Set-TemaBotaoGenerico -Botao $c
        }
        elseif ($c -is [System.Windows.Forms.DataGridView]) {
            $c.BackgroundColor = $script:Tema.Camada
            $c.BorderStyle = 'FixedSingle'
            $c.GridColor = $script:Tema.Contorno
            $c.EnableHeadersVisualStyles = $false
            $c.ColumnHeadersBorderStyle = 'Single'
            $c.ColumnHeadersDefaultCellStyle.BackColor = $script:Tema.Fundo
            $c.ColumnHeadersDefaultCellStyle.ForeColor = $script:Tema.Texto
            $c.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:Tema.Fundo
            $c.ColumnHeadersDefaultCellStyle.Font = $script:TemaFonteGrupo
            $c.DefaultCellStyle.SelectionBackColor = $script:Tema.VerdeSuave
            $c.DefaultCellStyle.SelectionForeColor = $script:Tema.Texto
            $c.AlternatingRowsDefaultCellStyle.BackColor = $script:Tema.CamadaAlt
            $c.RowHeadersDefaultCellStyle.BackColor = $script:Tema.Fundo
        }
        elseif ($c -is [System.Windows.Forms.TextBox] -or $c -is [System.Windows.Forms.ListBox] -or $c -is [System.Windows.Forms.TreeView] -or $c -is [System.Windows.Forms.RichTextBox]) {
            if ($c.BorderStyle -eq 'Fixed3D') { $c.BorderStyle = 'FixedSingle' }
            if ($c -is [System.Windows.Forms.TreeView]) { $c.LineColor = $script:Tema.Contorno }
        }
        elseif ($c -is [System.Windows.Forms.StatusStrip]) {
            $c.BackColor = $script:Tema.Fundo
        }
        elseif ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.CheckBox] -or $c -is [System.Windows.Forms.RadioButton]) {
            if ($c.ForeColor.ToArgb() -eq [System.Drawing.SystemColors]::ControlText.ToArgb()) { $c.ForeColor = $script:Tema.Texto }
            if ($c -is [System.Windows.Forms.Label] -and $c.Font.Size -ge 14 -and (-not $maiorRotulo -or $c.Font.Size -gt $maiorRotulo.Font.Size)) { $maiorRotulo = $c }
        }
        elseif ($c -is [System.Windows.Forms.Panel] -or $c -is [System.Windows.Forms.TableLayoutPanel] -or $c -is [System.Windows.Forms.FlowLayoutPanel] -or $c -is [System.Windows.Forms.TabPage] -or $c -is [System.Windows.Forms.SplitContainer]) {
            # Fundo cinza padrao explicito: segue o fundo do pai ja tematizado.
            if ($c.Parent -and $padroes -contains $c.BackColor.ToArgb()) { $c.BackColor = $c.Parent.BackColor }
        }

        if ($c.Controls.Count -gt 0 -and -not ($c -is [System.Windows.Forms.DataGridView]) -and -not ($c -is [System.Windows.Forms.NumericUpDown])) {
            $c.Remove_ControlAdded($script:TemaControleAdicionado)
            $c.Add_ControlAdded($script:TemaControleAdicionado)
            for ($i = $c.Controls.Count - 1; $i -ge 0; $i--) { $pilha.Push($c.Controls[$i]) }
        }
    }

    if (-not $SemTitulo -and $maiorRotulo) {
        if ($Titulo) { $maiorRotulo.Text = $Titulo }
        $maiorRotulo.ForeColor = $script:Tema.VerdeEscuro
        if ($maiorRotulo.AutoSize) { $maiorRotulo.Font = Get-TemaFonte -Papel Titulo -Tamanho 15 }
    }
}

function Set-TemaLayoutEtapas {
    <#
      Layout que se ajusta a altura para as janelas originais embutidas (02, 05).
      Os originais empilham as etapas com tamanho minimo fixo (760-810px) e eram
      cortados em telas com escala (ex.: 1920x1080 a 125% = 793px maximizado).

        - Superiores: etapas de altura fixa, empilhadas no topo
        - Principal:  etapa que recebe a sobra de altura (lista/grade)
        - Inferiores: 1 ou 2 blocos lado a lado no rodape

      Posicao e tamanho das etapas sao calculados aqui (ancoras Top/Left); os
      controles internos continuam seguindo as ancoras originais.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Form]$Janela,
        [System.Windows.Forms.Control[]]$Superiores = @(),
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Principal,
        [System.Windows.Forms.Control[]]$Inferiores = @(),
        [int]$Topo = 55,
        [int]$AlturaInferiorMin = 80,
        [int]$AlturaInferiorMax = 200,
        [double]$ProporcaoInferior = 0.58,
        [int]$AlturaPrincipalMin = 200,
        [int]$LarguraMinima = 1000
    )

    $ancora = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    foreach ($g in @($Superiores) + @($Principal) + @($Inferiores)) { if ($g) { $g.Anchor = $ancora } }
    $barraStatus = @($Janela.Controls | Where-Object { $_ -is [System.Windows.Forms.StatusStrip] }) | Select-Object -First 1

    $alturaSuperiores = 0
    foreach ($g in $Superiores) { $alturaSuperiores += $g.Height + 10 }
    $alturaMinima = $Topo + $alturaSuperiores + $AlturaPrincipalMin + 12 + $AlturaInferiorMin + 8 + $(if ($barraStatus) { $barraStatus.Height } else { 0 })
    $Janela.MinimumSize = New-Object System.Drawing.Size($LarguraMinima, $alturaMinima)
    $Janela.Dock = [System.Windows.Forms.DockStyle]::Fill

    $superioresLocal = @($Superiores)
    $inferioresLocal = @($Inferiores)
    $ajustar = {
        $w = $Janela.ClientSize.Width
        $h = $Janela.ClientSize.Height
        $barra = if ($barraStatus -and $barraStatus.Visible) { $barraStatus.Height } else { 0 }
        $margem = 20; $espaco = 12
        $util = $w - (2 * $margem)
        $y = $Topo
        foreach ($g in $superioresLocal) { $g.SetBounds($margem, $y, $util, $g.Height); $y += $g.Height + 10 }

        $alturaInf = 0
        if ($inferioresLocal.Count -gt 0) {
            $alturaInf = [Math]::Max($AlturaInferiorMin, [Math]::Min($AlturaInferiorMax, [int](($h - $y) * 0.30)))
            $topoInf = $h - $barra - 8 - $alturaInf
            if ($inferioresLocal.Count -eq 1) {
                $inferioresLocal[0].SetBounds($margem, $topoInf, $util, $alturaInf)
            }
            else {
                $larguraA = [int](($util - $espaco) * $ProporcaoInferior)
                $inferioresLocal[0].SetBounds($margem, $topoInf, $larguraA, $alturaInf)
                $inferioresLocal[1].SetBounds(($margem + $larguraA + $espaco), $topoInf, ($util - $larguraA - $espaco), $alturaInf)
            }
        }
        else { $topoInf = $h - $barra - 8 + $espaco }

        $Principal.SetBounds($margem, $y, $util, [Math]::Max($AlturaPrincipalMin, ($topoInf - $espaco - $y)))
    }.GetNewClosure()
    $Janela.Add_Resize($ajustar)
    & $ajustar
}

function Set-TemaBotao {
    param(
        [System.Windows.Forms.Button]$Botao,
        [ValidateSet('Principal', 'Secundario', 'Discreto')][string]$Tipo = 'Secundario'
    )
    $Botao.FlatStyle = 'Flat'
    $Botao.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Botao.UseVisualStyleBackColor = $false
    switch ($Tipo) {
        'Principal' {
            $Botao.BackColor = $script:Tema.VerdeEscuro
            $Botao.ForeColor = [System.Drawing.Color]::White
            $Botao.FlatAppearance.BorderColor = $script:Tema.VerdeEscuro
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.VerdeHover
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeHover
            $Botao.Font = Get-TemaFonte -Papel TextoForte -Tamanho 10
        }
        'Secundario' {
            $Botao.BackColor = $script:Tema.Camada
            $Botao.ForeColor = $script:Tema.Texto
            $Botao.FlatAppearance.BorderColor = $script:Tema.Contorno
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.CamadaAlt
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeSuave
            $Botao.Font = Get-TemaFonte -Papel Texto
        }
        'Discreto' {
            $Botao.BackColor = $script:Tema.Fundo
            $Botao.ForeColor = $script:Tema.Texto
            $Botao.FlatAppearance.BorderSize = 0
            $Botao.FlatAppearance.MouseOverBackColor = $script:Tema.VerdeSuave
            $Botao.FlatAppearance.MouseDownBackColor = $script:Tema.VerdeSuave
            $Botao.Font = Get-TemaFonte -Papel Texto
        }
    }
}
