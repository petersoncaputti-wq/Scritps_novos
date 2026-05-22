# =========================================================
# PROJECTWISE - USUARIOS INATIVOS COM INTERFACE VISUAL
# Relatorio, desativacao e exclusao com travas de seguranca.
# =========================================================

Clear-Host
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$script:ConectadoPW = $false
$script:UsuarioAtual = $null
$script:Resultados = @()
$script:InativosElegiveis = @()

function Write-UiLog {
    param([string]$Mensagem)

    if ($script:LogBox) {
        $script:LogBox.AppendText(("[$(Get-Date -Format 'HH:mm:ss')] {0}" -f $Mensagem) + [Environment]::NewLine)
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Show-UiError {
    param([string]$Mensagem)

    [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-UiInfo {
    param([string]$Mensagem)

    [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        "Informacao",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Ask-UiYesNo {
    param(
        [string]$Mensagem,
        [string]$Titulo = "Confirmacao"
    )

    $resposta = [System.Windows.Forms.MessageBox]::Show(
        $Mensagem,
        $Titulo,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($resposta -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Confirmar-TextoDialog {
    param(
        [string]$TextoEsperado,
        [string]$Mensagem
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Confirmacao obrigatoria"
    $dialog.Size = New-Object System.Drawing.Size(560, 210)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Mensagem
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $label.Size = New-Object System.Drawing.Size(510, 55)
    $dialog.Controls.Add($label)

    $labelTexto = New-Object System.Windows.Forms.Label
    $labelTexto.Text = "Digite exatamente '$TextoEsperado' para confirmar:"
    $labelTexto.Location = New-Object System.Drawing.Point(15, 78)
    $labelTexto.Size = New-Object System.Drawing.Size(510, 22)
    $dialog.Controls.Add($labelTexto)

    $textConfirmacao = New-Object System.Windows.Forms.TextBox
    $textConfirmacao.Location = New-Object System.Drawing.Point(18, 104)
    $textConfirmacao.Size = New-Object System.Drawing.Size(505, 25)
    $dialog.Controls.Add($textConfirmacao)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Confirmar"
    $btnOk.Location = New-Object System.Drawing.Point(335, 138)
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnOk)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(433, 138)
    $btnCancelar.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnCancelar)

    $dialog.AcceptButton = $btnOk
    $dialog.CancelButton = $btnCancelar

    $resultado = $dialog.ShowDialog($form)
    if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
        return $false
    }

    return ($textConfirmacao.Text -ceq $TextoEsperado)
}

function Get-ValorSeguroPropriedade {
    param(
        [object]$Objeto,
        [string[]]$PossiveisNomes
    )

    if (-not $Objeto) {
        return ""
    }

    foreach ($nome in $PossiveisNomes) {
        $propriedade = $Objeto.PSObject.Properties[$nome]
        if ($propriedade -and $null -ne $propriedade.Value) {
            $valor = $propriedade.Value.ToString().Trim()
            if ($valor -ne "") {
                return $valor
            }
        }
    }

    return ""
}

function Get-StatusUsuario {
    param(
        [object]$UltimoAcesso,
        [datetime]$DataLimite
    )

    if ([string]::IsNullOrWhiteSpace([string]$UltimoAcesso)) {
        return "Inativo"
    }

    if ([string]$UltimoAcesso -eq "Sem registro") {
        return "Inativo"
    }

    $dataUltimoAcesso = [datetime]::MinValue
    $culturaBrasil = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")

    if ([datetime]::TryParse([string]$UltimoAcesso, $culturaBrasil, [System.Globalization.DateTimeStyles]::None, [ref]$dataUltimoAcesso)) {
        if ($dataUltimoAcesso.Date -ge $DataLimite.Date) {
            return "Ativo"
        }
    }

    return "Inativo"
}

function Get-UltimoAcessoUsuario {
    param(
        [object]$Usuario,
        [string]$StartDate,
        [string]$EndDate
    )

    try {
        $auditoria = Get-PWUserAuditTrailRecords `
            -Users $Usuario `
            -StartDate $StartDate `
            -EndDate $EndDate `
            -WarningAction SilentlyContinue `
            -ErrorAction Stop

        $ultimoLogin = @($auditoria |
            Where-Object { $_.Action -eq "User Login" } |
            Sort-Object ActionDate -Descending |
            Select-Object -First 1)

        if (-not $ultimoLogin -or $ultimoLogin.Count -eq 0) {
            return [PSCustomObject]@{
                UltimoAcesso = "Sem registro"
                Consulta     = "OK"
                Mensagem     = ""
            }
        }

        $actionDate = $ultimoLogin[0].ActionDate

        if ($actionDate -is [datetime]) {
            $ultimoAcesso = $actionDate.ToString("dd/MM/yyyy")
        }
        else {
            $dataUltimoAcesso = [datetime]::MinValue
            if ([datetime]::TryParse([string]$actionDate, [ref]$dataUltimoAcesso)) {
                $ultimoAcesso = $dataUltimoAcesso.ToString("dd/MM/yyyy")
            }
            else {
                $ultimoAcesso = [string]$actionDate
            }
        }

        return [PSCustomObject]@{
            UltimoAcesso = $ultimoAcesso
            Consulta     = "OK"
            Mensagem     = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            UltimoAcesso = "Erro ao consultar Audit Trail"
            Consulta     = "Erro"
            Mensagem     = $_.Exception.Message
        }
    }
}

function Exportar-Xlsx {
    param(
        [object[]]$Dados,
        [string]$Caminho
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $intervaloCabecalho = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Usuarios"

        $colunas = @(
            "Nome",
            "Email",
            "ID",
            "Ultimo acesso",
            "Status",
            "Status acesso",
            "Status ProjectWise",
            "Elegivel exclusao",
            "Motivo",
            "Acao executada",
            "Resultado"
        )

        for ($coluna = 0; $coluna -lt $colunas.Count; $coluna++) {
            $worksheet.Cells.Item(1, $coluna + 1).Value2 = $colunas[$coluna]
        }

        for ($linha = 0; $linha -lt $Dados.Count; $linha++) {
            for ($coluna = 0; $coluna -lt $colunas.Count; $coluna++) {
                $valor = $Dados[$linha].$($colunas[$coluna])
                $worksheet.Cells.Item($linha + 2, $coluna + 1).Value2 = [string]$valor
            }
        }

        $intervaloCabecalho = $worksheet.Range(
            $worksheet.Cells.Item(1, 1),
            $worksheet.Cells.Item(1, $colunas.Count)
        )
        $intervaloCabecalho.Font.Bold = $true

        $worksheet.Columns.AutoFit() | Out-Null
        $workbook.SaveAs($Caminho, 51)
    }
    catch {
        throw "Nao foi possivel gerar o arquivo XLSX. Verifique se o Microsoft Excel esta instalado. Detalhes: $($_.Exception.Message)"
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }

        if ($intervaloCabecalho) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($intervaloCabecalho) | Out-Null
        }

        if ($worksheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        }

        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        if ($excel) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Atualizar-Grade {
    $script:GridUsuarios.DataSource = $null
    if ($script:Resultados.Count -gt 0) {
        $script:GridUsuarios.DataSource = [System.Collections.ArrayList]@($script:Resultados)
        $script:GridUsuarios.AutoResizeColumns()
    }

    $script:LabelResumo.Text = "Usuarios: $($script:Resultados.Count) | Inativos elegiveis: $($script:InativosElegiveis.Count)"
    $temResultados = ($script:Resultados.Count -gt 0)
    $temElegiveis = ($script:InativosElegiveis.Count -gt 0)

    $script:BtnExportar.Enabled = $temResultados
    $script:BtnDesativar.Enabled = $temElegiveis
    $script:BtnExcluir.Enabled = ($temElegiveis -and $script:CheckPermitirExclusao.Checked)
}

function Conectar-ProjectWise {
    try {
        Write-UiLog "Carregando modulo PWPS_DAB..."
        Import-Module PWPS_DAB -ErrorAction Stop

        Write-UiLog "Abrindo login do ProjectWise..."
        New-PWLogin | Out-Null

        $script:UsuarioAtual = Get-PWCurrentUser -ErrorAction SilentlyContinue
        $nomeAtual = Get-ValorSeguroPropriedade -Objeto $script:UsuarioAtual -PossiveisNomes @("Name", "UserName", "LoginName")

        $script:ConectadoPW = $true
        $script:BtnConsultar.Enabled = $true
        $script:BtnConectar.Enabled = $false
        $script:LabelConexao.Text = "Conectado: $nomeAtual"
        Write-UiLog "Conectado ao ProjectWise como '$nomeAtual'."
    }
    catch {
        Show-UiError "Nao foi possivel conectar ao ProjectWise.`n`n$($_.Exception.Message)"
        Write-UiLog "Erro de conexao: $($_.Exception.Message)"
    }
}

function Consultar-UsuariosInativos {
    if (-not $script:ConectadoPW) {
        Show-UiError "Conecte ao ProjectWise antes de consultar."
        return
    }

    $dias = [int]$script:NumericDias.Value
    $script:Resultados = @()
    $script:InativosElegiveis = @()
    Atualizar-Grade

    try {
        $script:BtnConsultar.Enabled = $false
        $script:BtnDesativar.Enabled = $false
        $script:BtnExcluir.Enabled = $false
        $script:BtnExportar.Enabled = $false

        $dataLimiteStatus = (Get-Date).Date.AddDays(-$dias)
        $startDate = $dataLimiteStatus.ToString("yyyy-MM-dd")
        $endDate = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")

        Write-UiLog "Consulta iniciada. Criterio: sem login nos ultimos $dias dias ou usuario desabilitado no ProjectWise."
        Write-UiLog "Periodo consultado: $startDate ate $endDate."

        $idUsuarioAtual = Get-ValorSeguroPropriedade -Objeto $script:UsuarioAtual -PossiveisNomes @("ID", "Id", "UserID", "UserId")
        $nomeUsuarioAtual = Get-ValorSeguroPropriedade -Objeto $script:UsuarioAtual -PossiveisNomes @("Name", "UserName", "LoginName")

        $usuarios = @(Get-PWUsersByMatch)
        $usuariosDesabilitados = @(Get-PWUsersByMatch -Disabled -ErrorAction SilentlyContinue)
        $idsDesabilitados = @{}
        $nomesDesabilitados = @{}

        foreach ($usuarioDesabilitado in $usuariosDesabilitados) {
            $idDesabilitado = Get-ValorSeguroPropriedade -Objeto $usuarioDesabilitado -PossiveisNomes @("ID", "Id", "UserID", "UserId")
            $nomeDesabilitado = Get-ValorSeguroPropriedade -Objeto $usuarioDesabilitado -PossiveisNomes @("Name", "UserName", "LoginName")

            if (-not [string]::IsNullOrWhiteSpace($idDesabilitado)) {
                $idsDesabilitados[$idDesabilitado] = $true
            }

            if (-not [string]::IsNullOrWhiteSpace($nomeDesabilitado)) {
                $nomesDesabilitados[$nomeDesabilitado.ToLowerInvariant()] = $true
            }
        }

        $totalUsuarios = $usuarios.Count
        if ($totalUsuarios -eq 0) {
            Show-UiInfo "Nenhum usuario encontrado."
            return
        }

        $listaResultado = New-Object System.Collections.Generic.List[object]

        for ($indice = 0; $indice -lt $usuarios.Count; $indice++) {
            $usuario = $usuarios[$indice]
            $contador = $indice + 1

            $nome = Get-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("Name", "UserName", "LoginName")
            $email = Get-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("Email", "EmailAddress")
            $id = Get-ValorSeguroPropriedade -Objeto $usuario -PossiveisNomes @("ID", "Id", "UserID", "UserId")

            $script:ProgressConsulta.Value = [Math]::Min(100, [int](($contador / $totalUsuarios) * 100))
            $script:LabelResumo.Text = "Processando $contador de $totalUsuarios - $nome"
            if (($contador % 10) -eq 0 -or $contador -eq 1 -or $contador -eq $totalUsuarios) {
                Write-UiLog "Processando $contador de $totalUsuarios - $nome"
            }
            [System.Windows.Forms.Application]::DoEvents()

            $consultaAcesso = Get-UltimoAcessoUsuario -Usuario $usuario -StartDate $startDate -EndDate $endDate
            $statusAcesso = Get-StatusUsuario -UltimoAcesso $consultaAcesso.UltimoAcesso -DataLimite $dataLimiteStatus

            $estaDesabilitado = $false
            if ($id -ne "" -and $idsDesabilitados.ContainsKey($id)) {
                $estaDesabilitado = $true
            }
            if ($nome -ne "" -and $nomesDesabilitados.ContainsKey($nome.ToLowerInvariant())) {
                $estaDesabilitado = $true
            }

            $statusProjectWise = if ($estaDesabilitado) { "Inativo/Desabilitado" } else { "Ativo/Habilitado" }
            $statusFinal = if ($statusAcesso -eq "Inativo" -or $estaDesabilitado) { "Inativo" } else { "Ativo" }

            $elegivelExclusao = "Nao"
            $motivo = ""

            if ($statusFinal -eq "Inativo" -and $consultaAcesso.Consulta -eq "OK") {
                $elegivelExclusao = "Sim"
                if ($statusAcesso -eq "Inativo" -and $estaDesabilitado) {
                    $motivo = "Usuario sem login nos ultimos $dias dias e ja inativo/desabilitado no ProjectWise"
                }
                elseif ($statusAcesso -eq "Inativo") {
                    $motivo = "Usuario sem login nos ultimos $dias dias"
                }
                else {
                    $motivo = "Usuario ja inativo/desabilitado no ProjectWise"
                }
            }

            if ($consultaAcesso.Consulta -eq "Erro") {
                $motivo = "Nao elegivel: erro ao consultar Audit Trail - $($consultaAcesso.Mensagem)"
            }

            if (($id -ne "" -and $idUsuarioAtual -ne "" -and $id -eq $idUsuarioAtual) -or ($nome -ne "" -and $nomeUsuarioAtual -ne "" -and $nome -ieq $nomeUsuarioAtual)) {
                $elegivelExclusao = "Nao"
                $motivo = "Nao elegivel: usuario conectado na sessao atual"
            }

            $listaResultado.Add([PSCustomObject]@{
                Nome                 = $nome
                Email                = $email
                ID                   = $id
                "Ultimo acesso"      = $consultaAcesso.UltimoAcesso
                Status               = $statusFinal
                "Status acesso"      = $statusAcesso
                "Status ProjectWise" = $statusProjectWise
                "Elegivel exclusao"  = $elegivelExclusao
                Motivo               = $motivo
                "Acao executada"     = ""
                Resultado            = ""
            })
        }

        $script:Resultados = @($listaResultado | Sort-Object Nome)
        $script:InativosElegiveis = @($script:Resultados | Where-Object { $_.Status -eq "Inativo" -and $_."Elegivel exclusao" -eq "Sim" })
        $script:ProgressConsulta.Value = 100
        Atualizar-Grade
        Write-UiLog "Consulta concluida. Usuarios: $($script:Resultados.Count). Elegiveis: $($script:InativosElegiveis.Count)."
    }
    catch {
        Show-UiError "Erro durante a consulta.`n`n$($_.Exception.Message)"
        Write-UiLog "Erro durante a consulta: $($_.Exception.Message)"
    }
    finally {
        $script:BtnConsultar.Enabled = $script:ConectadoPW
        Atualizar-Grade
    }
}

function Exportar-Relatorio {
    if ($script:Resultados.Count -eq 0) {
        Show-UiError "Consulte os usuarios antes de exportar."
        return
    }

    $caminho = $script:TextCaminhoSaida.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($caminho)) {
        Show-UiError "Informe o caminho do arquivo Excel."
        return
    }

    try {
        Exportar-Xlsx -Dados $script:Resultados -Caminho $caminho
        Write-UiLog "Relatorio exportado: $caminho"
        Show-UiInfo "Relatorio exportado com sucesso."
    }
    catch {
        Show-UiError $_.Exception.Message
        Write-UiLog "Erro ao exportar: $($_.Exception.Message)"
    }
}

function Executar-AcaoUsuarios {
    param(
        [ValidateSet("Desativar", "Excluir")]
        [string]$Acao
    )

    if ($script:InativosElegiveis.Count -eq 0) {
        Show-UiError "Nao ha usuarios inativos elegiveis para executar a acao."
        return
    }

    if ($Acao -eq "Excluir" -and -not $script:CheckPermitirExclusao.Checked) {
        Show-UiError "Marque a trava 'Permitir exclusao definitiva' antes de excluir usuarios."
        return
    }

    $textoConfirmacao = if ($Acao -eq "Excluir") { "EXCLUIR" } else { "DESATIVAR" }
    $mensagem = if ($Acao -eq "Excluir") {
        "ATENCAO: esta acao vai excluir definitivamente $($script:InativosElegiveis.Count) usuarios elegiveis."
    }
    else {
        "Esta acao vai desativar $($script:InativosElegiveis.Count) usuarios elegiveis."
    }

    if (-not (Confirmar-TextoDialog -TextoEsperado $textoConfirmacao -Mensagem $mensagem)) {
        Show-UiInfo "Acao cancelada."
        Write-UiLog "$Acao cancelado pelo usuario."
        return
    }

    if (-not (Ask-UiYesNo -Mensagem "Confirmar $Acao de $($script:InativosElegiveis.Count) usuarios?" -Titulo "Ultima confirmacao")) {
        Show-UiInfo "Acao cancelada."
        Write-UiLog "$Acao cancelado na ultima confirmacao."
        return
    }

    try {
        $script:BtnDesativar.Enabled = $false
        $script:BtnExcluir.Enabled = $false
        $script:BtnConsultar.Enabled = $false

        $total = $script:InativosElegiveis.Count
        for ($indice = 0; $indice -lt $script:InativosElegiveis.Count; $indice++) {
            $item = $script:InativosElegiveis[$indice]
            $contador = $indice + 1

            $script:ProgressConsulta.Value = [Math]::Min(100, [int](($contador / $total) * 100))
            $script:LabelResumo.Text = "$Acao $contador de $total - $($item.Nome)"
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $usuario = Get-PWUsersByMatch -UserId ([int]$item.ID) -ErrorAction Stop

                if ($Acao -eq "Desativar") {
                    Set-PWUserDisabled -InputUser @($usuario) -ErrorAction Stop | Out-Null
                    $item."Acao executada" = "Desativado"
                    Write-UiLog "Desativado: $($item.Nome)"
                }
                else {
                    Remove-PWUserByMatch -InputUsers @($usuario) -ErrorAction Stop | Out-Null
                    $item."Acao executada" = "Excluido"
                    Write-UiLog "Excluido: $($item.Nome)"
                }

                $item.Resultado = "Sucesso"
            }
            catch {
                $item."Acao executada" = $Acao
                $item.Resultado = "Erro: $($_.Exception.Message)"
                Write-UiLog "Erro em $($item.Nome): $($_.Exception.Message)"
            }
        }

        Atualizar-Grade
        Exportar-Relatorio
        Show-UiInfo "$Acao concluido. Consulte a coluna Resultado no relatorio."
    }
    finally {
        $script:BtnConsultar.Enabled = $script:ConectadoPW
        Atualizar-Grade
    }
}

# ---------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Usuarios Inativos ProjectWise"
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.MinimumSize = New-Object System.Drawing.Size(1040, 760)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true

$fontePadrao = New-Object System.Drawing.Font("Segoe UI", 10)

$labelTitulo = New-Object System.Windows.Forms.Label
$labelTitulo.Text = "Usuarios Inativos ProjectWise"
$labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$labelTitulo.AutoSize = $true
$labelTitulo.Location = New-Object System.Drawing.Point(20, 15)
$form.Controls.Add($labelTitulo)

$btnFechar = New-Object System.Windows.Forms.Button
$btnFechar.Text = "Fechar"
$btnFechar.Location = New-Object System.Drawing.Point(1055, 18)
$btnFechar.Size = New-Object System.Drawing.Size(90, 28)
$btnFechar.Anchor = "Top,Right"
$btnFechar.Add_Click({ $form.Close() })
$form.Controls.Add($btnFechar)

$groupConfig = New-Object System.Windows.Forms.GroupBox
$groupConfig.Text = "Etapa 1 - Criterios e seguranca"
$groupConfig.Font = $fontePadrao
$groupConfig.Location = New-Object System.Drawing.Point(20, 55)
$groupConfig.Size = New-Object System.Drawing.Size(1125, 135)
$groupConfig.Anchor = "Top,Left,Right"
$form.Controls.Add($groupConfig)

$script:BtnConectar = New-Object System.Windows.Forms.Button
$script:BtnConectar.Text = "Conectar ProjectWise"
$script:BtnConectar.Location = New-Object System.Drawing.Point(15, 28)
$script:BtnConectar.Size = New-Object System.Drawing.Size(165, 32)
$script:BtnConectar.Add_Click({ Conectar-ProjectWise })
$groupConfig.Controls.Add($script:BtnConectar)

$script:LabelConexao = New-Object System.Windows.Forms.Label
$script:LabelConexao.Text = "Conexao: nao conectado"
$script:LabelConexao.Location = New-Object System.Drawing.Point(195, 35)
$script:LabelConexao.Size = New-Object System.Drawing.Size(500, 22)
$groupConfig.Controls.Add($script:LabelConexao)

$labelDias = New-Object System.Windows.Forms.Label
$labelDias.Text = "Dias sem acesso:"
$labelDias.Location = New-Object System.Drawing.Point(15, 78)
$labelDias.Size = New-Object System.Drawing.Size(115, 22)
$groupConfig.Controls.Add($labelDias)

$script:NumericDias = New-Object System.Windows.Forms.NumericUpDown
$script:NumericDias.Location = New-Object System.Drawing.Point(130, 75)
$script:NumericDias.Size = New-Object System.Drawing.Size(80, 25)
$script:NumericDias.Minimum = 1
$script:NumericDias.Maximum = 3650
$script:NumericDias.Value = 180
$groupConfig.Controls.Add($script:NumericDias)

$script:CheckPermitirExclusao = New-Object System.Windows.Forms.CheckBox
$script:CheckPermitirExclusao.Text = "Permitir exclusao definitiva"
$script:CheckPermitirExclusao.Location = New-Object System.Drawing.Point(245, 76)
$script:CheckPermitirExclusao.Size = New-Object System.Drawing.Size(230, 24)
$script:CheckPermitirExclusao.Add_CheckedChanged({ Atualizar-Grade })
$groupConfig.Controls.Add($script:CheckPermitirExclusao)

$labelCriterios = New-Object System.Windows.Forms.Label
$labelCriterios.Text = "Criterios mesclados: sem login no periodo informado ou usuario ja inativo/desabilitado no ProjectWise."
$labelCriterios.Location = New-Object System.Drawing.Point(500, 78)
$labelCriterios.Size = New-Object System.Drawing.Size(595, 22)
$labelCriterios.Anchor = "Top,Left,Right"
$groupConfig.Controls.Add($labelCriterios)

$groupSaida = New-Object System.Windows.Forms.GroupBox
$groupSaida.Text = "Etapa 2 - Arquivo de saida"
$groupSaida.Font = $fontePadrao
$groupSaida.Location = New-Object System.Drawing.Point(20, 200)
$groupSaida.Size = New-Object System.Drawing.Size(1125, 80)
$groupSaida.Anchor = "Top,Left,Right"
$form.Controls.Add($groupSaida)

$labelSaida = New-Object System.Windows.Forms.Label
$labelSaida.Text = "Excel:"
$labelSaida.Location = New-Object System.Drawing.Point(15, 34)
$labelSaida.Size = New-Object System.Drawing.Size(50, 22)
$groupSaida.Controls.Add($labelSaida)

$script:TextCaminhoSaida = New-Object System.Windows.Forms.TextBox
$script:TextCaminhoSaida.Location = New-Object System.Drawing.Point(65, 31)
$script:TextCaminhoSaida.Size = New-Object System.Drawing.Size(860, 25)
$script:TextCaminhoSaida.Anchor = "Top,Left,Right"
$script:TextCaminhoSaida.Text = Join-Path ([Environment]::GetFolderPath("Desktop")) "usuarios_pw_inativos_180_dias.xlsx"
$groupSaida.Controls.Add($script:TextCaminhoSaida)

$script:BtnSelecionarSaida = New-Object System.Windows.Forms.Button
$script:BtnSelecionarSaida.Text = "Selecionar"
$script:BtnSelecionarSaida.Location = New-Object System.Drawing.Point(940, 29)
$script:BtnSelecionarSaida.Size = New-Object System.Drawing.Size(100, 28)
$script:BtnSelecionarSaida.Anchor = "Top,Right"
$script:BtnSelecionarSaida.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Salvar relatorio de usuarios inativos ProjectWise"
    $dialog.Filter = "Arquivo Excel (*.xlsx)|*.xlsx"
    $dialog.FileName = "usuarios_pw_inativos_$($script:NumericDias.Value)_dias.xlsx"
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:TextCaminhoSaida.Text = $dialog.FileName
    }
})
$groupSaida.Controls.Add($script:BtnSelecionarSaida)

$script:BtnExportar = New-Object System.Windows.Forms.Button
$script:BtnExportar.Text = "Exportar"
$script:BtnExportar.Location = New-Object System.Drawing.Point(1048, 29)
$script:BtnExportar.Size = New-Object System.Drawing.Size(70, 28)
$script:BtnExportar.Anchor = "Top,Right"
$script:BtnExportar.Enabled = $false
$script:BtnExportar.Add_Click({ Exportar-Relatorio })
$groupSaida.Controls.Add($script:BtnExportar)

$groupUsuarios = New-Object System.Windows.Forms.GroupBox
$groupUsuarios.Text = "Etapa 3 - Consulta e usuarios encontrados"
$groupUsuarios.Font = $fontePadrao
$groupUsuarios.Location = New-Object System.Drawing.Point(20, 290)
$groupUsuarios.Size = New-Object System.Drawing.Size(1125, 325)
$groupUsuarios.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($groupUsuarios)

$script:BtnConsultar = New-Object System.Windows.Forms.Button
$script:BtnConsultar.Text = "Consultar usuarios"
$script:BtnConsultar.Location = New-Object System.Drawing.Point(15, 28)
$script:BtnConsultar.Size = New-Object System.Drawing.Size(145, 30)
$script:BtnConsultar.Enabled = $false
$script:BtnConsultar.Add_Click({ Consultar-UsuariosInativos })
$groupUsuarios.Controls.Add($script:BtnConsultar)

$script:LabelResumo = New-Object System.Windows.Forms.Label
$script:LabelResumo.Text = "Usuarios: 0 | Inativos elegiveis: 0"
$script:LabelResumo.Location = New-Object System.Drawing.Point(175, 34)
$script:LabelResumo.Size = New-Object System.Drawing.Size(620, 22)
$groupUsuarios.Controls.Add($script:LabelResumo)

$script:ProgressConsulta = New-Object System.Windows.Forms.ProgressBar
$script:ProgressConsulta.Location = New-Object System.Drawing.Point(815, 32)
$script:ProgressConsulta.Size = New-Object System.Drawing.Size(290, 22)
$script:ProgressConsulta.Anchor = "Top,Right"
$groupUsuarios.Controls.Add($script:ProgressConsulta)

$script:GridUsuarios = New-Object System.Windows.Forms.DataGridView
$script:GridUsuarios.Location = New-Object System.Drawing.Point(15, 70)
$script:GridUsuarios.Size = New-Object System.Drawing.Size(1090, 240)
$script:GridUsuarios.Anchor = "Top,Bottom,Left,Right"
$script:GridUsuarios.ReadOnly = $true
$script:GridUsuarios.AllowUserToAddRows = $false
$script:GridUsuarios.AllowUserToDeleteRows = $false
$script:GridUsuarios.SelectionMode = "FullRowSelect"
$script:GridUsuarios.AutoSizeColumnsMode = "DisplayedCells"
$script:GridUsuarios.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupUsuarios.Controls.Add($script:GridUsuarios)

$groupAcao = New-Object System.Windows.Forms.GroupBox
$groupAcao.Text = "Etapa 4 - Acao apos confirmacao"
$groupAcao.Font = $fontePadrao
$groupAcao.Location = New-Object System.Drawing.Point(20, 625)
$groupAcao.Size = New-Object System.Drawing.Size(1125, 75)
$groupAcao.Anchor = "Bottom,Left,Right"
$form.Controls.Add($groupAcao)

$script:BtnDesativar = New-Object System.Windows.Forms.Button
$script:BtnDesativar.Text = "Desativar elegiveis"
$script:BtnDesativar.Location = New-Object System.Drawing.Point(15, 28)
$script:BtnDesativar.Size = New-Object System.Drawing.Size(155, 32)
$script:BtnDesativar.Enabled = $false
$script:BtnDesativar.Add_Click({ Executar-AcaoUsuarios -Acao "Desativar" })
$groupAcao.Controls.Add($script:BtnDesativar)

$script:BtnExcluir = New-Object System.Windows.Forms.Button
$script:BtnExcluir.Text = "Excluir elegiveis"
$script:BtnExcluir.Location = New-Object System.Drawing.Point(185, 28)
$script:BtnExcluir.Size = New-Object System.Drawing.Size(140, 32)
$script:BtnExcluir.Enabled = $false
$script:BtnExcluir.Add_Click({ Executar-AcaoUsuarios -Acao "Excluir" })
$groupAcao.Controls.Add($script:BtnExcluir)

$labelTrava = New-Object System.Windows.Forms.Label
$labelTrava.Text = "A exclusao exige a trava marcada, digitacao de EXCLUIR e uma ultima confirmacao."
$labelTrava.Location = New-Object System.Drawing.Point(345, 34)
$labelTrava.Size = New-Object System.Drawing.Size(745, 22)
$labelTrava.Anchor = "Top,Left,Right"
$groupAcao.Controls.Add($labelTrava)

$groupLog = New-Object System.Windows.Forms.GroupBox
$groupLog.Text = "Acompanhamento"
$groupLog.Font = $fontePadrao
$groupLog.Location = New-Object System.Drawing.Point(20, 710)
$groupLog.Size = New-Object System.Drawing.Size(1125, 60)
$groupLog.Anchor = "Bottom,Left,Right"
$form.Controls.Add($groupLog)

$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Multiline = $true
$script:LogBox.ReadOnly = $true
$script:LogBox.ScrollBars = "Vertical"
$script:LogBox.Location = New-Object System.Drawing.Point(15, 23)
$script:LogBox.Size = New-Object System.Drawing.Size(1090, 25)
$script:LogBox.Anchor = "Top,Bottom,Left,Right"
$script:LogBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$groupLog.Controls.Add($script:LogBox)

$form.Add_Shown({
    Write-UiLog "Interface carregada. Conecte ao ProjectWise para iniciar."
})

[void]$form.ShowDialog()
