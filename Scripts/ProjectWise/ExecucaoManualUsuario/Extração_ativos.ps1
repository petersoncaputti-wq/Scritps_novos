Clear-Host
Import-Module PWPS_DAB -ErrorAction Stop
Add-Type -AssemblyName System.Windows.Forms

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

function Exportar-Xlsx {
    param(
        [object[]]$Dados,
        [string]$Caminho
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $intervaloDados = $null
    $intervaloCabecalho = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Usuários"

        $colunas = @("Nome", "Email", "ID", "Último acesso", "Status")

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
        throw "Não foi possível gerar o arquivo XLSX. Verifique se o Microsoft Excel está instalado. Detalhes: $($_.Exception.Message)"
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

        if ($intervaloDados) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($intervaloDados) | Out-Null
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

Write-Host "=== EXTRAÇÃO DE USUÁRIOS PW - ÚLTIMO ACESSO 6 MESES ===" -ForegroundColor Cyan

New-PWLogin

$dialog = New-Object System.Windows.Forms.SaveFileDialog
$dialog.Title = "Salvar relatório de usuários ProjectWise"
$dialog.Filter = "Arquivo Excel (*.xlsx)|*.xlsx"
$dialog.FileName = "usuarios_pw_ultimo_acesso_6_meses.xlsx"

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Operação cancelada." -ForegroundColor Yellow
    exit
}

$caminhoSaida = $dialog.FileName

# Últimos 6 meses
$startDate = (Get-Date).AddMonths(-6).ToString("yyyy-MM-dd")
$endDate = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")
$dataLimiteStatus = (Get-Date).Date.AddDays(-180)

Write-Host "Período consultado: $startDate até $endDate" -ForegroundColor Yellow

$usuarios = Get-PWUsersByMatch
$totalUsuarios = $usuarios.Count
$resultadoFinal = New-Object System.Collections.Generic.List[object]

$contador = 0

foreach ($usuario in $usuarios) {
    $contador++

    Write-Progress `
        -Activity "Consultando usuários ProjectWise" `
        -Status "Processando $contador de $totalUsuarios - $($usuario.Name)" `
        -PercentComplete (($contador / $totalUsuarios) * 100)

    Write-Host "[$contador/$totalUsuarios] $($usuario.Name)" -ForegroundColor DarkGray

    $ultimoAcesso = "Sem registro"

    try {
        $auditoria = Get-PWUserAuditTrailRecords `
            -Users $usuario `
            -StartDate $startDate `
            -EndDate $endDate `
            -WarningAction SilentlyContinue `
            -ErrorAction Stop

        if ($auditoria -and $auditoria.Rows.Count -gt 0) {
            $ultimoLogin = $auditoria |
                Where-Object { $_.Action -eq "User Login" } |
                Sort-Object ActionDate -Descending |
                Select-Object -First 1

            if ($ultimoLogin) {
                if ($ultimoLogin.ActionDate -is [datetime]) {
                    $ultimoAcesso = $ultimoLogin.ActionDate.ToString("dd/MM/yyyy")
                }
                else {
                    $dataUltimoAcesso = [datetime]::MinValue
                    if ([datetime]::TryParse([string]$ultimoLogin.ActionDate, [ref]$dataUltimoAcesso)) {
                        $ultimoAcesso = $dataUltimoAcesso.ToString("dd/MM/yyyy")
                    }
                    else {
                        $ultimoAcesso = $ultimoLogin.ActionDate
                    }
                }
            }
        }
    }
    catch {
        $ultimoAcesso = "Erro ao consultar Audit Trail"
    }

    $status = Get-StatusUsuario -UltimoAcesso $ultimoAcesso -DataLimite $dataLimiteStatus

    $resultadoFinal.Add([PSCustomObject]@{
        Nome            = $usuario.Name
        Email           = $usuario.Email
        ID              = $usuario.ID
        "Último acesso" = $ultimoAcesso
        Status          = $status
    })
}

Write-Progress -Activity "Consultando usuários ProjectWise" -Completed

$resultadoOrdenado = @($resultadoFinal | Sort-Object Nome)
Exportar-Xlsx -Dados $resultadoOrdenado -Caminho $caminhoSaida

Write-Host ""
Write-Host "Arquivo gerado com sucesso em:" -ForegroundColor Green
Write-Host $caminhoSaida

Write-Host ""
Write-Host "Total de usuários exportados: $($resultadoFinal.Count)" -ForegroundColor Cyan
