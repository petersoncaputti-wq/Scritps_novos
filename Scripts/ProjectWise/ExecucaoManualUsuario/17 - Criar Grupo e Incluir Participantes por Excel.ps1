<#
.SYNOPSIS
Cria/reutiliza um grupo ProjectWise e inclui usuarios existentes por e-mail.
.EXAMPLE
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1'
.EXAMPLE
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1' -NomeGrupo 'ARTESP' -CaminhoExcel 'C:\Dados\Participantes.xlsx' -SomenteValidar
.EXAMPLE
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1' -GerarModelo -CaminhoExcel 'C:\Dados\Modelo_Grupo.xlsx'
#>
[CmdletBinding()]
param(
    [string]$NomeGrupo,
    [string]$DescricaoGrupo,
    [string]$CaminhoExcel,
    [string]$NomeAba,
    [string]$PastaRelatorios = (Join-Path $PSScriptRoot 'Logs'),
    [switch]$SomenteValidar,
    [switch]$GerarModelo
)
$ErrorActionPreference = 'Stop'

function Get-CampoParticipante {
    param($Objeto, [string[]]$Nomes)
    foreach ($nome in $Nomes) {
        $prop = $Objeto.PSObject.Properties[$nome]
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return ([string]$prop.Value).Trim()
        }
    }
    return ''
}

function Test-MembroGrupo {
    param($Usuario, [object[]]$Membros)
    $id = Get-CampoParticipante $Usuario @('UserID', 'ID')
    $nome = Get-CampoParticipante $Usuario @('Name', 'UserName', 'LoginName')
    foreach ($membro in $Membros) {
        $idMembro = Get-CampoParticipante $membro @('UserID', 'ID')
        if ($id -and $idMembro) {
            if ($id -eq $idMembro) { return $true }
            continue
        }
        $nomeMembro = Get-CampoParticipante $membro @('Name', 'UserName', 'LoginName')
        if ($nome -and $nome -ieq $nomeMembro) { return $true }
    }
    return $false
}

function Importar-Participantes {
    param([string]$Path, [string]$Aba)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Planilha nao encontrada: $Path" }
    if ([IO.Path]::GetExtension($Path) -notin @('.xlsx', '.xlsm')) { throw 'Use uma planilha .xlsx ou .xlsm.' }
    $parametros = @{ Path = $Path; ErrorAction = 'Stop' }
    if ($Aba) { $parametros.WorksheetName = $Aba }
    $linhas = @(Import-Excel @parametros)
    if (-not $linhas.Count) { throw 'A planilha nao possui participantes.' }
    $cabecalhos = @($linhas[0].PSObject.Properties.Name)
    if (-not @($cabecalhos | Where-Object { $_ -in @('Email', 'E-mail', 'Mail') }).Count) {
        throw 'A primeira linha da planilha deve conter a coluna Email ou E-mail.'
    }
    return $linhas
}

function Processar-ParticipantesGrupo {
    param([object[]]$Linhas, [string]$Grupo, [switch]$Validar)
    $grupos = @(Get-PWGroups -GroupName $Grupo -ErrorAction Stop | Where-Object { $_.Name -ieq $Grupo })
    if ($grupos.Count -gt 1) { throw "Mais de um grupo corresponde ao nome '$Grupo'." }
    $membros = @()
    if ($grupos.Count) { $membros = @(Get-PWUsersInGroup -GroupName $Grupo -ErrorAction Stop) }
    $grupoDisponivel = $grupos.Count -eq 1
    $vistos = @{}
    $linha = 1
    foreach ($item in $Linhas) {
        $linha++
        $email = Get-CampoParticipante $item @('Email', 'E-mail', 'Mail')
        $resultado = [pscustomobject]@{
            Linha = $linha; Grupo = $Grupo; Email = $email; Usuario = ''; Status = ''; Mensagem = ''
        }
        try {
            if ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'E-mail vazio ou em formato invalido.' }
            if ($vistos.ContainsKey($email)) {
                $resultado.Status = 'DUPLICADO_PLANILHA'
                $resultado.Mensagem = 'E-mail repetido na planilha; consulte o resultado da primeira ocorrencia.'
            }
            else {
                $vistos[$email] = $true
                $usuarios = @(Get-PWUsersByMatch -Email $email -ErrorAction Stop | Where-Object {
                    (Get-CampoParticipante $_ @('Email', 'EMail', 'Mail')) -ieq $email
                })
                if ($usuarios.Count -eq 0) { throw 'Usuario nao encontrado pelo e-mail. Cadastre-o primeiro pela rotina de usuarios.' }
                if ($usuarios.Count -gt 1) { throw 'E-mail encontrado em mais de um cadastro. Revise os usuarios antes de incluir.' }
                $usuario = $usuarios[0]
                $resultado.Usuario = Get-CampoParticipante $usuario @('Name', 'UserName', 'LoginName')
                if (Test-MembroGrupo $usuario $membros) {
                    $resultado.Status = 'JA_MEMBRO'
                    $resultado.Mensagem = 'Usuario ja pertence ao grupo.'
                }
                elseif ($Validar) {
                    $resultado.Status = 'VALIDO'
                    $resultado.Mensagem = 'Usuario localizado; pode ser incluido. Nenhuma alteracao realizada.'
                }
                else {
                    if (-not $grupoDisponivel) {
                        $descricao = $DescricaoGrupo
                        if (-not $descricao) { $descricao = $Grupo }
                        New-PWGroupByName -GroupName $Grupo -GroupDescription $descricao -ErrorAction Stop | Out-Null
                        $confirmacao = @(Get-PWGroups -GroupName $Grupo -ErrorAction Stop | Where-Object { $_.Name -ieq $Grupo })
                        if ($confirmacao.Count -ne 1) { throw 'Nao foi possivel confirmar a criacao do grupo.' }
                        $grupoDisponivel = $true
                    }
                    Add-PWUserToGroup -InputUser @($usuario) -GroupName $Grupo -ErrorAction Stop | Out-Null
                    $membros = @(Get-PWUsersInGroup -GroupName $Grupo -ErrorAction Stop)
                    if (-not (Test-MembroGrupo $usuario $membros)) { throw 'Inclusao executada, mas o vinculo nao foi confirmado. Verifique o grupo no ProjectWise.' }
                    $resultado.Status = 'ADICIONADO'
                    $resultado.Mensagem = 'Inclusao no grupo confirmada.'
                }
            }
        }
        catch {
            $resultado.Status = 'ERRO'
            $resultado.Mensagem = $_.Exception.Message
        }
        Write-Host ("Linha {0}: {1} | {2} | {3}" -f $linha, $email, $resultado.Status, $resultado.Mensagem)
        $resultado
    }
}

$loginIniciado = $false
try {
    Import-Module ImportExcel -ErrorAction Stop
    if ($GerarModelo) {
        if (-not $CaminhoExcel) { $CaminhoExcel = Join-Path $PSScriptRoot 'Modelo_Participantes_Grupo.xlsx' }
        if (Test-Path -LiteralPath $CaminhoExcel) { throw 'O arquivo de destino ja existe. Escolha outro caminho para o modelo.' }
        @([pscustomobject]@{ Email = 'pessoa@empresa.com'; Nome = 'Nome da pessoa (opcional)' }) |
            Export-Excel -Path $CaminhoExcel -WorksheetName 'Participantes' -AutoSize -FreezeTopRow
        Write-Host "Modelo criado: $CaminhoExcel"
        return
    }
    Add-Type -AssemblyName System.Windows.Forms
    if (-not $NomeGrupo) {
        Add-Type -AssemblyName Microsoft.VisualBasic
        $NomeGrupo = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o nome do grupo (ex.: SIGLA-PROJETO-ARTESP):', 'Criar grupo ProjectWise', '')
    }
    if ([string]::IsNullOrWhiteSpace($NomeGrupo)) { Write-Host 'Operacao cancelada.'; return }
    $NomeGrupo = $NomeGrupo.Trim()
    if (-not $CaminhoExcel) {
        $dialogo = New-Object System.Windows.Forms.OpenFileDialog
        $dialogo.Title = 'Selecione a planilha de participantes'
        $dialogo.Filter = 'Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm'
        try {
            if ($dialogo.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $CaminhoExcel = $dialogo.FileName
        }
        finally { $dialogo.Dispose() }
    }
    $linhas = @(Importar-Participantes -Path $CaminhoExcel -Aba $NomeAba)
    New-Item -ItemType Directory -Path $PastaRelatorios -Force | Out-Null
    $relatorio = Join-Path $PastaRelatorios ('PW_Grupo_Participantes_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
    Import-Module PWPS_DAB -ErrorAction Stop
    foreach ($comando in @('New-PWLogin', 'Undo-PWLogin', 'Get-PWGroups', 'Get-PWUsersByMatch', 'Get-PWUsersInGroup', 'New-PWGroupByName', 'Add-PWUserToGroup')) {
        if (-not (Get-Command $comando -ErrorAction SilentlyContinue)) { throw "Comando indisponivel: $comando" }
    }
    $login = New-PWLogin -UseGui -ErrorAction Stop
    if (-not $login) { throw 'Login no ProjectWise cancelado ou nao confirmado.' }
    $loginIniciado = $true
    $resultados = @(Processar-ParticipantesGrupo -Linhas $linhas -Grupo $NomeGrupo -Validar:$SomenteValidar)
    $resultados | Export-Csv -LiteralPath $relatorio -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "`nResumo do grupo: $NomeGrupo"
    $resultados | Group-Object Status | ForEach-Object { Write-Host ("{0}: {1}" -f $_.Name, $_.Count) }
    Write-Host "Relatorio: $relatorio"
    if ($SomenteValidar) { Write-Host 'Modo de validacao: nenhuma alteracao realizada.' }
    Write-Host 'Este script configura membros do grupo. As permissoes em arquivos/pastas devem ser configuradas separadamente.'
}
finally {
    if ($loginIniciado) { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null }
}
