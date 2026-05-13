<#
.SYNOPSIS
    Testa o fluxo completo de selecionar concessao/projeto via ProjectWise e diagnosticar dados PW Web/PWDM.

.DESCRIPTION
    Este script e somente leitura e orquestra dois scripts V2:
    1. pw_listar_concessoes_projetos_v2.ps1
    2. diagnostico_pw_projectwise_project_v2.ps1

    Ele lista as concessoes/projetos diretamente do ProjectWise, permite escolher por numero
    e roda o diagnostico no projeto selecionado.

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\teste_fluxo_selecionar_projeto_pw_v2.ps1"

.EXAMPLE
    powershell.exe -NoProfile -MTA -ExecutionPolicy Bypass -File ".\teste_fluxo_selecionar_projeto_pw_v2.ps1" -ConcessaoNome "Ecovias"
#>

[CmdletBinding()]
param(
    [string]$ConcessaoNome = "",
    [int]$ProfundidadeProjetos = 1
)

$ErrorActionPreference = "Stop"

$PastaLogs = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $PastaLogs)) {
    New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
}

function Assert-Mta {
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartment -ne [System.Threading.ApartmentState]::MTA) {
        throw "Execute este teste usando powershell.exe -MTA. Estado atual: $apartment"
    }
}

function Ler-Numero {
    param(
        [string]$Mensagem,
        [int]$Minimo,
        [int]$Maximo
    )

    while ($true) {
        $entrada = Read-Host $Mensagem
        if ($entrada -match '^\d+$') {
            $numero = [int]$entrada
            if ($numero -ge $Minimo -and $numero -le $Maximo) {
                return $numero
            }
        }

        Write-Host "Informe um numero entre $Minimo e $Maximo." -ForegroundColor Yellow
    }
}

function Mostrar-Concessoes {
    param([array]$Concessoes)

    Write-Host ""
    Write-Host "Concessoes encontradas:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Concessoes.Count; $i++) {
        $item = $Concessoes[$i]
        $qtd = @($item.projetos).Count
        $status = if ($item.pastaProjetosEncontrada) { "OK" } else { "sem pasta Projetos" }
        "{0:00}. {1} ({2} projeto(s)) [{3}]" -f ($i + 1), $item.nome, $qtd, $status | Write-Host
    }
}

function Mostrar-Projetos {
    param([array]$Projetos)

    Write-Host ""
    Write-Host "Projetos encontrados:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Projetos.Count; $i++) {
        $item = $Projetos[$i]
        "{0:000}. {1} | ID: {2} | GUID: {3}" -f ($i + 1), $item.nome, $item.id, $item.guid | Write-Host
        if (-not [string]::IsNullOrWhiteSpace($item.caminho)) {
            Write-Host "     $($item.caminho)" -ForegroundColor DarkGray
        }
    }
}

Assert-Mta

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$arquivoLista = Join-Path $PastaLogs "pw_fluxo_lista_concessoes_projetos_$timestamp.json"

Write-Host "Etapa 1/3 - Listando concessoes/projetos pelo ProjectWise..." -ForegroundColor Cyan
$parametrosLista = @{
    Saida = $arquivoLista
    ProfundidadeProjetos = $ProfundidadeProjetos
}
if (-not [string]::IsNullOrWhiteSpace($ConcessaoNome)) {
    $parametrosLista["ConcessaoNome"] = $ConcessaoNome
}

& (Join-Path $PSScriptRoot "pw_listar_concessoes_projetos_v2.ps1") @parametrosLista

if (-not (Test-Path -LiteralPath $arquivoLista)) {
    throw "Arquivo de lista nao foi gerado: $arquivoLista"
}

$dados = Get-Content -Raw -Path $arquivoLista | ConvertFrom-Json
$concessoes = @($dados.concessoes | Where-Object { $_ -ne $null })
if ($concessoes.Count -eq 0) {
    throw "Nenhuma concessao foi retornada pelo ProjectWise."
}

Write-Host ""
Write-Host "Etapa 2/3 - Selecionando concessao e projeto..." -ForegroundColor Cyan
Mostrar-Concessoes -Concessoes $concessoes
$indiceConcessao = Ler-Numero -Mensagem "Numero da concessao" -Minimo 1 -Maximo $concessoes.Count
$concessaoSelecionada = $concessoes[$indiceConcessao - 1]

$projetos = @($concessaoSelecionada.projetos | Where-Object { $_ -ne $null })
if ($projetos.Count -eq 0) {
    throw "A concessao selecionada nao possui projetos no JSON gerado."
}

Mostrar-Projetos -Projetos $projetos
$indiceProjeto = Ler-Numero -Mensagem "Numero do projeto" -Minimo 1 -Maximo $projetos.Count
$projetoSelecionado = $projetos[$indiceProjeto - 1]

Write-Host ""
Write-Host "Selecionado:" -ForegroundColor Cyan
Write-Host "Concessao: $($concessaoSelecionada.nome)"
Write-Host "Projeto  : $($projetoSelecionado.nome)"
Write-Host "ID       : $($projetoSelecionado.id)"
Write-Host "GUID     : $($projetoSelecionado.guid)"
Write-Host "Caminho  : $($projetoSelecionado.caminho)"

Write-Host ""
Write-Host "Etapa 3/3 - Rodando diagnostico do ProjectWise Project..." -ForegroundColor Cyan

if (-not [string]::IsNullOrWhiteSpace($projetoSelecionado.id)) {
    & (Join-Path $PSScriptRoot "diagnostico_pw_projectwise_project_v2.ps1") -FolderId $projetoSelecionado.id
}
elseif (-not [string]::IsNullOrWhiteSpace($projetoSelecionado.caminho)) {
    & (Join-Path $PSScriptRoot "diagnostico_pw_projectwise_project_v2.ps1") -FolderPath $projetoSelecionado.caminho
}
else {
    throw "Projeto selecionado nao possui ID nem caminho para diagnostico."
}

Write-Host ""
Write-Host "[OK] Fluxo de teste concluido." -ForegroundColor Green
