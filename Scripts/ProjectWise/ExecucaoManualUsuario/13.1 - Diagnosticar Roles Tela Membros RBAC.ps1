$ErrorActionPreference = "Stop"

$PastaScript = Join-Path $PSScriptRoot "PWDM_Gerenciamento_Participantes_V2"
$ScriptPython = Join-Path $PastaScript "diagnostico_rbac_roles_membros_v2.py"

if (-not (Test-Path -LiteralPath $ScriptPython)) {
    throw "Script Python nao encontrado: $ScriptPython"
}

Push-Location $PastaScript
try {
    python $ScriptPython
}
finally {
    Pop-Location
}
