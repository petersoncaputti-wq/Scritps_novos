param([switch]$Executar, [string]$Arquivo, [switch]$SemPausa, [switch]$SomentePrevia)
$ErrorActionPreference = 'Stop'
$codigo = 0
try {
    $python = Get-Command python -ErrorAction Stop
    Write-Host "Python: $($python.Source)"
    $argumentos = @('-u', (Join-Path $PSScriptRoot 'criar_usuarios.py'))
    if ($Arquivo) { $argumentos += @('--arquivo', $Arquivo) }
    if ($SomentePrevia) { $argumentos += '--somente-previa' }
    & $python.Source @argumentos
    $codigo = $LASTEXITCODE
    if ($codigo -ne 0) {
        Write-Host 'O script encontrou uma falha. Consulte a mensagem acima e o diagnostico na pasta Logs.' -ForegroundColor Red
    }
}
catch {
    $codigo = 1
    Write-Host "Falha ao iniciar: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if (-not $SemPausa) {
        Read-Host 'Pressione ENTER para fechar esta janela' | Out-Null
    }
}
exit $codigo
