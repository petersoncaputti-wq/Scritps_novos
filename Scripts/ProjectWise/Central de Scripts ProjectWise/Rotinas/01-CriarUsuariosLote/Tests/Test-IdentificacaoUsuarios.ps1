$ErrorActionPreference = 'Stop'
$arquivo = Join-Path $PSScriptRoot '../CriarUsuariosLote.Integrado.ps1'
$tokens = $null
$erros = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $arquivo), [ref]$tokens, [ref]$erros)
if ($erros.Count) { throw ($erros | Out-String) }
foreach ($nome in @('Get-ValorSeguroPropriedade', 'Test-UsuarioExistePW', 'Resolver-NomeUsuarioPW', 'Buscar-UsuarioPW', 'Criar-UsuarioPW')) {
    $funcao = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $nome }, $true)
    Invoke-Expression $funcao.Extent.Text
}
function Get-PWUsersByMatch {
    [CmdletBinding()]
    param($UserName, $Email)
    if ($script:falhaConsulta) { throw 'Consulta indisponivel' }
    # Retorna candidatos amplos para testar o filtro exato da rotina.
    $script:usuarios
}
function Assert-Erro($registro, $trecho) {
    try { $null = Test-UsuarioExistePW $registro }
    catch { if ($_.Exception.Message -like "*$trecho*") { return }; throw }
    throw "Era esperado erro: $trecho"
}
$registro = [pscustomobject]@{ UserName = 'Joao'; Email = 'novo@empresa.com' }
$script:usuarios = @([pscustomobject]@{ Name = 'Joao'; Email = 'outro@empresa.com' })
if (-not (Test-UsuarioExistePW $registro).ConflitoNome) { throw 'Conflito nao detectado' }
$script:usuarios = @([pscustomobject]@{ Name = 'Joao'; Email = '' })
if (-not (Test-UsuarioExistePW $registro).ConflitoNome) { throw 'Conflito nao detectado' }
$script:usuarios = @([pscustomobject]@{ Name = 'OutroLogin'; Email = ' NOVO@EMPRESA.COM ' })
$r = Test-UsuarioExistePW $registro
if (-not $r.Existe -or $r.EncontradoPor -ne 'E-mail' -or $r.Usuario.Name -ne 'OutroLogin') { throw 'Falha ao identificar por email' }
$script:usuarios = @($script:usuarios[0], $script:usuarios[0])
Assert-Erro $registro 'Mais de um cadastro'
$script:usuarios = @([pscustomobject]@{ Name = 'Joao Silva'; Email = 'prefixo.novo@empresa.com' })
if ((Test-UsuarioExistePW $registro).Existe) { throw 'Correspondencia parcial aceita' }
$script:usuarios = @()
if ((Test-UsuarioExistePW $registro).Existe) { throw 'Usuario novo considerado existente' }
$script:falhaConsulta = $true
Assert-Erro $registro 'Consulta indisponivel'
$script:falhaConsulta = $false
function Write-Status { param($Mensagem, $Nivel) }
function Solicitar-NomeAlternativo {
    param($Registro, $EmailConflitante)
    $script:pedidos++
    $valor = $script:respostas.Dequeue()
    return $valor
}
$script:usuarios = @([pscustomobject]@{ Name = 'Joao'; Email = 'outro@empresa.com' })
$script:respostas = New-Object System.Collections.Queue
foreach ($nome in @('', 'Joao', 'Joao.Novo')) { $script:respostas.Enqueue($nome) }
$script:pedidos = 0
$r = Resolver-NomeUsuarioPW $registro
if ($r.Existe -or $r.ConflitoNome -or $registro.UserName -ne 'Joao.Novo' -or $script:pedidos -ne 3) { throw 'Falha ao resolver nome alternativo' }
$registro.UserName = 'Joao'
$script:respostas.Enqueue($null)
try { $null = Resolver-NomeUsuarioPW $registro; throw 'Cancelamento ignorado' }
catch { if ($_.Exception.Message -notlike '*cancelado pelo operador*') { throw } }
$script:usuarios = @([pscustomobject]@{ Name = 'LoginExistente'; Email = 'novo@empresa.com' })
$script:pedidos = 0
$r = Resolver-NomeUsuarioPW $registro
if (-not $r.Existe -or $script:pedidos -ne 0) { throw 'Cadastro pelo email pediu nome desnecessariamente' }
function New-PWUserSimple {
    [CmdletBinding()]
    param($UserNames, $Email, $IMSUser, $Description)
    $script:criado = [pscustomobject]@{ Name = $UserNames[0]; Email = $Email }
    $script:usuarios += $script:criado
}
function Set-IdentidadeBentleyIMS {
    param($Usuario, $Email)
    $script:identidade = [pscustomobject]@{ Usuario = $Usuario; Email = $Email }
}
$registro.UserName = 'Joao'
$script:usuarios = @([pscustomobject]@{ Name = 'Joao'; Email = 'outro@empresa.com' })
$original = $script:usuarios[0]
$script:respostas.Enqueue('Joao.Novo')
$null = Resolver-NomeUsuarioPW $registro
$novo = Criar-UsuarioPW $registro
if ($novo.Name -ne 'Joao.Novo' -or $novo.Email -ne 'novo@empresa.com') { throw 'Cadastro criado incorretamente' }
if ($script:identidade.Usuario -ne $novo -or $script:identidade.Email -ne $registro.Email) { throw 'IMS vinculado ao cadastro errado' }
if ($original.Name -ne 'Joao' -or $original.Email -ne 'outro@empresa.com') { throw 'Cadastro original alterado' }
Write-Output 'OK: 11 cenarios, incluindo criacao e identidade IMS; sintaxe validada. Nenhuma conexao ao ProjectWise.'
