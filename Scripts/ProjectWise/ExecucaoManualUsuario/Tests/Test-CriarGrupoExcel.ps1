$ErrorActionPreference = 'Stop'
$arquivo = Join-Path $PSScriptRoot '../17 - Criar Grupo e Incluir Participantes por Excel.ps1'
$tokens = $null; $erros = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $arquivo), [ref]$tokens, [ref]$erros)
if ($erros.Count) { throw ($erros | Out-String) }
$ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { Invoke-Expression $_.Extent.Text }
function Get-PWGroups { [CmdletBinding()] param($GroupName) if ($script:grupoCriado) { [pscustomobject]@{ Name = $GroupName } } }
function New-PWGroupByName { [CmdletBinding()] param($GroupName,$GroupDescription) $script:grupoCriado=$true; $script:criacao++ }
function Get-PWUsersByMatch { [CmdletBinding()] param($Email) $script:usuarios }
function Get-PWUsersInGroup { [CmdletBinding()] param($GroupName) $script:membros }
function Add-PWUserToGroup { [CmdletBinding()] param($InputUser,$GroupName) $script:membros += $InputUser; $script:inclusao++ }
function Assert($condicao, $mensagem) { if (-not $condicao) { throw $mensagem } }
$script:usuarios = @([pscustomobject]@{ Name='Joao'; Email='pessoa@empresa.com'; UserID=10 })
$script:membros=@(); $script:grupoCriado=$false; $script:criacao=0; $script:inclusao=0
$linhas=@([pscustomobject]@{ Email=' PESSOA@EMPRESA.COM ' },[pscustomobject]@{ Email='pessoa@empresa.com' },[pscustomobject]@{ Email='ausente@empresa.com' },[pscustomobject]@{ Email='invalido' })
$r=@(Processar-ParticipantesGrupo $linhas 'ARTESP' -Validar)
Assert ($r[0].Status -eq 'VALIDO' -and $r[1].Status -eq 'DUPLICADO_PLANILHA' -and $r[2].Status -eq 'ERRO' -and $r[3].Status -eq 'ERRO') 'Validacao incorreta'
Assert ($script:criacao -eq 0 -and $script:inclusao -eq 0) 'Validacao alterou dados'
$r=@(Processar-ParticipantesGrupo $linhas 'ARTESP')
Assert ($r[0].Status -eq 'ADICIONADO' -and $script:criacao -eq 1 -and $script:inclusao -eq 1) 'Inclusao incorreta'
$r=@(Processar-ParticipantesGrupo @($linhas[0]) 'ARTESP')
Assert ($r[0].Status -eq 'JA_MEMBRO' -and $script:inclusao -eq 1) 'Membro foi incluido novamente'
$script:usuarios += [pscustomobject]@{Name='OutraPessoa'; Email='pessoa@empresa.com';UserID=11}
$r=@(Processar-ParticipantesGrupo @($linhas[0]) 'ARTESP')
Assert ($r[0].Status -eq 'ERRO' -and $script:inclusao -eq 1) 'Email ambiguo aceito'
Assert (-not (Test-MembroGrupo $script:usuarios[0] @([pscustomobject]@{Name='Joao';UserID=99}))) 'IDs diferentes tratados como mesma pessoa'
function Get-PWUsersByMatch { [CmdletBinding()] param($Email) throw 'Falha de consulta' }
$r=@(Processar-ParticipantesGrupo @($linhas[0]) 'ARTESP')
Assert ($r[0].Status -eq 'ERRO' -and $script:inclusao -eq 1) 'Falha de consulta alterou dados'
Write-Output 'OK: sintaxe e cenarios locais de validacao, criacao, inclusao, duplicidade, ambiguidade e falha de consulta.'
