[CmdletBinding()]
param(
 [string]$Datasource = '01SSRV305.ECSC.ECORODOVIAS.CORP:Ecorodovias-01',
 [string]$Grupo = 'NP-01.03.01-consulta ultima rev',
 [string]$Projeto = '01.03.01'
)
$ErrorActionPreference='Stop'
Import-Module PWPS_DAB -WarningAction SilentlyContinue
$pasta = Join-Path $PSScriptRoot ('Logs\DiagnosticoGrupo_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $pasta -Force | Out-Null
$login=$false
try {
 if (-not (New-PWLogin -DatasourceName $Datasource -UseGui -ErrorAction Stop)) { throw 'Login nao confirmado.' }
 $login=$true
 $grupos=@(Get-PWGroups -GroupName $Grupo -ErrorAction Stop | Where-Object {$_.Name -ieq $Grupo})
 if ($grupos.Count -ne 1) { throw "Esperado um grupo exato; encontrados: $($grupos.Count)." }
 $grupoReal=$grupos[0].Name
 $grupos | Export-Clixml (Join-Path $pasta 'Grupo.xml')
 @(Get-PWUsersInGroup -GroupName $grupoReal -ErrorAction Stop) | Select-Object Name,UserID,Email,Description | Export-Csv (Join-Path $pasta 'Participantes.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
 $fila=New-Object System.Collections.Queue
 $fila.Enqueue([pscustomobject]@{Nome=$grupoReal;Tipo='group'})
 $visitados=@{}; $vinculos=New-Object System.Collections.Generic.List[object]
 while ($fila.Count) {
  $item=$fila.Dequeue(); $chave="$($item.Tipo):$($item.Nome)"
  if ($visitados.ContainsKey($chave)) { continue }; $visitados[$chave]=$true
  $retorno=Get-PWUserListsByMember -MemberName $item.Nome -MemberType $item.Tipo -ErrorAction Stop
  $listas=if ($retorno -is [System.Collections.IDictionary]) { @($retorno.Values) } else { @($retorno) }
  foreach ($lista in $listas) {
   $lista | Export-Clixml (Join-Path $pasta 'Ultimo_Retorno_Listas.xml')
   $nome=if ($lista -is [string]) {$lista} else { foreach ($campo in @('Name','UserListName','ListName','userlistname','o_name','Value')) { if ($lista.$campo) { [string]$lista.$campo; break } } }
   if (-not $nome) { $lista | Format-List * | Out-Host; throw 'Consulta de listas retornou objeto sem nome reconhecido.' }
   $vinculos.Add([pscustomobject]@{Membro=$item.Nome;Tipo=$item.Tipo;Lista=$nome})
   $fila.Enqueue([pscustomobject]@{Nome=$nome;Tipo='userlist'})
  }
 }
 $vinculos | Export-Csv (Join-Path $pasta 'Vinculos_Listas.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
 $pastas=@(Get-PWFolders -FolderName $Projeto -PopulatePaths -ErrorAction Stop | Where-Object {$_.Name -ieq $Projeto})
 $pastas = @($pastas | ForEach-Object { Get-PWFolders -FolderID $_.ProjectID -JustOne -PopulatePaths -ErrorAction Stop })
 $pastas | Select-Object Name,ProjectID,FullPath | Export-Csv (Join-Path $pasta 'Projetos_Localizados.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
 $erros=New-Object System.Collections.Generic.List[string]
 foreach ($raiz in $pastas) {
  $caminhoRaiz = $raiz.GetFullPath()
  Write-Output "PROJETO: $caminhoRaiz"
  if ([string]::IsNullOrWhiteSpace($caminhoRaiz)) { throw 'Caminho do projeto vazio; consulta de descendentes interrompida.' }
  $alvos=@($raiz)+@(Get-PWFolders -FolderPath $caminhoRaiz -PopulatePaths -ErrorAction Stop)
  $vistos=@{}
  foreach ($alvo in $alvos) {
   if ($vistos.ContainsKey([string]$alvo.ProjectID)) { continue }; $vistos[[string]$alvo.ProjectID]=$true
   try {
    $seguranca=Get-PWFolderSecurity -InputFolder @($alvo) -ErrorAction Stop
    $seguranca | Export-Clixml (Join-Path $pasta ("Seguranca_{0}.xml" -f $alvo.ProjectID))
   } catch { $erros.Add("$($alvo.FullPath): $($_.Exception.Message)") }
  }
 }
 $erros | Set-Content (Join-Path $pasta 'Falhas_Consulta.txt') -Encoding UTF8
 @('Consulta somente leitura.',"Datasource: $Datasource","Grupo: $grupoReal",'Seguranca coletada apenas nas pastas localizadas para o projeto informado; nao e uma varredura completa do datasource.','Permissoes individuais de documentos e acessos efetivos por usuario exigem analise adicional.') | Set-Content (Join-Path $pasta 'Escopo.txt') -Encoding UTF8
 Write-Output "RELATORIO: $pasta"
 Write-Output "Listas/vinculos: $($vinculos.Count); projetos localizados: $($pastas.Count); falhas de seguranca: $($erros.Count)"
} finally { if ($login) { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } }
