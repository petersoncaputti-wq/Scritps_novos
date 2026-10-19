# Planejamento, leitura e aplicacao explicita com preservacao de participantes.
function Obter-BloqueioArquivoARTESP {
    param([string]$NomeArquivo)
    $nome=[IO.Path]::GetFileName($NomeArquivo.Trim())
    if($nome -match '^PL'){return 'Nome do arquivo comeca com PL'}
    $base=[IO.Path]::GetFileNameWithoutExtension($nome)
    if($base -match '-(?<Numero>\d+)-R[^-]+$'){
        $numero=$Matches.Numero.TrimStart('0')
        if($numero.Length -gt 3 -or ($numero.Length -eq 3 -and [int]$numero -ge 900)){
            return 'Numero antes da revisao igual ou maior que 900'
        }
    }
    return ''
}
function Confirmar-ArquivoPermitidoARTESP {
    param($Documento)
    $nome=[string]$Documento.FileName
    if([string]::IsNullOrWhiteSpace($nome)){$nome=[string]$Documento.Name}
    if([string]::IsNullOrWhiteSpace($nome)){$nome=[string]$Documento.DocumentName}
    if([string]::IsNullOrWhiteSpace($nome)){throw 'Documento sem nome para validar as regras de exclusao.'}
    $bloqueio=Obter-BloqueioArquivoARTESP $nome
    if($bloqueio){throw "Arquivo bloqueado: $nome. $bloqueio. Nenhuma permissao deve ser aplicada."}
}
function Novo-ContextoACLARTESP {
    @{Pastas=@{};Origens=@{};Membros=@{};Mascaras=@{};GrupoID=0;Tempos=@{};UltimoResultado=''}
}
function Medir-OperacaoACL {
    param([string]$Nome,$Contexto,[scriptblock]$Acao)
    $relogio=[Diagnostics.Stopwatch]::StartNew()
    try { & $Acao }
    finally {
        $relogio.Stop()
        if ($Contexto) {
            if (-not $Contexto.Tempos.ContainsKey($Nome)) { $Contexto.Tempos[$Nome]=@{Quantidade=0;Milissegundos=0.0} }
            $Contexto.Tempos[$Nome].Quantidade++
            $Contexto.Tempos[$Nome].Milissegundos += $relogio.Elapsed.TotalMilliseconds
        }
    }
}
function Mesclar-ACLARTESP {
    param([object[]]$ACL, [int]$GrupoID)
    if ($GrupoID -le 0) { throw 'ID do grupo ARTESP invalido.' }
    $mapa = @{}
    foreach ($entrada in $ACL) {
        if ($null -eq $entrada -or $null -eq $entrada.MemberType -or
            $null -eq $entrada.MemberId -or $null -eq $entrada.AccessMask) { throw 'ACL incompleta.' }
        $chave = '{0}:{1}' -f [int]$entrada.MemberType,[int]$entrada.MemberId
        if ($mapa.ContainsKey($chave)) { throw "Participante duplicado na ACL: $chave" }
        $mapa[$chave] = [pscustomobject]@{
            MemberType = [int]$entrada.MemberType; MemberId = [int]$entrada.MemberId
            AccessMask = [uint32]$entrada.AccessMask
        }
    }
    $artesp = '2:' + $GrupoID
    if ($mapa.ContainsKey($artesp)) {
        if ($mapa[$artesp].AccessMask -ne 20) { throw 'ARTESP ja possui direitos diferentes; nao substituir automaticamente.' }
    } else {
        $mapa[$artesp] = [pscustomobject]@{MemberType=2;MemberId=$GrupoID;AccessMask=[uint32]20}
    }
    @($mapa.Values | Sort-Object MemberType,MemberId)
}

function Confirmar-ACLCompleta {
    param([object[]]$Esperada, [object[]]$Atual)
    $e = @($Esperada | ForEach-Object { '{0}:{1}:{2}' -f $_.MemberType,$_.MemberId,$_.AccessMask } | Sort-Object)
    $a = @($Atual | ForEach-Object { '{0}:{1}:{2}' -f $_.MemberType,$_.MemberId,$_.AccessMask } | Sort-Object)
    if ($e.Count -eq 0 -and $a.Count -eq 0) { return $true }
    if ($e.Count -ne $a.Count -or @(Compare-Object $e $a).Count -gt 0) {
        throw 'A ACL completa diverge: participante removido/adicionado ou direito alterado. Interromper o lote.'
    }
    return $true
}

function Ler-ACLWorkflowPropria {
    param($Documento)
    $w = [PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWWrapper', $true)
    $b = [Reflection.BindingFlags]'Public,NonPublic,Static'
    $select = @($w.GetMethods($b) | Where-Object { $_.Name -eq 'aaApi_SelectAccessControlItems' -and $_.GetParameters()[0].ParameterType -eq [int] })[0]
    $numeric = @($w.GetMethods($b) | Where-Object { $_.Name -eq 'aaApi_GetAccessControlItemNumericProperty' -and $_.GetParameters()[0].ParameterType -eq [int] })[0]
    $pt = $w.GetNestedType('AccessObjectProperty',[Reflection.BindingFlags]'Public,NonPublic')
    $n = [int]$select.Invoke($null,[object[]]@(4,[int]$Documento.ProjectID,[int]$Documento.DocumentID,[int]$Documento.WorkflowId,[int]$Documento.StateId,-1,-1))
    if ($n -lt 0) { throw 'Erro ao enumerar a ACL completa do documento.' }
    for ($i=0;$i -lt $n;$i++) {
        $r=[ordered]@{}
        foreach($p in @('ObjType','ObjID1','ObjID2','Workflow','State','MemberType','MemberId','AccessMask')) {
            $r[$p]=[int]$numeric.Invoke($null,[object[]]@([int][Enum]::Parse($pt,$p),$i))
        }
        if ($r.ObjType -ne 4 -or $r.ObjID1 -ne $Documento.ProjectID -or $r.ObjID2 -ne $Documento.DocumentID -or
            $r.Workflow -ne $Documento.WorkflowId -or $r.State -ne $Documento.StateId) { throw 'ACL fora do documento/workflow/estado esperado.' }
        [pscustomobject]$r
    }
}

function Ler-ACLWorkflowOrigem {
    param($Documento,$Contexto,[switch]$Atualizar)
    $chave='{0}:{1}:{2}' -f $Documento.ProjectID,$Documento.WorkflowId,$Documento.StateId
    if ($Contexto -and -not $Atualizar -and $Contexto.Origens.ContainsKey($chave)) { return $Contexto.Origens[$chave] }
    $idPasta=[string]$Documento.ProjectID
    if ($Contexto -and -not $Atualizar -and $Contexto.Pastas.ContainsKey($idPasta)) {
        $dados=$Contexto.Pastas[$idPasta]
    } else {
    $pastas = @(Get-PWFolders -FolderID $Documento.ProjectID -JustOne -ErrorAction Stop | Where-Object { $null -ne $_ -and $_ -isnot [bool] })
    if ($pastas.Count -ne 1) { throw 'Pasta do documento nao localizada de forma unica.' }
    $dados = Medir-OperacaoACL 'SegurancaPasta' $Contexto { Get-PWFolderSecurity -InputFolder $pastas[0] -ErrorAction Stop }
    if ($dados -is [Data.DataTable]) { $dados = $dados.Rows }
    $dados=@($dados)
    if ($Contexto) { $Contexto.Pastas[$idPasta]=$dados }
    }
    $origem = @($dados | Where-Object { $_.SecurityType -eq 'Document' -and $_.Workflow -eq $Documento.Workflow -and $_.State -eq $Documento.WorkflowState })
    if ($origem.Count -eq 0) { throw 'Nao foi possivel identificar todas as regras herdadas do workflow/estado na pasta. Escrita bloqueada.' }
    $c = [PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWPS_DAB.CommonAccessControlMethods', $true)
    $m = $c.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static')
    $idMetodo = @($m | Where-Object Name -eq 'getMemberID')[0]
    $mascaraMetodo = @($m | Where-Object Name -eq 'ConvertMaskStringToInt')[0]
    $resultado=@(foreach ($r in $origem) {
        $tipo=0; $id=0
        switch ([string]$r.Type) {
            'User' { $tipo=1 }
            'Group' { $tipo=2 }
            'UserList' { $tipo=3 }
            'All_Users (*Everyone)' { $tipo=4 }
            default { throw "Tipo de participante desconhecido: $($r.Type)" }
        }
        if ($tipo -ne 4) {
            $chaveMembro='{0}:{1}' -f $tipo,$r.Name
            if($Contexto -and -not $Atualizar -and $Contexto.Membros.ContainsKey($chaveMembro)){
                $id=[int]$Contexto.Membros[$chaveMembro]
            } else {
                $id=[int](Medir-OperacaoACL 'ResolverParticipante' $Contexto { $idMetodo.Invoke($null,[object[]]@([string]$r.Name,$tipo)) })
                if($Contexto){$Contexto.Membros[$chaveMembro]=$id}
            }
            if ($id -le 0) { throw "Participante nao resolvido: $($r.Name)" }
        }
        $chaveMask=[string]$r.Access_Control_Settings
        if($Contexto -and $Contexto.Mascaras.ContainsKey($chaveMask)){
            $mask=[int]$Contexto.Mascaras[$chaveMask]
        } else {
            $mask=[int]$mascaraMetodo.Invoke($null,[object[]]@($chaveMask,'Document'))
            if($Contexto){$Contexto.Mascaras[$chaveMask]=$mask}
        }
        if ($mask -lt 0) { throw "Mascara invalida para $($r.Name)" }
        [pscustomobject]@{MemberType=$tipo;MemberId=$id;AccessMask=[uint32]$mask;Name=[string]$r.Name;Origem=[string]$r.InheritingFrom}
    })
    if ($Contexto) { $Contexto.Origens[$chave]=$resultado }
    $resultado
}

function Atualizar-ACLWorkflowARTESP {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GUID, [Parameter(Mandatory)][string]$Destino,
        [switch]$Aplicar, [switch]$RecuperarOrigem, [string]$PlanoAprovado,
        $Contexto,$DocumentoLido,$SelecaoEsperada,[switch]$Silencioso,[switch]$Reverter)
    if($Reverter){
        Reverter-ACLWorkflowARTESP -GUID $GUID -Destino $Destino -Aplicar:$Aplicar -PlanoAprovado $PlanoAprovado -Contexto $Contexto -DocumentoLido $DocumentoLido -SelecaoEsperada $SelecaoEsperada -Silencioso:$Silencioso
        return
    }
    if (-not $Contexto) { $Contexto=Novo-ContextoACLARTESP }
    $Contexto.UltimoResultado=''
    if ($DocumentoLido -and -not $Aplicar) { $docs=@($DocumentoLido) }
    else { $docs=@(Medir-OperacaoACL 'ConsultarDocumento' $Contexto { Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop | Where-Object { $null -ne $_ -and $_ -isnot [bool] } }) }
    if ($docs.Count -ne 1) { throw 'Documento nao localizado de forma unica.' }
    $doc=$docs[0]
    Confirmar-ArquivoPermitidoARTESP $doc
    if ($SelecaoEsperada -and ($doc.ProjectID -ne [int]$SelecaoEsperada.FolderID -or
        $doc.Workflow -ne $SelecaoEsperada.Workflow -or $doc.WorkflowState -ne $SelecaoEsperada.Estado)) { throw 'Documento mudou de pasta/workflow/estado desde a selecao.' }
    if ($doc.ProjectID -le 0 -or $doc.DocumentID -le 0 -or $doc.WorkflowId -le 0 -or $doc.StateId -le 0) { throw 'Identificadores incompletos.' }
    if ($RecuperarOrigem -and $GUID -ne '5fc0819d-db6f-4841-a979-75a387bb1d01') { throw 'Recuperacao restrita ao PDF testado.' }
    $w=[PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWWrapper',$true)
    $metodos=$w.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static')
    $gm=@($metodos | Where-Object Name -eq 'GetGroupsByName')[0]
    if ($Contexto.GrupoID -le 0) {
        $grupos=Medir-OperacaoACL 'LocalizarARTESP' $Contexto { $gm.Invoke($null,@()) }
        $g=@($grupos.Values | Where-Object Name -eq 'ARTESP')
        if ($g.Count -ne 1) { throw 'Grupo ARTESP nao localizado.' }
        $Contexto.GrupoID=[int]$g[0].ID
    }
    $grupoID=[int]$Contexto.GrupoID
    $antes=@(Medir-OperacaoACL 'ACLPropria' $Contexto { Ler-ACLWorkflowPropria $doc })
    $base=@($antes); $herdadas=@()
    if ($antes.Count -eq 0 -or $RecuperarOrigem) {
        # Reutiliza origem no planejamento; consulta novamente antes de aplicar.
        $herdadas=@(Ler-ACLWorkflowOrigem $doc -Contexto $Contexto -Atualizar:$Aplicar)
        foreach($h in $herdadas) {
            $existente=@($base | Where-Object { $_.MemberType -eq $h.MemberType -and $_.MemberId -eq $h.MemberId })
            if ($existente.Count -eq 0) { $base += $h }
            elseif ($existente.Count -ne 1 -or $existente[0].AccessMask -ne $h.AccessMask) { throw 'Conflito entre ACL propria e origem. Revisao necessaria.' }
        }
    }
    $plano=@(Mesclar-ACLARTESP $base $grupoID)
    if ($PlanoAprovado) {
        $aprovado=Get-Content -LiteralPath $PlanoAprovado -Raw | ConvertFrom-Json
        if ($aprovado.GUID -ne $GUID -or $aprovado.WorkflowID -ne $doc.WorkflowId -or $aprovado.StateID -ne $doc.StateId) { throw 'Documento/workflow/estado diverge do plano aprovado.' }
        $null=Confirmar-ACLCompleta @($aprovado.ACLPropriaAntes) $antes
        $null=Confirmar-ACLCompleta @($aprovado.ACLEsperada) $plano
    }
    New-Item -ItemType Directory -Path $Destino -Force | Out-Null
    # Ja permissionado: valida a ACL inteira e a identidade, sem entrar no
    # caminho de escrita nem duplicar os arquivos de planejamento aprovados.
    if ($Aplicar -and $antes.Count -eq $plano.Count) {
        $null=Confirmar-ACLCompleta $plano $antes
        $confere=@(Medir-OperacaoACL 'ValidarACLJaCorreta' $Contexto { Ler-ACLWorkflowPropria $doc })
        $null=Confirmar-ACLCompleta $plano $confere
        $final=@(Medir-OperacaoACL 'ValidarDocumentoJaCorreto' $Contexto { Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop })
        if ($final.Count -ne 1 -or $final[0].ProjectID -ne $doc.ProjectID -or $final[0].DocumentID -ne $doc.DocumentID -or
            $final[0].WorkflowId -ne $doc.WorkflowId -or $final[0].StateId -ne $doc.StateId) { throw 'Documento mudou durante a verificacao de acesso existente.' }
        $Contexto.UltimoResultado='Sem alteracao - ARTESP e ACL completa conferidos'
        [pscustomobject]@{GUID=$GUID;Resultado=$Contexto.UltimoResultado;PlanoAprovado=$PlanoAprovado;Participantes=$confere.Count} |
            ConvertTo-Json | Set-Content (Join-Path $Destino 'ResultadoSemAlteracao.json') -Encoding UTF8
        return
    }
    $doc | Export-Clixml (Join-Path $Destino 'DocumentoAntes.xml') -Depth 5
    [pscustomobject]@{GUID=$GUID;WorkflowID=$doc.WorkflowId;StateID=$doc.StateId;ACLPropriaAntes=$antes;ACLOrigem=$herdadas;ACLEsperada=$plano} |
        ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Destino 'PlanoPreservacao.json') -Encoding UTF8
    if (-not $Silencioso) { $plano | Format-Table MemberType,MemberId,AccessMask -AutoSize | Out-Host }
    if (-not $Aplicar) { if (-not $Silencioso) { Write-Host 'Plano salvo; nenhuma alteracao aplicada.' }; return }
    $atual=@(Medir-OperacaoACL 'ValidarDocumentoAntes' $Contexto { Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop })
    if ($atual.Count -ne 1 -or $atual[0].ProjectID -ne $doc.ProjectID -or $atual[0].DocumentID -ne $doc.DocumentID -or
        $atual[0].WorkflowId -ne $doc.WorkflowId -or $atual[0].StateId -ne $doc.StateId) { throw 'Documento mudou desde o planejamento.' }
    Confirmar-ArquivoPermitidoARTESP $atual[0]
    $null=Confirmar-ACLCompleta $antes @(Medir-OperacaoACL 'ValidarACLAntes' $Contexto { Ler-ACLWorkflowPropria $doc })
    $assign=@($metodos | Where-Object Name -eq 'aaApi_AssignAccessList')[0]
    try {
        # Preserva/materializa todos os demais participantes antes de acrescentar ARTESP.
        foreach($r in @($plano | Sort-Object @{Expression={if($_.MemberType -eq 2 -and $_.MemberId -eq $grupoID){1}else{0}}})) {
            $existente=@($antes | Where-Object { $_.MemberType -eq $r.MemberType -and $_.MemberId -eq $r.MemberId })
            if ($existente.Count -gt 0) { continue }
            $ok=[bool](Medir-OperacaoACL 'EscreverPermissao' $Contexto { $assign.Invoke($null,[object[]]@(4,[int]$doc.ProjectID,[int]$doc.DocumentID,[int]$doc.WorkflowId,[int]$doc.StateId,[int]$r.MemberType,[int]$r.MemberId,[uint32]$r.AccessMask)) })
            if (-not $ok) { throw "Falha ao aplicar participante $($r.MemberType):$($r.MemberId). Resultado pode ser parcial; interrompido." }
        }
        $depois=@(Medir-OperacaoACL 'ValidarACLDepois' $Contexto { Ler-ACLWorkflowPropria $doc })
        $depois | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Destino 'ACLDepois.json') -Encoding UTF8
        $null=Confirmar-ACLCompleta $plano $depois
        $final=@(Medir-OperacaoACL 'ValidarDocumentoDepois' $Contexto { Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop })
        if ($final.Count -ne 1 -or $final[0].WorkflowId -ne $doc.WorkflowId -or $final[0].StateId -ne $doc.StateId) { throw 'Workflow/estado divergente apos aplicacao.' }
        'ACL completa confirmada; demais participantes e direitos preservados. Conferencia visual pendente.' |
            Set-Content (Join-Path $Destino 'Resultado.txt') -Encoding UTF8
        $Contexto.UltimoResultado='Concluido - ACL completa verificada'
    }
    catch { $_ | Out-String | Set-Content (Join-Path $Destino 'Erro.txt') -Encoding UTF8; throw }
}

function Remover-EntradaWorkflowARTESP {
    param($Documento,[int]$GrupoID)
    $w=[PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWWrapper',$true)
    $m=@($w.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static') | Where-Object {$_.Name -eq 'aaApi_RemoveAccessList' -and $_.GetParameters()[0].ParameterType -eq [int]})[0]
    if(-not $m){throw 'API de remocao indisponivel.'}
    [bool]$m.Invoke($null,[object[]]@(4,[int]$Documento.ProjectID,[int]$Documento.DocumentID,[int]$Documento.WorkflowId,[int]$Documento.StateId,2,$GrupoID))
}
function Reverter-ACLWorkflowARTESP {
    [CmdletBinding()]
    param([string]$GUID,[string]$Destino,[switch]$Aplicar,[string]$PlanoAprovado,
        $Contexto,$DocumentoLido,$SelecaoEsperada,[switch]$Silencioso)
    if(-not $Contexto){$Contexto=Novo-ContextoACLARTESP}
    $Contexto.UltimoResultado=''
    if($DocumentoLido -and -not $Aplicar){$docs=@($DocumentoLido)}
    else{$docs=@(Medir-OperacaoACL 'ConsultarDocumentoReversao' $Contexto {Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop})}
    if($docs.Count -ne 1){throw 'Documento nao localizado para reversao.'}
    $doc=$docs[0]
    $nome=[string]$doc.FileName
    if(-not $nome){$nome=[string]$doc.Name}
    $motivo=Obter-BloqueioArquivoARTESP $nome
    if(-not $motivo){throw 'Reversao bloqueada: documento nao atende as regras PL ou numero >= 900.'}
    if($doc.ProjectID -le 0 -or $doc.DocumentID -le 0 -or $doc.WorkflowId -le 0 -or $doc.StateId -le 0){throw 'IDs incompletos para reversao.'}
    if($SelecaoEsperada -and ($doc.ProjectID -ne [int]$SelecaoEsperada.FolderID -or $doc.Workflow -ne $SelecaoEsperada.Workflow -or $doc.WorkflowState -ne $SelecaoEsperada.Estado)){throw 'Documento mudou desde a selecao.'}
    if($Contexto.GrupoID -le 0){
        $w=[PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWWrapper',$true)
        $gm=@($w.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static')|Where-Object Name -eq 'GetGroupsByName')[0]
        $grupos=$gm.Invoke($null,@());$g=@($grupos.Values|Where-Object Name -eq 'ARTESP')
        if($g.Count -ne 1){throw 'Grupo ARTESP nao localizado.'}
        $Contexto.GrupoID=[int]$g[0].ID
    }
    $grupoID=[int]$Contexto.GrupoID
    $antes=@(Medir-OperacaoACL 'ACLAntesReversao' $Contexto {Ler-ACLWorkflowPropria $doc})
    $artesp=@($antes|Where-Object {$_.MemberType -eq 2 -and $_.MemberId -eq $grupoID})
    if($artesp.Count -gt 1){throw 'Entradas ARTESP duplicadas. Revisao necessaria.'}
    $esperada=@($antes|Where-Object {-not ($_.MemberType -eq 2 -and $_.MemberId -eq $grupoID)})
    $herdado=$false
    if($esperada.Count -eq 0){
        $origem=@(Ler-ACLWorkflowOrigem $doc -Contexto $Contexto -Atualizar:$Aplicar)
        $herdado=@($origem|Where-Object {$_.MemberType -eq 2 -and $_.MemberId -eq $grupoID -and $_.AccessMask -ne 0}).Count -gt 0
    }
    if($PlanoAprovado){
        $aprovado=Get-Content -LiteralPath $PlanoAprovado -Raw|ConvertFrom-Json
        if($aprovado.Operacao -ne 'Reverter' -or $aprovado.GUID -ne $GUID -or $aprovado.WorkflowID -ne $doc.WorkflowId -or $aprovado.StateID -ne $doc.StateId){throw 'Reversao diverge do plano aprovado.'}
        $null=Confirmar-ACLCompleta @($aprovado.ACLPropriaAntes) $antes
        $null=Confirmar-ACLCompleta @($aprovado.ACLEsperada) $esperada
        if([bool]$aprovado.AcessoHerdadoRestante -ne $herdado){throw 'Heranca mudou desde o plano de reversao.'}
    }
    New-Item -ItemType Directory -Path $Destino -Force|Out-Null
    [pscustomobject]@{Operacao='Reverter';GUID=$GUID;WorkflowID=$doc.WorkflowId;StateID=$doc.StateId;Motivo=$motivo;ACLPropriaAntes=$antes;ACLEsperada=$esperada;AcessoHerdadoRestante=$herdado}|
        ConvertTo-Json -Depth 8|Set-Content (Join-Path $Destino 'PlanoPreservacao.json') -Encoding UTF8
    if(-not $Aplicar){if(-not $Silencioso){Write-Host "Reversao planejada: $nome. Entrada propria ARTESP: $($artesp.Count). Heranca restante: $herdado"};return}
    try{
        # Uma segunda leitura protege os demais participantes contra mudancas recentes.
        $null=Confirmar-ACLCompleta $antes @(Ler-ACLWorkflowPropria $doc)
        $atual=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop)
        if($atual.Count -ne 1 -or $atual[0].ProjectID -ne $doc.ProjectID -or $atual[0].DocumentID -ne $doc.DocumentID -or
            $atual[0].WorkflowId -ne $doc.WorkflowId -or $atual[0].StateId -ne $doc.StateId){throw 'Documento mudou antes da reversao.'}
        $nomeAtual=[string]$atual[0].FileName;if(-not $nomeAtual){$nomeAtual=[string]$atual[0].Name}
        if(-not (Obter-BloqueioArquivoARTESP $nomeAtual)){throw 'Arquivo deixou de atender as regras; reversao bloqueada.'}
        if($artesp.Count -gt 0 -and -not $herdado){
            $ok=Medir-OperacaoACL 'RemoverPermissao' $Contexto {Remover-EntradaWorkflowARTESP $doc $grupoID}
            if(-not $ok){throw 'API nao confirmou remocao; resultado pode ser parcial.'}
        }
        $depois=@(Ler-ACLWorkflowPropria $doc)
        $alvoConferencia=@(if($herdado){$antes}else{$esperada})
        $null=Confirmar-ACLCompleta @($alvoConferencia) $depois
        $depois|ConvertTo-Json -Depth 6|Set-Content (Join-Path $Destino 'ACLDepois.json') -Encoding UTF8
        $final=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop)
        if($final.Count -ne 1 -or $final[0].ProjectID -ne $doc.ProjectID -or $final[0].DocumentID -ne $doc.DocumentID -or $final[0].WorkflowId -ne $doc.WorkflowId -or $final[0].StateId -ne $doc.StateId){throw 'Identidade/workflow/estado divergente apos reversao.'}
        $Contexto.UltimoResultado=if($herdado){'Nao revertido - ARTESP herdado; revisar origem'}elseif($artesp.Count -eq 0){'Sem alteracao - ARTESP ausente no workflow'}else{'Revertido - ARTESP removido; demais acessos preservados'}
        $Contexto.UltimoResultado|Set-Content (Join-Path $Destino 'Resultado.txt') -Encoding UTF8
    }catch{$_|Out-String|Set-Content (Join-Path $Destino 'Erro.txt') -Encoding UTF8;throw}
}
function Executar-LoteARTESP {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Linhas,[Parameter(Mandatory)][string]$Destino,[switch]$Reverter)
    if ($Linhas.Count -eq 0) { throw 'Nenhum documento selecionado.' }
    $guids=@{}
    $resultados=@($Linhas | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_.DocumentGUID)) { throw 'Documento selecionado sem GUID.' }
        if ($guids.ContainsKey($_.DocumentGUID)) { throw 'GUID duplicado na selecao.' }
        $guids[$_.DocumentGUID]=$true
        [pscustomobject]@{Documento=$_.Documento;GUID=$_.DocumentGUID;Pasta=$_.Pasta;FolderID=$_.FolderID;Workflow=$_.Workflow;Estado=$_.Estado;Resultado='Pendente';Detalhe='';PastaRelatorio=''}
    })
    New-Item -ItemType Directory -Path $Destino -Force | Out-Null
    $csv=Join-Path $Destino 'ResultadoDocumentos.csv'
    $contexto=Novo-ContextoACLARTESP
    $cronometro=[Diagnostics.Stopwatch]::StartNew()
    $duracoes=New-Object 'System.Collections.Generic.List[object]'
    $journal=Join-Path $Destino 'ProgressoDocumentos.csv'
    $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
    $fase='Planejamento'
    try {
        for($i=0;$i -lt $Linhas.Count;$i++) {
            $linha=$Linhas[$i];$registro=$resultados[$i]
            $registro.PastaRelatorio=Join-Path $Destino ('Documento_'+($i+1))
            $atuais=@(Medir-OperacaoACL 'ConsultarDocumento' $contexto { Get-PWDocumentsByGUIDs -DocumentGUIDs @($linha.DocumentGUID) -ErrorAction Stop })
            if ($atuais.Count -ne 1 -or $atuais[0].ProjectID -ne [int]$linha.FolderID -or
                $atuais[0].Workflow -ne $linha.Workflow -or $atuais[0].WorkflowState -ne $linha.Estado) { throw 'Documento mudou de pasta/workflow/estado desde a selecao.' }
            Write-Host ('Planejando {0}/{1}: {2}' -f ($i+1),$Linhas.Count,$linha.Documento) -ForegroundColor Cyan
            Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino $registro.PastaRelatorio -DocumentoLido $atuais[0] -Contexto $contexto -Silencioso -Reverter:$Reverter
            $registro.Resultado='Planejado'
            $registro | Select-Object *,@{n='Horario';e={Get-Date -Format o}} | Export-Csv $journal -Append -NoTypeInformation -Encoding UTF8
        }
        $duracoes.Add([pscustomobject]@{Etapa='Planejamento';Segundos=$cronometro.Elapsed.TotalSeconds})
        $cronometro.Stop()
        $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        $comandoConfirmacao=if($Reverter){'REVERTER LOTE'}else{'APLICAR LOTE'}
        if($Reverter){Write-Host 'Reversao: remover somente ARTESP da ACL propria de Workflow dos arquivos bloqueados; preservar demais acessos. Heranca ARTESP exige revisao da origem.' -ForegroundColor Yellow}
        else{
            Write-Host ("Plano completo: {0} documentos. ARTESP recebera Read + File Read no workflow/estado atual, preservando os demais acessos." -f $Linhas.Count) -ForegroundColor Yellow
            Write-Host 'A heranca atual sera materializada como permissoes proprias quando necessario; futuras mudancas na origem nao serao herdadas por essa ACL.' -ForegroundColor Yellow
        }
        if ((Read-Host "Digite $comandoConfirmacao para executar; Enter encerra com o plano salvo").Trim() -ine $comandoConfirmacao) {
            foreach($r in $resultados){$r.Resultado='Nao aplicado - confirmacao cancelada'}
            $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
            return
        }
        $fase='Aplicacao'
        $cronometro.Restart()
        for($i=0;$i -lt $Linhas.Count;$i++) {
            $linha=$Linhas[$i];$registro=$resultados[$i]
            $registro.Resultado='Em aplicacao - resultado pode ser parcial'
            $registro | Select-Object *,@{n='Horario';e={Get-Date -Format o}} | Export-Csv $journal -Append -NoTypeInformation -Encoding UTF8
            Write-Host ('Aplicando {0}/{1}: {2}' -f ($i+1),$Linhas.Count,$linha.Documento) -ForegroundColor Cyan
            Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino (Join-Path $registro.PastaRelatorio 'Aplicacao') -PlanoAprovado (Join-Path $registro.PastaRelatorio 'PlanoPreservacao.json') -Aplicar -Contexto $contexto -SelecaoEsperada $linha -Silencioso -Reverter:$Reverter
            $registro.Resultado=if($contexto.UltimoResultado){$contexto.UltimoResultado}else{'Concluido - ACL completa verificada'}
            $registro | Select-Object *,@{n='Horario';e={Get-Date -Format o}} | Export-Csv $journal -Append -NoTypeInformation -Encoding UTF8
        }
        $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        Write-Host "Lote concluido. Resultado por documento: $csv" -ForegroundColor Green
    }
    catch {
        if ($null -ne $registro) { $registro.Resultado="Falha em $fase - revisar; alteracoes podem ser parciais"; $registro.Detalhe=$_.Exception.Message }
        foreach($r in $resultados){if($r.Resultado -in @('Pendente','Planejado')){$r.Resultado='Nao executado - lote interrompido'}}
        $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        throw
    }
    finally {
        $cronometro.Stop()
        if ($fase -eq 'Aplicacao') { $duracoes.Add([pscustomobject]@{Etapa='Aplicacao';Segundos=$cronometro.Elapsed.TotalSeconds}) }
        elseif ($duracoes.Count -eq 0) { $duracoes.Add([pscustomobject]@{Etapa='Planejamento interrompido';Segundos=$cronometro.Elapsed.TotalSeconds}) }
        $duracoes | Export-Csv (Join-Path $Destino 'TemposEtapas.csv') -NoTypeInformation -Encoding UTF8
        @($contexto.Tempos.GetEnumerator() | ForEach-Object {
            [pscustomobject]@{Operacao=$_.Key;Quantidade=$_.Value.Quantidade;Segundos=$_.Value.Milissegundos/1000}
        }) | Export-Csv (Join-Path $Destino 'TemposOperacoes.csv') -NoTypeInformation -Encoding UTF8
    }
}
