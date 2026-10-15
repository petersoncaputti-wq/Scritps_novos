# Planejamento, leitura e aplicacao explicita com preservacao de participantes.
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
    param($Documento)
    $pastas = @(Get-PWFolders -FolderID $Documento.ProjectID -JustOne -ErrorAction Stop | Where-Object { $null -ne $_ -and $_ -isnot [bool] })
    if ($pastas.Count -ne 1) { throw 'Pasta do documento nao localizada de forma unica.' }
    $dados = Get-PWFolderSecurity -InputFolder $pastas[0] -ErrorAction Stop
    if ($dados -is [Data.DataTable]) { $dados = $dados.Rows }
    $origem = @($dados | Where-Object { $_.SecurityType -eq 'Document' -and $_.Workflow -eq $Documento.Workflow -and $_.State -eq $Documento.WorkflowState })
    if ($origem.Count -eq 0) { throw 'Nao foi possivel identificar todas as regras herdadas do workflow/estado na pasta. Escrita bloqueada.' }
    $c = [PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWPS_DAB.CommonAccessControlMethods', $true)
    $m = $c.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static')
    $idMetodo = @($m | Where-Object Name -eq 'getMemberID')[0]
    $mascaraMetodo = @($m | Where-Object Name -eq 'ConvertMaskStringToInt')[0]
    foreach ($r in $origem) {
        $tipo=0; $id=0
        switch ([string]$r.Type) {
            'User' { $tipo=1 }
            'Group' { $tipo=2 }
            'UserList' { $tipo=3 }
            'All_Users (*Everyone)' { $tipo=4 }
            default { throw "Tipo de participante desconhecido: $($r.Type)" }
        }
        if ($tipo -ne 4) {
            $id=[int]$idMetodo.Invoke($null,[object[]]@([string]$r.Name,$tipo))
            if ($id -le 0) { throw "Participante nao resolvido: $($r.Name)" }
        }
        $mask=[int]$mascaraMetodo.Invoke($null,[object[]]@([string]$r.Access_Control_Settings,'Document'))
        if ($mask -lt 0) { throw "Mascara invalida para $($r.Name)" }
        [pscustomobject]@{MemberType=$tipo;MemberId=$id;AccessMask=[uint32]$mask;Name=[string]$r.Name;Origem=[string]$r.InheritingFrom}
    }
}

function Atualizar-ACLWorkflowARTESP {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GUID, [Parameter(Mandatory)][string]$Destino,
        [switch]$Aplicar, [switch]$RecuperarOrigem, [string]$PlanoAprovado)
    $docs=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop | Where-Object { $null -ne $_ -and $_ -isnot [bool] })
    if ($docs.Count -ne 1) { throw 'Documento nao localizado de forma unica.' }
    $doc=$docs[0]
    if ($doc.ProjectID -le 0 -or $doc.DocumentID -le 0 -or $doc.WorkflowId -le 0 -or $doc.StateId -le 0) { throw 'Identificadores incompletos.' }
    if ($RecuperarOrigem -and $GUID -ne '5fc0819d-db6f-4841-a979-75a387bb1d01') { throw 'Recuperacao restrita ao PDF testado.' }
    $w=[PWPS_DAB.SetPWDocumentGroupAccess].Assembly.GetType('PWWrapper',$true)
    $metodos=$w.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static')
    $gm=@($metodos | Where-Object Name -eq 'GetGroupsByName')[0]
    $grupos=$gm.Invoke($null,@())
    $g=@($grupos.Values | Where-Object Name -eq 'ARTESP')
    if ($g.Count -ne 1) { throw 'Grupo ARTESP nao localizado.' }
    $antes=@(Ler-ACLWorkflowPropria $doc)
    $base=@($antes); $herdadas=@()
    if ($antes.Count -eq 0 -or $RecuperarOrigem) {
        $herdadas=@(Ler-ACLWorkflowOrigem $doc)
        foreach($h in $herdadas) {
            $existente=@($base | Where-Object { $_.MemberType -eq $h.MemberType -and $_.MemberId -eq $h.MemberId })
            if ($existente.Count -eq 0) { $base += $h }
            elseif ($existente.Count -ne 1 -or $existente[0].AccessMask -ne $h.AccessMask) { throw 'Conflito entre ACL propria e origem. Revisao necessaria.' }
        }
    }
    $plano=@(Mesclar-ACLARTESP $base ([int]$g[0].ID))
    if ($PlanoAprovado) {
        $aprovado=Get-Content -LiteralPath $PlanoAprovado -Raw | ConvertFrom-Json
        if ($aprovado.GUID -ne $GUID -or $aprovado.WorkflowID -ne $doc.WorkflowId -or $aprovado.StateID -ne $doc.StateId) { throw 'Documento/workflow/estado diverge do plano aprovado.' }
        $null=Confirmar-ACLCompleta @($aprovado.ACLPropriaAntes) $antes
        $null=Confirmar-ACLCompleta @($aprovado.ACLEsperada) $plano
    }
    New-Item -ItemType Directory -Path $Destino -Force | Out-Null
    $doc | Export-Clixml (Join-Path $Destino 'DocumentoAntes.xml') -Depth 5
    [pscustomobject]@{GUID=$GUID;WorkflowID=$doc.WorkflowId;StateID=$doc.StateId;ACLPropriaAntes=$antes;ACLOrigem=$herdadas;ACLEsperada=$plano} |
        ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Destino 'PlanoPreservacao.json') -Encoding UTF8
    $plano | Format-Table MemberType,MemberId,AccessMask -AutoSize | Out-Host
    if (-not $Aplicar) { Write-Host 'Plano salvo; nenhuma alteracao aplicada.'; return }
    $atual=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop)
    if ($atual.Count -ne 1 -or $atual[0].ProjectID -ne $doc.ProjectID -or $atual[0].DocumentID -ne $doc.DocumentID -or
        $atual[0].WorkflowId -ne $doc.WorkflowId -or $atual[0].StateId -ne $doc.StateId) { throw 'Documento mudou desde o planejamento.' }
    $null=Confirmar-ACLCompleta $antes @(Ler-ACLWorkflowPropria $doc)
    $assign=@($metodos | Where-Object Name -eq 'aaApi_AssignAccessList')[0]
    try {
        # Preserva/materializa todos os demais participantes antes de acrescentar ARTESP.
        foreach($r in @($plano | Sort-Object @{Expression={if($_.MemberType -eq 2 -and $_.MemberId -eq $g[0].ID){1}else{0}}})) {
            $existente=@($antes | Where-Object { $_.MemberType -eq $r.MemberType -and $_.MemberId -eq $r.MemberId })
            if ($existente.Count -gt 0) { continue }
            $ok=[bool]$assign.Invoke($null,[object[]]@(4,[int]$doc.ProjectID,[int]$doc.DocumentID,[int]$doc.WorkflowId,[int]$doc.StateId,[int]$r.MemberType,[int]$r.MemberId,[uint32]$r.AccessMask))
            if (-not $ok) { throw "Falha ao aplicar participante $($r.MemberType):$($r.MemberId). Resultado pode ser parcial; interrompido." }
        }
        $depois=@(Ler-ACLWorkflowPropria $doc)
        $depois | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Destino 'ACLDepois.json') -Encoding UTF8
        $null=Confirmar-ACLCompleta $plano $depois
        $final=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($GUID) -ErrorAction Stop)
        if ($final.Count -ne 1 -or $final[0].WorkflowId -ne $doc.WorkflowId -or $final[0].StateId -ne $doc.StateId) { throw 'Workflow/estado divergente apos aplicacao.' }
        'ACL completa confirmada; demais participantes e direitos preservados. Conferencia visual pendente.' |
            Set-Content (Join-Path $Destino 'Resultado.txt') -Encoding UTF8
    }
    catch { $_ | Out-String | Set-Content (Join-Path $Destino 'Erro.txt') -Encoding UTF8; throw }
}

function Executar-LoteARTESP {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Linhas,[Parameter(Mandatory)][string]$Destino)
    if ($Linhas.Count -eq 0) { throw 'Nenhum documento selecionado.' }
    $guids=@{}
    $resultados=@($Linhas | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_.DocumentGUID)) { throw 'Documento selecionado sem GUID.' }
        if ($guids.ContainsKey($_.DocumentGUID)) { throw 'GUID duplicado na selecao.' }
        $guids[$_.DocumentGUID]=$true
        [pscustomobject]@{Documento=$_.Documento;GUID=$_.DocumentGUID;Workflow=$_.Workflow;Estado=$_.Estado;Resultado='Pendente';Detalhe='';PastaRelatorio=''}
    })
    New-Item -ItemType Directory -Path $Destino -Force | Out-Null
    $csv=Join-Path $Destino 'ResultadoDocumentos.csv'
    $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
    $fase='Planejamento'
    try {
        for($i=0;$i -lt $Linhas.Count;$i++) {
            $linha=$Linhas[$i];$registro=$resultados[$i]
            $registro.PastaRelatorio=Join-Path $Destino ('Documento_'+($i+1))
            $atuais=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($linha.DocumentGUID) -ErrorAction Stop)
            if ($atuais.Count -ne 1 -or $atuais[0].ProjectID -ne [int]$linha.FolderID -or
                $atuais[0].Workflow -ne $linha.Workflow -or $atuais[0].WorkflowState -ne $linha.Estado) { throw 'Documento mudou de pasta/workflow/estado desde a selecao.' }
            Write-Host ('Planejando {0}/{1}: {2}' -f ($i+1),$Linhas.Count,$linha.Documento) -ForegroundColor Cyan
            Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino $registro.PastaRelatorio
            $registro.Resultado='Planejado'
            $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        }
        Write-Host ("Plano completo: {0} documentos. ARTESP recebera Read + File Read no workflow/estado atual, preservando os demais acessos." -f $Linhas.Count) -ForegroundColor Yellow
        Write-Host 'A heranca atual sera materializada como permissoes proprias quando necessario; futuras mudancas na origem nao serao herdadas por essa ACL.' -ForegroundColor Yellow
        if ((Read-Host 'Digite APLICAR LOTE para executar; Enter encerra com o plano salvo').Trim() -ine 'APLICAR LOTE') {
            foreach($r in $resultados){$r.Resultado='Nao aplicado - confirmacao cancelada'}
            $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
            return
        }
        $fase='Aplicacao'
        for($i=0;$i -lt $Linhas.Count;$i++) {
            $linha=$Linhas[$i];$registro=$resultados[$i]
            # Revalida tambem a pasta imediatamente antes da aplicacao.
            $atuais=@(Get-PWDocumentsByGUIDs -DocumentGUIDs @($linha.DocumentGUID) -ErrorAction Stop)
            if ($atuais.Count -ne 1 -or $atuais[0].ProjectID -ne [int]$linha.FolderID -or
                $atuais[0].Workflow -ne $linha.Workflow -or $atuais[0].WorkflowState -ne $linha.Estado) { throw 'Documento mudou desde a revisao do lote.' }
            $registro.Resultado='Em aplicacao - resultado pode ser parcial'
            $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
            Write-Host ('Aplicando {0}/{1}: {2}' -f ($i+1),$Linhas.Count,$linha.Documento) -ForegroundColor Cyan
            Atualizar-ACLWorkflowARTESP -GUID $linha.DocumentGUID -Destino (Join-Path $registro.PastaRelatorio 'Aplicacao') -PlanoAprovado (Join-Path $registro.PastaRelatorio 'PlanoPreservacao.json') -Aplicar
            $registro.Resultado='Concluido - ACL completa verificada'
            $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Lote concluido. Resultado por documento: $csv" -ForegroundColor Green
    }
    catch {
        if ($null -ne $registro) { $registro.Resultado="Falha em $fase - revisar; alteracoes podem ser parciais"; $registro.Detalhe=$_.Exception.Message }
        foreach($r in $resultados){if($r.Resultado -in @('Pendente','Planejado')){$r.Resultado='Nao executado - lote interrompido'}}
        $resultados | Export-Csv $csv -NoTypeInformation -Encoding UTF8
        throw
    }
}
