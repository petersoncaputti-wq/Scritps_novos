function ZipXml($Archive,[string]$Name){
    $entry=$Archive.GetEntry($Name);if(!$entry){return $null}
    $stream=$entry.Open();$reader=[Xml.XmlReader]::Create($stream,[Xml.XmlReaderSettings]@{DtdProcessing=[Xml.DtdProcessing]::Prohibit;XmlResolver=$null})
    try{$doc=New-Object Xml.XmlDocument;$doc.XmlResolver=$null;$doc.Load($reader);return ,$doc}finally{$reader.Dispose();$stream.Dispose()}
}
function LoadProjects([string]$Path){
    Add-Type -AssemblyName System.IO.Compression.FileSystem,System.IO.Compression
    $zip=[IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($Path))
    try{
        $strings=New-Object Collections.Generic.List[string]
        $xml=ZipXml $zip 'xl/sharedStrings.xml'
        if($xml){foreach($si in $xml.SelectNodes('/*[local-name()="sst"]/*[local-name()="si"]')){$text='';foreach($t in $si.SelectNodes('.//*[local-name()="t"]')){$text+=$t.InnerText};$strings.Add($text)}}
        $book=ZipXml $zip 'xl/workbook.xml';$rels=ZipXml $zip 'xl/_rels/workbook.xml.rels';$selected=$null
        foreach($s in $book.SelectNodes('//*[local-name()="sheets"]/*[local-name()="sheet"]')){if(!$Sheet -or $s.GetAttribute('name') -ceq $Sheet){$selected=$s;break}}
        if(!$selected){throw 'Aba nao encontrada na planilha.'}
        $rid=$selected.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships');$target=$null
        foreach($r in $rels.DocumentElement.ChildNodes){if($r.Id -eq $rid){$target=$r.Target}}
        if(!$target){throw 'Relacionamento da aba nao encontrado.'}
        $uri=[uri]::new([uri]'https://xlsx.local/xl/workbook.xml',[string]$target);$name=$uri.AbsolutePath.TrimStart('/')
        $xml=ZipXml $zip $name;if(!$xml){throw 'XML da aba nao encontrado.'}
        $slash=$name.LastIndexOf('/');$relName=$name.Substring(0,$slash)+'/_rels/'+$name.Substring($slash+1)+'.rels'
        $sheetRels=ZipXml $zip $relName;$links=@{};$hyperlinks=@{}
        if($sheetRels){foreach($r in $sheetRels.DocumentElement.ChildNodes){if($r.TargetMode -eq 'External'){$links[$r.Id]=$r.Target}}}
        foreach($h in $xml.SelectNodes('//*[local-name()="hyperlinks"]/*[local-name()="hyperlink"]')){
            $id=$h.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            if($id -and $links.ContainsKey($id)){$hyperlinks[$h.GetAttribute('ref')]=$links[$id]}
        }
        $projects=New-Object Collections.ArrayList;$issues=New-Object Collections.ArrayList;$seen=@{}
        foreach($row in $xml.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')){
            $number=[int]$row.GetAttribute('r');$cells=@{}
            foreach($cell in $row.SelectNodes('./*[local-name()="c"]')){
                $column=$cell.GetAttribute('r') -replace '[0-9]','';$v=$cell.SelectSingleNode('./*[local-name()="v"]');$value=if($v){$v.InnerText}else{''}
                if($cell.GetAttribute('t') -eq 's'){$value=if($value){$strings[[int]$value]}else{''}}
                elseif($cell.GetAttribute('t') -eq 'inlineStr'){$value='';foreach($t in $cell.SelectNodes('.//*[local-name()="t"]')){$value+=$t.InnerText}}
                $cells[$column]=$value.Trim()
            }
            if($cells['E'] -ieq 'prioridade' -or !($cells['C'] -or $cells['E'] -or $cells['F'])){continue}
            if($ProjectRows -and $number -notin $ProjectRows){continue}
            $p=0
            if(![int]::TryParse([string]$cells['E'],[ref]$p) -or $p -lt 1){[void]$issues.Add(@{Linha=$number;Motivo='Prioridade ausente ou invalida.'});continue}
            if($Priority -and $p -notin $Priority){continue}
            $url=$hyperlinks['F'+$number];if(!$url){$url=$cells['F']}
            try{$parsed=ParseUrl $url}catch{[void]$issues.Add(@{Linha=$number;Motivo='Link ausente ou nao suportado na coluna F.'});continue}
            if($seen.ContainsKey($parsed.Url)){[void]$issues.Add(@{Linha=$number;Motivo='Pasta duplicada na fila.'});continue};$seen[$parsed.Url]=$true
            $folder=LocalName ('Ano_'+$cells['A']+'_Lote_'+$cells['B']+'_SH_'+$cells['C']+'_'+(HashText $parsed.Url).Substring(0,12))
            [void]$projects.Add([pscustomobject]@{Row=$number;Priority=$p;Folder=$folder;Parsed=$parsed})
        }
        [pscustomobject]@{Projects=@($projects|Sort-Object Priority,Row);Issues=@($issues.ToArray())}
    }finally{$zip.Dispose()}
}
function RunQueue($Queue,[string]$Destination,[switch]$FirstFolderOpened){
    [void][IO.Directory]::CreateDirectory($Destination)
    $reportPath=Join-Path $Destination ('fila_projetos_resumo_'+[guid]::NewGuid().ToString('N').Substring(0,12)+'.json')
    $records=New-Object Collections.ArrayList;$report=@{pendencias_planilha=$Queue.Issues;projetos=$records}
    SaveJson $reportPath $report;$failures=$Queue.Issues.Count
    $first=$true
    foreach($project in $Queue.Projects){
        $target=Join-Path $Destination $project.Folder
        $record=@{linha=$project.Row;prioridade=$project.Priority;destino=$target;estado='em_andamento'}
        [void]$records.Add($record);SaveJson $reportPath $report
        Write-Host "Projeto linha $($project.Row), prioridade $($project.Priority)"
        try{
            if(!$first -or !$FirstFolderOpened){OpenFolder $project.Parsed}
            $result=RunBatch $project.Parsed $target
            if($result.Failed){$record.estado='com_pendencias';$failures++}else{$record.estado='concluido'}
        }catch [System.Management.Automation.PipelineStoppedException]{$record.estado='interrompido';throw}
        catch{$record.estado='falha';$record.erro=[SharePointPS.Transfer]::Clean($_.Exception.Message);$failures++;Write-Warning "Projeto com falha: $($record.erro). Seguindo para o proximo."}
        finally{$first=$false;SaveJson $reportPath $report;Log 'FILA_PROJETO_FIM' $record}
    }
    Write-Host "Fila finalizada. Resumo: $reportPath"
    if($failures){return 1};return 0
}
function Diagnose([string]$Remote){
    if(!$Remote.StartsWith($script:RemoteRoot+'/',[StringComparison]::OrdinalIgnoreCase)){throw 'Arquivo de diagnostico fora da pasta informada.'}
    $api=RemoteEndpoint File $Remote;$jar=CookieJar;$metadata=$null
    try{$metadata=BrowserJson ($api+'?$select=Name,Length,Exists,UniqueId,TimeLastModified');Write-Host "Metadados no navegador: $($metadata.Name), $($metadata.Length) bytes";Log 'DIAGNOSTICO_METADADOS' @{nome=$metadata.Name;bytes=$metadata.Length}}
    catch{Write-Warning 'Metadados no navegador: leitura recusada.'}
    $tests=New-Object Collections.ArrayList
    [void]$tests.Add(@{Label='metadados_terminal';Url=$api+'?$select=Name,Length,Exists,UniqueId,TimeLastModified'})
    $i=0;foreach($route in @(Routes $Remote)){[void]$tests.Add(@{Label=@('rest_caminho','url_direta','download_aspx')[$i];Url=$route});$i++}
    $id=[guid]::Empty
    if($metadata -and [guid]::TryParse([string]$metadata.UniqueId,[ref]$id)){[void]$tests.Add(@{Label='conteudo_por_id';Url=$script:Web+"/_api/web/GetFileById('$id')/`$value"})}
    foreach($test in $tests){$result=[SharePointPS.Transfer]::Probe($test.Url,$jar);Write-Host "$($test.Label): $result";Log 'DIAGNOSTICO_RESULTADO' @{teste=$test.Label;resultado=$result;remoto=$Remote}}
    Write-Host 'Diagnostico concluido: amostras de ate 32 KiB; nenhum documento salvo.'
}
function BrowserTest($Parsed,[string]$Destination){
    $remote=$Parsed.Path
    while($true){
        $api=RemoteEndpoint Folder $remote
        $files=@(Collection ($api+'/Files?$select=Name,Length,ServerRelativeUrl,ETag&$top=20')|Select-Object -First 20)
        $folders=@(Collection ($api+'/Folders?$select=Name,ServerRelativeUrl&$top=200')|Select-Object -First 200)
        Write-Host "Pasta: $remote"
        for($n=0;$n -lt $folders.Count;$n++){Write-Host "P$($n+1). $($folders[$n].Name)"}
        $choice=Read-Host 'P e numero para abrir subpasta; ENTER para usar esta pasta'
        if(!$choice){break}
        if($choice -match '^P([0-9]+)$' -and [int]$Matches[1] -ge 1 -and [int]$Matches[1] -le $folders.Count){
            $next=$folders[[int]$Matches[1]-1].ServerRelativeUrl
            if(!$next.StartsWith($remote+'/',[StringComparison]::OrdinalIgnoreCase)){throw 'Subpasta fora da origem.'};$remote=$next
        }else{Write-Warning 'Selecao invalida.'}
    }
    for($n=0;$n -lt $files.Count;$n++){Write-Host "$($n+1). $($files[$n].Name) ($($files[$n].Length) bytes)"}
    if(!$files.Count){Write-Host 'Nenhum arquivo direto nesta pasta.';return}
    $choice=Read-Host 'Numero do arquivo pequeno para testar (ENTER encerra)';if(!$choice){return}
    $n=0;if(![int]::TryParse($choice,[ref]$n) -or $n -lt 1 -or $n -gt $files.Count){throw 'Numero invalido.'}
    $item=$files[$n-1];if([long]$item.Length -gt 100MB){throw 'Teste limitado a arquivos de ate 100 MB.'}
    if(!$item.ServerRelativeUrl.StartsWith($remote+'/',[StringComparison]::OrdinalIgnoreCase)){throw 'Arquivo fora da origem.'}
    [void][IO.Directory]::CreateDirectory($Destination);$target=SafePath (Join-Path $Destination (LocalName $item.Name)) $Destination
    if(Test-Path -LiteralPath $target){throw 'Destino existente; preservado.'}
    $transfer=New-Object SharePointPS.Transfer
    try{
        $transfer.Start([string[]]@(Routes $item.ServerRelativeUrl),$target,[long]$item.Length,($item.ServerRelativeUrl+'|'+$item.ETag),(CookieJar))
        while(!$transfer.Task.IsCompleted){Write-Progress -Activity 'Teste de download' -Status (BytesText $transfer.Bytes);Start-Sleep -Milliseconds 200}
        if(!$transfer.Task.Result.Success){throw $transfer.Task.Result.Message}
        Write-Host "Download validado por tamanho: $target"
    }finally{$transfer.Cancel=$true;[void]$transfer.Task.Wait(65000);Write-Progress -Activity 'Teste de download' -Completed}
}
