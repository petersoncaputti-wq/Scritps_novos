# -----------------------------------------------
# Execução 
# -----------------------------------------------
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "Logando..."
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
$logado = New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-01' -Password $SecurePassword -UserName 'admin'
if (-not $logado) { return }
Write-Host "Login levou [$($sw.Elapsed)]"; $sw.Restart()

Write-Host "Buscando Superados..."
# Retorna todos os documentos dentro de pastas chamadas "Superados"
$listaDocumentosSuperados = Select-PWSQL "SELECT REPLACE(CONCAT(dbo.GetVaultPath(P.o_projectno, 0), D.o_itemname),'/','\') AS Fullpath        
FROM dms_proj AS P
INNER JOIN dms_doc AS D
ON D.o_projectno = P.o_projectno
WHERE P.o_projectname = 'Superados'
AND RIGHT(D.o_filename,4) = '.pdf' ;"
Write-Host "Buscando Superados levou [$($sw.Elapsed)], $($listaDocumentosSuperados.Rows.Count) registros encontrados"; $sw.Restart()

Write-Host "Buscando XFDF..."
# Retorna todos os documentos dentro de pastas chamadas "xfdf__"
$listaDocumentosXFDF = Select-PWSQL "SELECT REPLACE(CONCAT(dbo.GetVaultPath(P.o_projectno, 0), D.o_itemname),'/','\') AS Fullpath        
FROM dms_proj AS P
INNER JOIN dms_doc AS D
ON D.o_projectno = P.o_projectno
WHERE P.o_projectname = 'xfdf__'
AND RIGHT(D.o_filename,5) = '.xfdf' ;"
Write-Host "Buscando XFDF levou [$($sw.Elapsed)], $($listaDocumentosXFDF.Rows.Count) registros encontrados"; $sw.Restart()

Write-Host "Ajustando dados retornados..."
$listaDocumentosSuperados = $listaDocumentosSuperados.Rows.Fullpath | Sort-Object -Unique
$hashSetDocumentosSuperados = [System.Collections.Generic.HashSet[string]]::new([string[]]$listaDocumentosSuperados)

$listaDocumentosXFDF = $listaDocumentosXFDF.Rows.Fullpath | Where-Object { -not $_.Contains('\Superados\xfdf__\') } | Sort-Object -Unique
Write-Host "Ajustando dados levou [$($sw.Elapsed)], $($hashSetDocumentosSuperados.Count) Superados únicos, $($listaDocumentosXFDF.Count) XFDFs únicos"; $sw.Restart()

Write-Host "Executando loop de verificação..."
$listaXFDFParaMover = @()
$count = 0
$total = $listaDocumentosXFDF.Count
foreach ($documentoXFDF in $listaDocumentosXFDF) {
    $count++
    [Console]::WriteLine("[$count/$total] - Verificando XFDF: Path ($documentoXFDF)")
    
    $superadoCorrespondente = $documentoXFDF.Replace('\xfdf__', '\Superados').Replace('.xfdf', '.pdf')
    
    if ($hashSetDocumentosSuperados.Contains($superadoCorrespondente)) {
        [Console]::WriteLine("Superado correspondente encontrado")
        $listaXFDFParaMover += $documentoXFDF
    }
}
Write-Host "Executando Loop de verificação levou [$($sw.Elapsed)], $($listaXFDFParaMover.Count) XFDFs para mover encontrados"; $sw.Restart()

Write-Host "Executando loop de movimentação..."
$count = 0
$total = $listaXFDFParaMover.Count
foreach ($pathXFDF in $listaXFDFParaMover) {
    $count++
    Write-Host "[$count/$total] - Movendo XFDF: Path ($pathXFDF)"
    
    try {
        $pathPastaXFDF = ($pathXFDF.Substring(0, $pathXFDF.LastIndexOf('\')))
        $nomeXFDF = $pathXFDF.Substring($pathXFDF.LastIndexOf('\') + 1)
        
        # Buscar ou criar pasta destino do XFDF
        $pathDestinoXFDF = $pathPastaXFDF.Replace('\xfdf__', '\Superados\xfdf__')
        $pastaDestinoXFDF = Get-PWFolders -FolderPath $pathDestinoXFDF
        if (-not $pastaDestinoXFDF) {
            Write-Host "[AVISO] Criando diretório de destino: ($pathXFDF)"
            $pastaDestinoXFDF = New-PWFolder -FolderPath $pathDestinoXFDF -Environment 'dmsXFDF' -Workflow 'Workflow - Comentarios'
            if (-not $pastaDestinoXFDF) {
                throw "Diretório de destino do XFDF inexistente e falha na tentativa de criação."
            }
            $pastaNova = $true
        }
        else {
            $pastaNova = $false
        }

        # Verificar se já existe um XFDF para esse Superado
        if (-not $pastaNova) {
            $antigosXFDFs = Select-PWSQL "SELECT o_itemname FROM dms_doc WHERE o_itemname = '$nomeXFDF' AND o_projectno = $($pastaDestinoXFDF.ProjectID)"
            if ($antigosXFDFs -and $antigosXFDFs.Rows.o_itemname -ne '') {
                throw "Diretório de destino do XFDF já possui um arquivo com o mesmo nome, nenhuma ação será realizada"
            }
        }
        
        # Resgatar ultima versão do XFDF e MOVER para a pasta destino
        $documentoXFDF = Get-PWDocumentsBySearch -FolderPath $pathPastaXFDF -DocumentName $nomeXFDF
        if (-not $documentoXFDF) {
            throw "Arquivo do XFDF não foi encontrado para ser movido"
        }
        Move-PWDocumentsToFolder -InputDocument $documentoXFDF -TargetFolderPath $pathDestinoXFDF -Verbose
    }
    catch {
        Write-Host "[ERRO] - $($_.Exception)"
    }
    
}
Write-Host "Executando Loop de movimentação levou [$($sw.Elapsed)]"; $sw.Restart()
Write-Host "Tempo total [$($swTotal.Elapsed)]"; $sw.Stop(); $swTotal.Stop()
Undo-PWLogin | Out-Null