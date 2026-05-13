function ObterDocumentosASeremValidados {
    $documentos = Get-PWDocumentsBySearch -SearchName 'Scripts\Docs Recebidos GRD' -GetAttributes
    return $documentos
}
function ValidaDocumento {
    param ($documento, $documentos)

    $mensagemErro = ValidaNumeroDocumento -Documento $documento
    if ($mensagemErro) {
        return $mensagemErro
    }

    # $mensagemErro = validaSePrevistoLD -documento $documento
    # if ($mensagemErro) {
    #     return $mensagemErro
    # }

    $possuiPDFENativo = ValidaSeDocumentoEditavelPossuiPDFEmitido -Documento $documento -Documentos $documentos
    if (-not $possuiPDFENativo) {
        return "Nao foram emitidos todos os arquivos Nativo e PDF deste documento."
    }

    $mensagemErro = ValidaRevisaoEVersao -Documento $documento

    return $mensagemErro
}
function validaSePrevistoLD {
    param ($documento)

    if ($documento.Attributes[0].PrevistoLD -eq 1) {
        return ''
    }
    return 'Documento não previsto na LD'
    
}
function ValidaSeDocumentoEditavelPossuiPDFEmitido {
    param ($documento, $documentos)

 
    $nomeBase = ''
    if ($documento.Name) { $nomeBase = $documento.Name }
    elseif ($documento.FileName) { $nomeBase = $documento.FileName }

    $ext = ''
    if ($nomeBase) { $ext = [System.IO.Path]::GetExtension($nomeBase) }
    if (-not $ext -and $documento.FileExtension) {
        $ext = "." + ($documento.FileExtension.ToLower().Trim())
    }
    if ($ext) { $ext = $ext.ToLower().Trim() }

    $extensoesSemPDF = @(
        '.kmz', '.kml', '.xml', '.jpg', '.jpeg', '.png',
        '.tif', '.tiff', '.3ds', '.dae', '.jgw', '.pgw',
        '.ecw', '.eww'
    )

    # Melhoria Modelos Autorais
    if (ValidaSeModeloFederadoAutoral $documento) { 
        $extensoesSemPDF += '.ifcZIP', '.bcf', '.bcfzip', 
        '.xml', '.landxml', '.rvt', '.ifc', '.nwd', '.ndw',
        '.dwg', '.npd', '.npl', '.ndx', '.npa', '.nlm', '.nma',
        '.3ds', '.las', '.laz', '.xyz', '.nwc', '.zip'
    }
   
    if ($ext -and ($extensoesSemPDF -contains $ext)) {
        return $true
    }
    
    $numeroDocumento = $documento.Attributes[0].NumeroPoderConcedente

    $documentosMesmoNumero = @(
        $documentos | Where-Object {
            $_.Attributes[0].NumeroPoderConcedente -eq $numeroDocumento -and
            [string]::IsNullOrEmpty($_.Attributes[0].Sufixo)
        }
    )

    if ($documentosMesmoNumero.Count -lt 2) {
        return $false
    }

    $documentosPDF = @(
        $documentosMesmoNumero | Where-Object {
            $nm = ''
            if ($_.Name) { $nm = $_.Name }
            elseif ($_.FileName) { $nm = $_.FileName }

            $e = ''
            if ($nm) { $e = [System.IO.Path]::GetExtension($nm) }
            if (-not $e -and $_.FileExtension) {
                $e = "." + ($_.FileExtension.ToLower().Trim())
            }
            if ($e) { $e = $e.ToLower().Trim() }

            $e -eq '.pdf'
        }
    )

    if ($documentosPDF.Count -lt 1) {
        return $false
    }

    return $true
}
function ValidaNumeroDocumento {
    param ($documento)
    
    if ($null -eq $documento -or $null -eq $documento.Attributes[0].NumeroPoderConcedente -or $documento.Attributes[0].NumeroPoderConcedente -eq '') {
        return 'Erro de taxonomia'
    }

    if ($documento.Attributes[0].NumeroPoderConcedente.Substring(0, 5) -eq 'Erro:') {
        return $documento.Attributes[0].NumeroPoderConcedente.Substring(11)
    }

    return $null
}
function ValidaRevisaoEVersao {
    param ($documento)
    
    # Melhoria: Exceção para modelos autorais
    if (ValidaSeModeloFederadoAutoral $documento) { return '' }

    $revisaoCorreta = ValidaSeRevisaoCorreta -Documento $documento
    $versaoCorreta = ValidaSeVersaoCorreta -Documento $documento
    $validaRevisaoEVersao = $revisaoCorreta -bor $versaoCorreta

    $revisaoCalculada = $documento.Attributes[0].RevisaoCalculada
    $versaoCalculada = $documento.Attributes[0].VersaoCalculada
    
    if (-not $revisaoCalculada) {
        if (-not $versaoCalculada) {
            return "Nao foi possivel calcular a revisao e a versao correta"
        }
        return "Nao foi possível calcular a revisao correta"
    }
    if (-not $versaoCalculada) {
        return "Nao foi possivel calcular a versao correta"
    }

    $mensagemFinal = switch ($validaRevisaoEVersao) {
        1 { "Revisao incorreta. O correto seria: $revisaoCalculada" }
        2 { "Versao incorreta. A versao correta seria: $versaoCalculada" }
        3 { "Revisao e Versao incorreta. O correto seria, Revisao: $revisaoCalculada Versao: $versaoCalculada" }
        Default { "" }
    }

    return $mensagemFinal
}
function ValidaSeRevisaoCorreta {
    param ($documento)
    
    if ($documento.Attributes[0].Revisao -eq $documento.Attributes[0].RevisaoCalculada) {
        return 0
    }
    return 1
}
function ValidaSeVersaoCorreta {
    param ($documento)

    if ($documento.Attributes[0].Versao -eq $documento.Attributes[0].VersaoCalculada) {
        return 0
    }
    return 2
}

function ValidaSeModeloFederadoAutoral {
    param($documento)

    if ($documento.Attributes[0].PoderConcedente -eq 'ANTT' -and $documento.Attributes[0].Sequencial -as [int] -in 800..999) { return $true }
    if ($documento.Attributes[0].PoderConcedente -eq 'ARTESP' -and $documento.Attributes[0].TipoDocumento -in @('MB', 'MI')) { return $true }

    return $false
}

#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' -Password $SecurePassword -UserName 'admin'

$documentos = ObterDocumentosASeremValidados

foreach ($documento in $documentos) {
    $mensagemErro = ValidaDocumento -Documento $documento -Documentos $documentos
    
    if ($mensagemErro -eq 'Nao foram emitidos todos os arquivos Nativo e PDF deste documento.') {
        $tempoDesdeCriado = (Get-Date) - $documento.CreateDate
        if ($tempoDesdeCriado.TotalHours -lt 1) {
            continue
        }
    }

    if ($mensagemErro) {
        Update-PWDocumentAttributes -InputDocuments $documento -Attributes @{Erros = $mensagemErro }
        Set-PWDocumentState -InputDocuments $documento -State 'Nao Validado pelo Sistema' -Force
    }
    else {
        Set-PWDocumentState -InputDocuments $documento -State 'Validado pelo Sistema - A ser movido' -Force
    }
}

Undo-PWLogin