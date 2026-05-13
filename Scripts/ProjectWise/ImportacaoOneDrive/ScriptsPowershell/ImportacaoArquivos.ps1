Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
$my_other_script = [System.IO.Path]::GetFullPath($PSScriptRoot + "\..\..\Funcoes\CriarProjeto.ps1")
. $my_other_script

$script:pathPastaMapeamento = Resolve-Path -Path "$PSScriptRoot\..\Mapeamento\" | Select-Object -ExpandProperty "Path"

function JanelaSeletorExcel {
    param ()

    $janela = New-Object System.Windows.Forms.OpenFileDialog

    $janela.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $janela.Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
    $janela.Title = "Selecione Excel com os dados dos projetos a serem criados"

    return $janela
}
function SelecionarPlanilhaCadastrosDeProjetos {
    param ()
    
    $janela = JanelaSeletorExcel

    if ($janela.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $enderecoArquivoExcel = $janela.FileName     
        $dados = Import-Excel -Path $enderecoArquivoExcel

        ImportarProjetos -DadosProjetos $dados
        [System.Windows.Forms.MessageBox]::Show("Projetos importados com sucesso!", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}
function ImportarProjetos {
    param ($dadosProjetos)    

    foreach ($dadosProjeto in $dadosProjetos) {
        
        $nomeConcessao = $dadosProjeto.'Nome Concessao'
        $siglaConcessao = $dadosProjeto.'Sigla Concessao'
        $projeto = $dadosProjeto.'Nome Projeto'
        $descricao = $dadosProjeto.'Nome Projeto' + ' - ' + $dadosProjeto.'Descrição'
        $descricao = if($dadosProjeto.'TH') { $descricao + ' - ' + $dadosProjeto.'TH' } else{ $descricao }
        $projetista = $dadosProjeto.'Projetista'
        $gestorEngenharia = $dadosProjeto.'Gestor Engenharia'
        $assistenteEngenharia = $dadosProjeto.'Assistente Engenharia'
        $poderConcedente = $dadosProjeto.'Poder Concedente'
       
        if ($projeto.Length -gt 28) {
            $mensagemErroNomeProjetoGrande = "O nome do Projeto não pode ter mais que 28 caracteres!`r`n`r`nClique em OK para ignorar este projeto e continuar criando os demais.`r`n`r`nClique em CANCELAR para encerrar a execução do script.`r`n`r`nProjeto: ""$projeto"" possui $($projeto.Length) caracteres."
            Write-Host $mensagemErroNomeProjetoGrande
            continue            
        }
        
        $pastaRaizProjeto = CriarProjeto -PoderConcedente $poderConcedente -NomeConcessao $nomeConcessao -SiglaConcessao $siglaConcessao -Projeto $projeto -Descricao $descricao -Projetista $projetista -GestorEngenharia $gestorEngenharia -AssistenteEngenharia $assistenteEngenharia

        ImportarProjeto -projeto $projeto -poderConcedente $poderConcedente -siglaConcessao $siglaConcessao -pastaRaizProjeto $pastaRaizProjeto
    }
}
function ImportarProjeto {
    param ($projeto, $poderConcedente, $siglaConcessao, $pastaRaizProjeto)

    $cadastrosDisciplina = ObtemListaCadastrosDisciplina -poderConcedente $poderConcedente
    $cadastrosFaseProjeto = ObtemListaCadastrosFaseProjeto -poderConcedente $poderConcedente

    importarDocumentosSharepoint -projeto $projeto -poderConcedente $poderConcedente -siglaConcessao $siglaConcessao -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto
    ImportarDocumentosRede -projeto $projeto -poderConcedente $poderConcedente -siglaConcessao $siglaConcessao -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto
}
function importarDocumentosSharepoint {
    param ($projeto, $poderConcedente, $siglaConcessao, $pastaRaizProjeto, $cadastrosDisciplina, $cadastrosFaseProjeto)
    
    importarDocumentos -projeto $projeto -poderConcedente $poderConcedente -siglaConcessao $siglaConcessao -prefixo "OD" -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto
}
function ImportarDocumentosRede {
    param ($projeto, $poderConcedente, $siglaConcessao, $pastaRaizProjeto, $cadastrosDisciplina, $cadastrosFaseProjeto)

    importarDocumentos -projeto $projeto -poderConcedente $poderConcedente -siglaConcessao $siglaConcessao -prefixo "PR" -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto
}
function importarDocumentos {
    param ($projeto, $prefixo, $poderConcedente, $siglaConcessao, $pastaRaizProjeto, $cadastrosDisciplina, $cadastrosFaseProjeto)

    $dados = carregarArquivosMapeados -projeto $projeto -prefixo $prefixo
    if(-not $dados){
        return
    }

    $dadosAnalisados = validaListaImportacaoArquivos -dados $dados -siglaConcessao $siglaConcessao -poderConcedente $poderConcedente -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto

    salvarRelatorioArquivosMapeadosEAnalisados -dados $dadosAnalisados -projeto $projeto -prefixo $prefixo

    importarDocumentosPW -dadosAnalisados $dadosAnalisados -projeto $projeto -prefixo $prefixo
}
function carregarArquivosMapeados {
    param ($projeto, $prefixo)
    
    $nomeProjeto = $projeto.Replace("*", "(Asterisco)")
    $arquivoMapeamento = $script:pathPastaMapeamento + "$prefixo`_$nomeProjeto.csv"
    if(Test-Path $arquivoMapeamento){
        $dados = Import-Csv -Path $arquivoMapeamento -Delimiter ";"
        return $dados
    }
    return $null
}
function salvarRelatorioArquivosMapeadosEAnalisados {
    param ($dados, $projeto, $prefixo)
    
    $nomeProjeto = $projeto.Replace("*", "(Asterisco)")
    $enderecoRelatorio = $script:pathPastaMapeamento + "$prefixo`_$nomeProjeto`_Analisado.csv" 
    $dados | Export-Csv -Path $enderecoRelatorio -NoTypeInformation -Encoding Default
}
function validaListaImportacaoArquivos {
    param ($dados, $siglaConcessao, $poderConcedente, $pastaRaizProjeto, $cadastrosDisciplina, $cadastrosFaseProjeto)

    $nomesDuplicados = $dados |
    Group-Object -Property nome_arquivo |
    Where-Object { $_.Count -gt 1 } |
    Select-Object -ExpandProperty Name

    foreach ($arquivo in $dados) {
        $arquivo | Add-Member -NotePropertyName Observacoes -NotePropertyValue "" -Force
        $arquivo | Add-Member -NotePropertyName Status -NotePropertyValue "" -Force
        $arquivo | Add-Member -NotePropertyName NomeAjustado -NotePropertyValue "" -Force
        $arquivo | Add-Member -NotePropertyName PastaDestino -NotePropertyValue "" -Force

        if ($nomesDuplicados -contains $arquivo.nome_arquivo) {
            $arquivo.Status = "Erro - Duplicado"
            $arquivo.Observacoes = "Documento duplicado na lista de arquivos mapeados para importação"
        }
    }

    $preValidacao = $dados.Where( { $_.Status -eq ""} )

    foreach ($documento in $preValidacao) {
        $filename = $documento.nome_arquivo
        $filenameAjustado = ajustaNomeArquivo -filename $filename
        $documento.NomeAjustado = $filenameAjustado

        $documento = validaDocumento -documento $documento -siglaConcessao $siglaConcessao -poderConcedente $poderConcedente
        if($documento.Status -eq "Importar"){
            $pastaDestino = CalculaPastaDestino -documento $documento -poderConcedente $poderConcedente -pastaRaizProjeto $pastaRaizProjeto -cadastrosDisciplina $cadastrosDisciplina -cadastrosFaseProjeto $cadastrosFaseProjeto
            if(-not $pastaDestino){
                $documento.Status = "Erro"
                $documento.Observacoes = "Pasta de destino não localizada."                
            }
            else{
                $documento.PastaDestino = $pastaDestino
            }
        }
    }
    return $dados
}
function CalculaPastaDestino {
    param ($documento, $poderConcedente, $pastaRaizProjeto, $cadastrosDisciplina, $cadastrosFaseProjeto)

    if (-not $documento -or -not $pastaRaizProjeto) {
        return $nullO 
    }
    $pastaRaizProjeto = Get-PWFolderPathAndProperties -InputFolder $pastaRaizProjeto

    $disciplina = DefineDisciplinaPai -documento $documento -cadastrosDisciplina $cadastrosDisciplina -poderConcedente $poderConcedente
    if($null -eq $disciplina){
        return $null
    }
    $faseProjeto = DefineFaseProjeto -documento $documento -cadastrosFaseProjeto $cadastrosFaseProjeto -cadastrosDisciplina $cadastrosDisciplina -poderConcedente $poderConcedente
    if($null -eq $faseProjeto){
        return $null
    }
    
    $volume = ObtemVolumeFilename -filename $documento.NomeAjustado -poderConcedente $poderConcedente

    

    if($volume){
        $pastaDestino = $pastaRaizProjeto.FullPath + '\1 - Area de Trabalho\' + $faseProjeto + '\' + $volume + '\' + $disciplina
    }
    else{
        $pastaDestino = $pastaRaizProjeto.FullPath + '\1 - Area de Trabalho\' + $faseProjeto + '\' + $disciplina
    }

    $pasta = Get-PWFolders -FolderPath $pastaDestino
    if($pasta){
        return $pastaDestino
    }
    return $null
}
function ObtemListaCadastrosDisciplina {
    param ($poderConcedente)
    
    $tipoRegistro = switch ($poderConcedente) {
        "ARTESP" { 'Disciplinas ARTESP' }
        Default { 'Disciplinas ANTT'}
    }
    $cadastrosDisciplina = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{TipoRegistro = $tipoRegistro } -GetAttributes

    return $cadastrosDisciplina
}
function ObtemCadastroDisciplina {
    param ($disciplina, $cadastrosDisciplina)

    if($null -eq $disciplina -or $disciplina -eq ''){
        return $null
    }
    $cadastroDisciplina = $cadastrosDisciplina.Where({$_.Attributes[0].Codigo -eq $disciplina})
    if ($null -eq $cadastroDisciplina -or $null -eq $cadastroDisciplina.CustomAttributes -or $null -eq $cadastroDisciplina.CustomAttributes.Relacionamento -or $cadastroDisciplina.CustomAttributes.Relacionamento -eq '') {
        return $null
    }
    return $cadastroDisciplina
}
function ObterDisciplinaFileName {
    param ($filename, $poderConcedente)
    
    if($poderConcedente -eq 'ARTESP') {
        return $filename.Split('-')[4]
    }
    if($filename.Contains('+')){ return $filename.Split('-')[6] } else { return $filename.Split('-')[7] }
}
function DefineDisciplinaPai {
    param ($documento, $cadastrosDisciplina, $poderConcedente)

    $disciplina = ObterDisciplinaFileName -filename $documento.NomeAjustado -poderConcedente $poderConcedente
    $cadastroDisciplina = ObtemCadastroDisciplina -disciplina $disciplina -cadastrosDisciplina $cadastrosDisciplina
    if($null -eq $cadastroDisciplina ){ return $null }

    if($poderConcedente -eq "ARTESP"){
        $disciplinaPai = $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[1]
    }
    else{
        $disciplinaPai = $cadastroDisciplina.Attributes[0].Relacionamento
    }

    return $disciplinaPai
}
function ObtemListaCadastrosFaseProjeto {
    param ($poderConcedente)

    if($poderConcedente -eq "ARTESP"){ return $null }
    $cadastrosFaseProjeto = Get-PWDocumentsBySearch -Environment 'dmsRegistro' -Attributes @{TipoRegistro = 'Tipos de Projeto ANTT' } -GetAttributes

    return $cadastrosFaseProjeto
}
function ObtemCadastroFaseProjeto {
    param ($faseprojeto, $cadastrosFaseProjeto)

    if($null -eq $faseprojeto -or $faseProjeto -eq ''){
        return $null
    }
    $cadastroFaseProjeto = $cadastrosFaseProjeto.Where({$_.Attributes[0].Codigo -eq $faseProjeto});
    if($null -eq $cadastroFaseProjeto -or $null -eq $cadastroFaseProjeto.CustomAttributes -or $null -eq $cadastroFaseProjeto.CustomAttributes.Relacionamento -or $cadastroFaseProjeto.CustomAttributes.Relacionamento -eq ''){
        return $null;
    }
    return $cadastroFaseProjeto;
}
function ObterFaseProjetoFileName {
    param ($filename, $poderConcedente)
    
    if($poderConcedente -eq 'ARTESP') { return $null }
    return $filename.Split('-')[5]
}
function DefineFaseProjeto {
    param ($documento, $cadastrosFaseProjeto, $cadastrosDisciplina, $poderConcedente)

    if($poderConcedente -eq "ARTESP"){
        $disciplina = ObterDisciplinaFileName -filename $documento.NomeAjustado -poderConcedente $poderConcedente
        $cadastroDisciplina = ObtemCadastroDisciplina -disciplina $disciplina -cadastrosDisciplina $cadastrosDisciplina
        if($null -eq $cadastroDisciplina ){ return $null }

        $faseProjeto = $cadastroDisciplina.Attributes[0].Relacionamento.Split(';')[0]
        return $faseProjeto
    }
    $faseProjeto = ObterFaseProjetoFileName -filename $documento.NomeAjustado -poderConcedente $poderConcedente
    $cadastroFaseProjeto = ObtemCadastroFaseProjeto -faseprojeto $faseProjeto  -cadastrosFaseProjeto $cadastrosFaseProjeto -poderConcedente $poderConcedente
    if($null -eq $cadastroFaseProjeto){
        return $null
    }
    $faseProjeto = $cadastroFaseProjeto.Attributes[0].Relacionamento

    return $faseProjeto
}
function ObtemFaseProjetoFilename {
    param ($filename, $poderConcedente)
    
    if($poderConcedente -eq 'ARTESP') { return $null }

    if($filename.Contains('+')){ return $filename.Split('-')[4] } else { return $filename.Split('-')[5] }
}
function ObtemTipoDocumentoFilename {
    param ($filename, $poderConcedente)

    if($poderConcedente -eq 'ARTESP'){ return $filename.Split('-')[0] }
    if($filename.Contains('+')){ return $filename.Split('-')[5] } else { return $filename.Split('-')[6] }
}
function ObtemVolumeFilename {
    param ($filename, $poderConcedente)

    if($poderConcedente -eq 'ARTESP'){ return '' }
    $faseprojeto = ObtemFaseProjetoFilename -filename $filename -poderConcedente $poderConcedente
    if($faseProjeto -eq 'FUN'){ return '' }
    if($faseProjeto -eq 'ANT'){ 
        $tipoDocumento = ObtemTipoDocumentoFilename -filename $filename -poderConcedente $poderConcedente
        if($tipoDocumento -eq 'DE' -or $tipoDocumento -eq 'RT') { return 'Volume II' }
        return 'Volume I'
    }
    if($faseProjeto -eq 'EXE' -or $faseprojeto -eq 'EXO'){
        $disciplina = ObterDisciplinaFileName -filename $filename -poderConcedente $poderConcedente
        if($disciplina.Substring(0,1) -eq 'O') { return 'Volume III' }
        $tipoDocumento = ObtemTipoDocumentoFilename -filename $filename -poderConcedente $poderConcedente
        if($tipoDocumento -eq 'DE') { return 'Volume II' }
        if($disciplina.Substring(0,1) -eq 'V'){
            switch ($disciplina.Substring(1,1)) {
                '1' { return 'Volume I' }
                '2' { return 'Volume II' }
                '3' { return 'Volume III' }
                '4' { return 'Volume IV' }
                Default { return '' }
            }
        } 
        return 'Volume I'
    }  
    return ''
}
function validaDocumento {
    param ($documento, $siglaConcessao, $poderConcedente)

    $environment = defineEnvironment -poderConcedente $poderConcedente
    if(-not $environment){ 
        $documento.Observacoes = "Poder concedente não definido"
        $documento.Status = "Erro"

        return $documento
    }

    $erro = validaNomeArquivo -filename $documento.NomeAjustado -siglaConcessao $siglaConcessao -poderConcedente $poderConcedente
    if($erro){
        $documento.Observacoes = $erro
        $documento.Status = "Erro Taxonomia"        

        return $documento
    }

    $validacao = validaSeDocumentoJaImportadoNaEngenharia -filename $documento.NomeAjustado -environment $environment
    if(-not $validacao){
        $documento.Observacoes ="Documento já importado na engenharia"
        $documento.Status = "Já importado"

        return $documento
    }    

    $documento.Status = "Importar"
    return $documento
}
function ajustaNomeArquivo {
    param ($filename)
    
    $filename = $filename.Replace("_r", "-R")
    $filename = $filename.Replace("_R", "-R")
    return $filename
}
function validaNomeArquivo {
    param ($filename, $siglaConcessao, $poderConcedente)
    
    if($filename.Substring(0,4) -eq "GRD-"){ return '' }

    if($poderConcedente -eq 'ARTESP'){ $funcao = "[ValidaNumeroARTESP]"}
    else { $funcao = "[ValidaNumeroANTT]"}

    $query = "select [dbo].$funcao('$filename', '$siglaConcessao')"
    $resultado = Select-PWSQL -SQLSelectStatement $query
    if($resultado[0].Column1.Substring(0,5) -eq 'Erro:'){
        return $resultado[0].Column1.Substring(6)
    }
    return ''
}
function validaSeDocumentoDuplicadoNaListaDeImportacao {
    param ($filename, $listaDocumentos)
    
    $documentos = $listaDocumentos.Where({$_.nome_arquivo -eq $filename})
    if($documentos.Count -gt 1) { return $false }
    return $true
}
function validaSeDocumentoJaImportadoNaAreaDeTransferencia {
    param ($filename, $pastaDestino)
    
    $documentos = Get-PWDocumentsBySearch -FileName $filename -FolderPath $pastaDestino

    if($documentos) { return $false }
    return $true
}
function validaSeDocumentoJaImportadoNaEngenharia {
    param ($filename, $environment)

    $documentos = Get-PWDocumentsBySearch -DocumentName $filename -Environment $environment

    if($documentos) { return $false }
    return $true
}
function defineEnvironment {
    param ($poderConcedente)
    
    if($poderConcedente -eq "ARTESP"){ return "dmsEngenhariaARTESP"}
    if($poderConcedente -eq "ANTT"){ return "dmsEngenhariaANTT"}
    return ''
}
function ObterNumeroDocumentoFilename {
    param ($filename)
    
    return ($filename -Split "-R")[0]
}
function ObterRevisaoFilename {
    param ($filename)
    
    $parteRevisao = ($filename -split "-R")[1]
    return $parteRevisao.Substring(0,2)
}
function ObterVersaoFilename {
    param ($filename)

    $parteRevisao = ($filename -split "-R")[1]
    return $parteRevisao.Substring(2,1)   
}
function AjustaEmissoesERevisaPastaDestino {
    param ($documento)

    $numeroDocumento = ObterNumeroDocumentoFilename -filename $documento.NomeAjustado
    $revisao = ObterRevisaoFilename -filename $documento.NomeAjustado
    $versao = ObterVersaoFilename -filename $documento.NomeAjustado

    $numeroDocumentoPesquisa = $numeroDocumento+"%"
    $documentos = Get-PWDocumentsBySearch -FolderPath $documento.PastaDestino -FileName $numeroDocumentoPesquisa -PopulatePath -GetAttributes -JustThisFolder

    if(-not $documentos){
        return $documento.PastaDestino
    }
    
    $pastaSuperados =  $documento.PastaDestino+"\Superados"

    $emissoesSuperior = $documentos.Where({$revisao -lt $_.CustomAttributes.Revisao -or $revisao -eq $_.CustomAttributes.Revisao -and $versao -lt $_.CustomAttributes.Versao })
    if($emissoesSuperior){ return $pastaSuperados }

    $mesmaEmissao = $documentos.Where({$revisao -eq $_.CustomAttributes.Revisao -and $versao -eq $_.CustomAttributes.Versao })
    if($mesmaEmissao){ return $documento.PastaDestino }

    $emissoesAnterior = $documentos.Where({ $revisao -gt $_.CustomAttributes.Revisao -or $revisao -eq $_.CustomAttributes.Revisao -and $versao -gt $_.CustomAttributes.Versao})
    
    $documentosMovimentados = Move-PWDocumentsToFolder -InputDocument $emissoesAnterior -TargetFolderPath $pastaSuperados
    Set-PWDocumentState -InputDocuments $documentosMovimentados -State "Superado" -Force

    return $documento.PastaDestino
}
function importarDocumentosPW {
    param ($dadosAnalisados, $projeto, $prefixo)
 
    $nomeProjeto = $projeto.Replace("*", "(Asterisco)")
    $enderecoRelatorio = $script:pathPastaMapeamento + "$prefixo`_$nomeProjeto`_Importado.csv"

    $documentosValidos = $dadosAnalisados.Where({$_.Status -eq "Importar"})

    foreach ($documento in $documentosValidos) {        
        try {
            $pastaDestinoRevisado = AjustaEmissoesERevisaPastaDestino -documento $documento

            if($documento.full_path.Length -ge 260){
                if(-not (Test-Path "c:\Temp")){ New-Item -ItemType Directory -Path "C:\Temp" }
                $arquivoOrigem = "\\?\"+$documento.full_path    
                $nomeArquivoOrigem = Split-Path -Path $arquivoOrigem -Leaf
                Copy-Item -Path $arquivoOrigem -Destination "C:\Temp"
                $arquivoTemporario = "C:\Temp\"+$nomeArquivoOrigem
                $documentoImportado = New-PWDocument -FolderPath $pastaDestinoRevisado -DocumentName $documento.NomeAjustado -FilePath $arquivoTemporario
                $null = Update-PWDocumentAttributes -InputDocuments $documentoImportado -Attributes @{Label =" "}    
                Remove-Item -Path $arquivoTemporario
                $documento | Add-Member -MemberType NoteProperty -Name "Data Importacao" -Value (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Write-Host "Importado arquivo: " + $documento.NomeAjustado
            }else{
                $documentoImportado = New-PWDocument -FolderPath $pastaDestinoRevisado -DocumentName $documento.NomeAjustado -FilePath $documento.full_path
                $null = Update-PWDocumentAttributes -InputDocuments $documentoImportado -Attributes @{Label =" "}    
                $documento | Add-Member -MemberType NoteProperty -Name "Data Importacao" -Value (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Write-Host "Importado arquivo: " + $documento.NomeAjustado
            }
        }
        catch {
            $documento | Add-Member -MemberType NoteProperty -Name "Data Importacao" -Value "Erro na importacao"
            Write-Host "Erro ao importar arquivo: " + $documento.NomeAjustado
        }
        
        $documento | Export-Csv -Path $enderecoRelatorio -NoTypeInformation -Encoding Default -Append        
    }

}

#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------

$login = New-PWLogin

if (-not $login) {
    return
}

SelecionarPlanilhaCadastrosDeProjetos

Undo-PWLogin