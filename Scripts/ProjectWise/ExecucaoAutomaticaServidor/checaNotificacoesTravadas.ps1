function obterNotificacoesEmExecucao {
    $documentos = Get-PWDocumentsBySearch -SearchName 'Scripts\Notificacoes Em Execucao' -GetAttributes
    return $documentos
}
function obterNotificacoesTravadas {
    param ($notificacoes)

    $data = (Get-Date).AddHours(-1).ToString('yyyy-MM-dd HH:mm')

    $notificacoesTravadas = $notificacoes.Where({$_.CustomAttributes.LASTRUN_TIME -le $data })
    return $notificacoesTravadas
}
function destravarNotificacoes {
    param ($notificacoesTravadas)
    
    $null = Update-PWDocumentAttributes -InputDocuments $notificacoesTravadas -Attributes @{ NOTIFICATION_IS_RUNNING = 0 }
}

#-------------------------------------------------------
#Início da execução
#-------------------------------------------------------
$SecurePassword = ConvertTo-SecureString '123456' -AsPlainText -Force
New-PWLogin -DatasourceName '01SSRV305.ECSC.ECORODOVIAS.CORP:ecorodovias-pw-01' -Password $SecurePassword -UserName 'admin'

$notificacoes = obterNotificacoesEmExecucao

$notificacoesTravadas = obterNotificacoesTravadas -notificacoes $notificacoes

destravarNotificacoes -notificacoesTravadas $notificacoesTravadas

Undo-PWLogin