function SelectProjectList {
    Add-Type -AssemblyName System.Windows.Forms
    $owner=New-Object Windows.Forms.Form
    $dialog=New-Object Windows.Forms.OpenFileDialog
    try {
        $owner.TopMost=$true
        $dialog.Title='Selecione a planilha de projetos SharePoint'
        $dialog.Filter='Planilha Excel (*.xlsx)|*.xlsx'
        $dialog.CheckFileExists=$true
        $dialog.Multiselect=$false
        if($dialog.ShowDialog($owner) -eq [Windows.Forms.DialogResult]::OK){return $dialog.FileName}
        return $null
    }finally{$dialog.Dispose();$owner.Dispose()}
}
