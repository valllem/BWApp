$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '508,648'
$Form.text                       = "Remove Access to a Calendar"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User being removed"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(263,48)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User being removed"
$ListBoxUserGainingAccess.width  = 216
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(262,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 406
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 406
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(44,546)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(19,616)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$ListBoxMailboxToManage          = New-Object system.Windows.Forms.ListBox
$ListBoxMailboxToManage.text     = "Mailbox to Manage"
$ListBoxMailboxToManage.width    = 218
$ListBoxMailboxToManage.height   = 362
$ListBoxMailboxToManage.location  = New-Object System.Drawing.Point(19,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxMailboxToManage.Items.Add($mailbox.PrimarySmtpAddress)}

$LabelMailboxToManage            = New-Object system.Windows.Forms.Label
$LabelMailboxToManage.text       = "Mailbox to Manage"
$LabelMailboxToManage.AutoSize   = $true
$LabelMailboxToManage.width      = 25
$LabelMailboxToManage.height     = 10
$LabelMailboxToManage.location   = New-Object System.Drawing.Point(19,49)
$LabelMailboxToManage.Font       = 'Microsoft Sans Serif,10,style=Bold'

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage))


$ButtonRunScript.Add_Click({
    
    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
    $permlevel = $ListBoxPermission.SelectedItem

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Status: Setting Permissions"

    $permission = Get-MailboxFolderPermission -identity $mailbox':\calendar' -User "$userRequiringAccess"
if($permission -eq $null){
    
    Write-Host "User does not have access to this calendar" -foregroundColor yellow
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "$userRequiringAccess does not have access to $mailbox's Calendar."
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: $userRequiringAccess does not have access to $mailbox's Calendar."

}else{
    Remove-MailboxFolderPermission -Identity $mailbox':\calendar' -User "$userRequiringAccess" -confirm:$false
    Write-Host "$userRequiringAccess has been removed access" -foregroundColor Yellow
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "$userRequiringAccess has been removed access"
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Permissions exist, removing."
}


$ProgressBar1.Value = 100
$LabelStatus.Text = "Status: Tasks Complete."


})



$Form.ShowDialog()