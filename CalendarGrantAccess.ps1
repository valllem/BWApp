$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
Connect-ExchangeOnline
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '756,648'
$Form.text                       = "Grant Access to a Calendar"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User Gaining Access"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(263,48)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User Gaining Access"
$ListBoxUserGainingAccess.width  = 216
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(262,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 655
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 658
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

$LabelPermission                 = New-Object system.Windows.Forms.Label
$LabelPermission.text            = "Permission"
$LabelPermission.AutoSize        = $true
$LabelPermission.width           = 25
$LabelPermission.height          = 10
$LabelPermission.location        = New-Object System.Drawing.Point(505,49)
$LabelPermission.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ListBoxPermission               = New-Object system.Windows.Forms.ListBox
$ListBoxPermission.text          = "Permission"
$ListBoxPermission.width         = 199
$ListBoxPermission.height        = 167
$ListBoxPermission.location      = New-Object System.Drawing.Point(505,81)
@('AvailabilityOnly','LimitedDetails','Author','PublishingEditor','Owner') | ForEach-Object {[void] $ListBoxPermission.Items.Add($_)}


$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage,$LabelPermission,$ListBoxPermission))


$ButtonRunScript.Add_Click({
    
    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
    $permlevel = $ListBoxPermission.SelectedItem

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Status: Setting Permissions"

    $permission = Get-MailboxFolderPermission -identity $mailbox':\calendar' -User "$userRequiringAccess"
if($permission -eq $null){
    Add-MailboxFolderPermission $mailbox':\calendar' -User "$userRequiringAccess" -AccessRights $permlevel -ErrorAction SilentlyContinue
    Write-Host "Gave $permlevel access to $userRequiringAccess on $mailbox's Calendar." -foregroundColor Green
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Gave $permlevel access to $userRequiringAccess on $mailbox's Calendar."
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Gave $permlevel access to $userRequiringAccess on $mailbox's Calendar."

}else{
    Set-MailboxFolderPermission $mailbox':\calendar' -User "$userRequiringAccess" -AccessRights $permlevel -ErrorAction SilentlyContinue
    Write-Host "$userRequiringAccess already had access, so the permissions were updated to $permlevel" -foregroundColor Yellow
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "$userRequiringAccess already had access, so the permissions were updated to $permlevel"
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Permissions already exist, replacing."
}


$ProgressBar1.Value = 100
$LabelStatus.Text = "Status: Tasks Complete."

Disconnect-ExchangeOnline -Confirm:$false
})



$Form.ShowDialog()