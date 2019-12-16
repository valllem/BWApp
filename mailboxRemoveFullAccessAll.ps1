$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '374,648'
$Form.text                       = "Remove User Full Access to ALL Mailboxes"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User to be removed"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(19,47)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User to be removed"
$ListBoxUserGainingAccess.width  = 329
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(19,83)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 273
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'
#$Form.Controls.Add($ButtonRunScript)

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 272
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(45,545)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(19,616)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$CheckBoxAddToOutlook,$ButtonRunScript,$ProgressBar1,$LabelStatus))




$ButtonRunScript.Add_Click({

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Starting..."


$accessRight = "FullAccess"
 
$mailboxes = Get-mailbox
$userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "
    Start-Sleep -Milliseconds 300
 
    Write-Host "Removing $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
    Add-Content "$logfile" "Removing $userRequiringAccess $accessRight on $mailbox"
    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Removing $userRequiringAccess $accessRight on $mailbox"
    Remove-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue


} #close foreach


$ProgressBar1.Value = 100
$LabelStatus.Text = "Status: Completed Tasks, see log file for info"
})


$Form.ShowDialog()