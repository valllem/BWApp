$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '500,648'
$Form.text                       = "Grant User Send As to a Mailbox"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User Gaining Access"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(262,49)
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
$ButtonRunScript.width           = 397
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 398
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

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Status: Starting..."

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
            
    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Status: Adding Send As Permissions"
    Write-Host "Adding $userRequiringAccess SendAs on $mailbox" -foregroundColor Cyan
    Add-Content "$logfile" "Adding $userRequiringAccess SendAs on $mailbox"
try {
    Add-RecipientPermission "$mailbox" -AccessRights SendAs -Trustee "$userRequiringAccess" -Confirm:$false
    Write-Host "Added Permissions" -foregroundColor Green
    }
catch {
    Write-Host "Could not add SendAs permissions" -foregroundColor Red
    Add-Content "$logfile" "Could not add SendAs permissions for $userRequiringAccess on $mailbox"
}
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Completed Task"

 

    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Completed Task"
})


$Form.ShowDialog()