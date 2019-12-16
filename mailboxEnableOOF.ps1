$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '500,648'
$Form.text                       = "Enable Out of Office"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "Out of Office Message"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(262,49)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.TextBox
$ListBoxUserGainingAccess.multiline   = $true
$ListBoxUserGainingAccess.width  = 216
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(262,81)

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
$mailbox = $ListBoxMailboxToManage.SelectedItem
$oofMessage = $ListBoxUserGainingAccess.Text

$ProgressBar1.Value = 25
$LabelStatus.Text = "Status: Setting out of office message"

##Set-MailboxAutoReplyConfiguration -identity $y -AutoReplyState Enabled -InternalMessage $x  -ExternalMessage $x
$AutoReplyStatus = Get-MailboxAutoReplyConfiguration -Identity "$mailbox"

if ($AutoReplyStatus.AutoReplyState -eq "Enabled") {
    Add-Content "$logfile" "========OOF STATUS========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Out-of-Office is already Enabled"
    $LabelStatus.Text = "Status: Out-of-Office is already Enabled"
}
elseif ($AutoReplyStatus.AutoReplyState -eq "Disabled") {
    Set-MailboxAutoReplyConfiguration -identity $mailbox -AutoReplyState Enabled -InternalMessage $oofMessage  -ExternalMessage $oofMessage
    Add-Content "$logfile" "========SUCCESS========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Set Out-of-Office Message"
    $LabelStatus.Text = "Status: Set Out-of-Office Message completed"

}
elseif ($AutoReplyStatus.AutoReplyState -eq "Scheduled") {
    Add-Content "$logfile" "========OOF STATUS========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Out-of-Office is already Scheduled"
    $LabelStatus.Text = "Status: Out-of-Office is already Scheduled"
}
else{
    Add-Content "$logfile" "FAILED OOF STATUS"
    $LabelStatus.Text = "Status: FAILED OOF STATUS"
}
    
    
   
$ProgressBar1.Value = 100

})


$Form.ShowDialog()