$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '500,648'
$Form.text                       = "Forwarding Emails"
$Form.TopMost                    = $false

$LabelForwardToMailbox           = New-Object system.Windows.Forms.Label
$LabelForwardToMailbox.text      = "Mailbox to Forward to"
$LabelForwardToMailbox.AutoSize  = $true
$LabelForwardToMailbox.width     = 25
$LabelForwardToMailbox.height    = 10
$LabelForwardToMailbox.location  = New-Object System.Drawing.Point(262,49)
$LabelForwardToMailbox.Font      = 'Microsoft Sans Serif,10,style=Bold'
$LabelForwardToMailbox.ForeColor  = "#000000"

$ListBoxForwardToMailbox         = New-Object system.Windows.Forms.ListBox
$ListBoxForwardToMailbox.text    = "Mailbox to Forward to"
$ListBoxForwardToMailbox.width   = 216
$ListBoxForwardToMailbox.height  = 362
$ListBoxForwardToMailbox.location  = New-Object System.Drawing.Point(262,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxForwardToMailbox.Items.Add($mailbox.PrimarySmtpAddress)}

$CheckBoxKeepCopy                = New-Object system.Windows.Forms.CheckBox
$CheckBoxKeepCopy.text           = "Keep Copy in Original Mailbox?"
$CheckBoxKeepCopy.AutoSize       = $false
$CheckBoxKeepCopy.width          = 250
$CheckBoxKeepCopy.height         = 20
$CheckBoxKeepCopy.location       = New-Object System.Drawing.Point(19,472)
$CheckBoxKeepCopy.Font           = 'Microsoft Sans Serif,10,style=Bold'

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


$Form.controls.AddRange(@($LabelForwardToMailbox,$ListBoxForwardToMailbox,$CheckBoxKeepCopy,$ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage))

$ButtonRunScript.Add_Click({
    
   
    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Starting..."
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxForwardToMailbox.SelectedItem

if ($CheckBoxKeepCopy.Checked) {

    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Setting up Email Forwarding for $mailbox "
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    Write-Host "Forwarding Emails from $mailbox to $userRequiringAccess and keeping a copy in original mailbox. " -foregroundColor Yellow
    Add-Content "$logfile" "Forwarding Emails from $mailbox to $userRequiringAccess and keeping a copy in original mailbox."

    Set-Mailbox -Identity $mailbox -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $userRequiringAccess

}
else {

    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Setting up Email Forwarding for $mailbox "
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    Write-Host "Forwarding Emails from $mailbox to $userRequiringAccess. NOT keeping copy in original mailbox." -foregroundColor Yellow
    Add-Content "$logfile" "Forwarding Emails from $mailbox to $userRequiringAccess. NOT keeping copy in original mailbox."

    Set-Mailbox -Identity $mailbox -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $userRequiringAccess
}




$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
$ProgressBar1.Refresh()
$LabelStatus.Refresh()
})


$Form.ShowDialog()