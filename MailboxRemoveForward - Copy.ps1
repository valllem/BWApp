$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '350,648'
$Form.text                       = "Remove Forwarding on Mailbox..."
$Form.TopMost                    = $false

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 220
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(60,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 220
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(60,546)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(60,616)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$ListBoxMailboxToManage          = New-Object system.Windows.Forms.ListBox
$ListBoxMailboxToManage.text     = "Mailbox to Manage"
$ListBoxMailboxToManage.width    = 220
$ListBoxMailboxToManage.height   = 362
$ListBoxMailboxToManage.location  = New-Object System.Drawing.Point(60,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxMailboxToManage.Items.Add($mailbox.PrimarySmtpAddress)}

$LabelMailboxToManage            = New-Object system.Windows.Forms.Label
$LabelMailboxToManage.text       = "Mailbox to disable forwarding"
$LabelMailboxToManage.AutoSize   = $true
$LabelMailboxToManage.width      = 25
$LabelMailboxToManage.height     = 10
$LabelMailboxToManage.location   = New-Object System.Drawing.Point(60,49)
$LabelMailboxToManage.Font       = 'Microsoft Sans Serif,10,style=Bold'




$Form.controls.AddRange(@($ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage))

$ButtonRunScript.Add_Click({
    
   
    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Starting..."
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxForwardToMailbox.SelectedItem



    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Removing forwarding on $mailbox "
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    Write-Host "Disabling forwarding on $mailbox" -foregroundColor Yellow
    Add-Content "$logfile" "Disabling forwarding on $mailbox"

    Set-Mailbox $mailbox -ForwardingAddress $NULL -ForwardingSmtpAddress $NULL





    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Completed Tasks, see log file for info"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
})


$Form.ShowDialog()