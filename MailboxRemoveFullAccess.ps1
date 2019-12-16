$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '500,648'
$Form.text                       = "Remove User Full Access to 1 Mailbox"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User being Removed"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(262,49)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User being Removed"
$ListBoxUserGainingAccess.width  = 216
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(262,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$CheckBoxAddToOutlook            = New-Object system.Windows.Forms.CheckBox
$CheckBoxAddToOutlook.text       = "Add Mailbox to Outlook?"
$CheckBoxAddToOutlook.AutoSize   = $false
$CheckBoxAddToOutlook.width      = 194
$CheckBoxAddToOutlook.height     = 20
$CheckBoxAddToOutlook.location   = New-Object System.Drawing.Point(19,472)
$CheckBoxAddToOutlook.Font       = 'Microsoft Sans Serif,10,style=Bold'

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

$CheckBoxGiveSendAs              = New-Object system.Windows.Forms.CheckBox
$CheckBoxGiveSendAs.text         = "Give Send As Permission?"
$CheckBoxGiveSendAs.AutoSize     = $false
$CheckBoxGiveSendAs.width        = 202
$CheckBoxGiveSendAs.height       = 20
$CheckBoxGiveSendAs.location     = New-Object System.Drawing.Point(262,472)
$CheckBoxGiveSendAs.Font         = 'Microsoft Sans Serif,10,style=Bold'

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage))

$ButtonRunScript.Add_Click({

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Starting..."

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem



    $accessRight = "FullAccess"
 
    $mailboxes = Get-mailbox
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
   
        Write-Host "$userRequiringAccess being removed from $mailbox" -ForegroundColor Green
        Add-Content "$logfile" "$userRequiringAccess being removed from $mailbox"
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess being removed from $mailbox"
        
        Remove-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -Confirm:$false


        $ProgressBar1.Value = 75
        $LabelStatus.Text = "$userRequiringAccess being removed from $mailbox SendAs"
        Write-Host "$userRequiringAccess being removed from $mailbox SendAs" -ForegroundColor Green
        Add-Content "$logfile" "$userRequiringAccess being removed from $mailbox SendAs"
        Remove-RecipientPermission "$mailbox" -AccessRights SendAs -Trustee "$userRequiringAccess" -Confirm:$false




##############




$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
})


$Form.ShowDialog()