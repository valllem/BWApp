$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '500,648'
$Form.text                       = "Grant User Full Access to 1 Mailbox"
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

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$CheckBoxAddToOutlook,$ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage,$CheckBoxGiveSendAs))

$ButtonRunScript.Add_Click({

    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Starting..."

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

if ($CheckBoxAddToOutlook.Checked -and $CheckBoxGiveSendAs.Checked) {

    $accessRight = "FullAccess"
 
    $mailboxes = Get-mailbox
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "Automapping ON, Full Access and SendAs" -foregroundColor Cyan
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green

        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Automapping ON, Full Access and SendAs" -foregroundColor Cyan
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
        Add-Content "$logfile" "Adding $userRequiringAccess SendAs permission on $mailbox"

        $ProgressBar1.Value = 50
        $LabelStatus.Text = "Adding $userRequiringAccess $accessRight on $mailbox"

        Add-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -AutoMapping $true | Out-File -Encoding Ascii -append "$logfile"
        Add-RecipientPermission "$mailbox" -AccessRights SendAs -Trustee "$userRequiringAccess" -Confirm:$false




    } #close elseif
else {

        Write-Host "No task completed" -foregroundColor red
        Add-Content "$logfile" "No task completed"
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "No task completed, see log file"
} #close else



} #Add to Outlook with SendAs
elseif ($CheckBoxAddToOutlook.Checked){

    $accessRight = "FullAccess"
 
    $mailboxes = Get-mailbox
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "Automapping ON, Full Access Only" -foregroundColor Cyan
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green

        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Automapping ON, Full Access" -foregroundColor Cyan
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
        Add-Content "$logfile" "Adding $userRequiringAccess SendAs permission on $mailbox"

        $ProgressBar1.Value = 50
        $LabelStatus.Text = "Adding $userRequiringAccess $accessRight on $mailbox"
        
        Add-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -AutoMapping $true | Out-File -Encoding Ascii -append "$logfile"
        
    } #close elseif
else {

        Write-Host "No task completed" -foregroundColor red
        Add-Content "$logfile" "No task completed"
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "No task completed, see log file"
} #close else
} #Add to Outlook
elseif ($CheckBoxAddToOutlook -and $CheckBoxGiveSendAs.Checked){

    $accessRight = "FullAccess"
 
    $mailboxes = Get-mailbox
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "Automapping Off, FullAccess and SendAs" -ForegroundColor Cyan
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green
        
        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Automapping Off, FullAccess and SendAs" -ForegroundColor Cyan
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
        Add-Content "$logfile" "Adding $userRequiringAccess SendAs permission on $mailbox"

        $ProgressBar1.Value = 50
        $LabelStatus.Text = "Adding $userRequiringAccess $accessRight on $mailbox"

        Add-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -AutoMapping $false | Out-File -Encoding Ascii -append "$logfile"
        Add-RecipientPermission "$mailbox" -AccessRights SendAs -Trustee "$userRequiringAccess" -Confirm:$false

    } #close elseif
else {

        Write-Host "No task completed" -foregroundColor red
        Add-Content "$logfile" "No task completed"
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "No task completed, see log file"
} #close else
} #Webmail only with SendAs
elseif ($CheckBoxAddToOutlook){

    $accessRight = "FullAccess"
    
    $mailboxes = Get-mailbox
    $userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem

    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "Automapping Off, FullAccess Only" -ForegroundColor Cyan
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green
        
        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Automapping Off, FullAccess Only" -ForegroundColor Cyan
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
        Add-Content "$logfile" "Adding $userRequiringAccess SendAs permission on $mailbox"

        $ProgressBar1.Value = 50
        $LabelStatus.Text = "Adding $userRequiringAccess $accessRight on $mailbox"

        Add-MailboxPermission -Identity "$mailbox" -User "$userRequiringAccess" -AccessRights FullAccess -AutoMapping $false | Out-File -Encoding Ascii -append "$logfile"
        
    } #close elseif
else {

        Write-Host "No task completed" -foregroundColor red
        Add-Content "$logfile" "No task completed"
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "No task completed, see log file"
} #close else
} #Webmail only




##############




$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
})


$Form.ShowDialog()