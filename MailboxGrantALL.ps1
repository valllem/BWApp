$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '374,648'
$Form.text                       = "Grant User Full Access to ALL Mailboxes"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User Gaining Access"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(19,47)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User Gaining Access"
$ListBoxUserGainingAccess.width  = 329
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(19,83)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$CheckBoxAddToOutlook            = New-Object system.Windows.Forms.CheckBox
$CheckBoxAddToOutlook.text       = "Add Mailboxes to Outlook? (Not Recommended)"
$CheckBoxAddToOutlook.AutoSize   = $false
$CheckBoxAddToOutlook.width      = 333
$CheckBoxAddToOutlook.height     = 20
$CheckBoxAddToOutlook.location   = New-Object System.Drawing.Point(19,472)
$CheckBoxAddToOutlook.Font       = 'Microsoft Sans Serif,10,style=Bold'

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


if ($CheckBoxAddToOutlook.Checked) {

$accessRight = "FullAccess"
 
$mailboxes = Get-mailbox
$userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green
        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
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


} #close foreach

} #close if.checked
else{

$accessRight = "FullAccess"
 
$mailboxes = Get-mailbox
$userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = " "

     
    if ($accessRights.accessRights -eq "FullAccess") {
        Write-Host "$userRequiringAccess already has $accessRight on $mailbox" -ForegroundColor Green
        Add-Content "$logfile" "$userRequiringAccess already has $accessRight on $mailbox"
        
        $ProgressBar1.Value = 50
        $LabelStatus.Text = "$userRequiringAccess already has $accessRight on $mailbox"
       
    } #close if
elseif($accessRights.accessRights -ne "FullAccess"){
        Write-Host "Adding $userRequiringAccess $accessRight on $mailbox" -foregroundColor Yellow
        Add-Content "$logfile" "Adding $userRequiringAccess $accessRight on $mailbox"
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


} #close foreach



} #close else (if not checked)




##############




$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
})


$Form.ShowDialog()