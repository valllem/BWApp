$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '370,648'
$Form.text                       = "Remove user from All Calendars"
$Form.TopMost                    = $false

$LabelUserGainingAccess          = New-Object system.Windows.Forms.Label
$LabelUserGainingAccess.text     = "User being removed"
$LabelUserGainingAccess.AutoSize  = $true
$LabelUserGainingAccess.width    = 25
$LabelUserGainingAccess.height   = 10
$LabelUserGainingAccess.location  = New-Object System.Drawing.Point(19,47)
$LabelUserGainingAccess.Font     = 'Microsoft Sans Serif,10,style=Bold'
$LabelUserGainingAccess.ForeColor  = "#000000"

$ListBoxUserGainingAccess        = New-Object system.Windows.Forms.ListBox
$ListBoxUserGainingAccess.text   = "User being removed"
$ListBoxUserGainingAccess.width  = 329
$ListBoxUserGainingAccess.height  = 362
$ListBoxUserGainingAccess.location  = New-Object System.Drawing.Point(19,83)
foreach ($mailbox in $mailboxes) {[void] $ListBoxUserGainingAccess.Items.Add($mailbox.PrimarySmtpAddress)}

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 275
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'
#$Form.Controls.Add($ButtonRunScript)

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 275
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(44,545)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(19,616)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$ButtonRunScript,$ProgressBar1,$LabelStatus))




$ButtonRunScript.Add_Click({

$LabelUserGainingAccess



$ProgressBar1.Value = 25
$LabelStatus.Text = "Starting..."

$userRequiringAccess = $ListBoxUserGainingAccess.SelectedItem
$accessRight = $ListBoxPermission.SelectedItem
 
$mailboxes = Get-mailbox

foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess -erroraction SilentlyContinue
    
    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Updating: $mailbox"
    
    
    
if ($accessRights.accessRights -ne $null) {

    Start-Sleep -Milliseconds 300
    Write-Host -ForegroundColor Yellow "Removing $userRequiringAccess from $mailbox"
    Add-Content "$Logfile" "Removing $userRequiringAccess from $mailbox"
    Remove-MailboxFolderPermission -Identity $mailbox':\calendar' -User "$userRequiringAccess" -confirm:$false
    
    Start-Sleep -Milliseconds 300

    #check removal has applied
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess
    if ($accessRights.accessRights -eq $null) { 
        Start-Sleep -Milliseconds 300
        Write-Host -ForegroundColor Green "Removed $userRequiringAccess from $mailbox"
        Add-Content "$Logfile" "Removed $userRequiringAccess from $mailbox"
    
    
    }#close if
    else {
        Write-Host -ForegroundColor Red "ERROR removing $userRequiringAccess from $mailbox"
        Add-Content "$Logfile" "ERROR removing $userRequiringAccess from $mailbox"

    }#close else
    $ProgressBar1.Value = 80
    Start-Sleep -milliseconds 500

    }#close if
    

else { 
    Start-Sleep -Milliseconds 300
    Write-Host -ForegroundColor Green "No permissions exist for $userRequiringAccess on $mailbox, skipping"
    Add-Content "$Logfile" "No permissions exist for $userRequiringAccess on $mailbox, skipping"


}#close else

}#close foreach

$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
})


$Form.ShowDialog()