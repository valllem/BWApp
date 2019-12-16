$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,648'
$Form.text                       = "Grant User Access to ALL Calendars"
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

$LabelPermission                 = New-Object system.Windows.Forms.Label
$LabelPermission.text            = "Permission"
$LabelPermission.AutoSize        = $true
$LabelPermission.width           = 25
$LabelPermission.height          = 10
$LabelPermission.location        = New-Object System.Drawing.Point(360,47)
$LabelPermission.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ListBoxPermission               = New-Object system.Windows.Forms.ListBox
$ListBoxPermission.text          = "Permission"
$ListBoxPermission.width         = 199
$ListBoxPermission.height        = 167
$ListBoxPermission.location      = New-Object System.Drawing.Point(360,83)
@('AvailabilityOnly','LimitedDetails','Author','PublishingEditor','Owner') | ForEach-Object {[void] $ListBoxPermission.Items.Add($_)}

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 480
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(44,509)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'
#$Form.Controls.Add($ButtonRunScript)

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 480
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(45,545)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(19,616)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($LabelUserGainingAccess,$ListBoxUserGainingAccess,$ButtonRunScript,$ProgressBar1,$LabelStatus,$LabelPermission,$ListBoxPermission))




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
    
Start-Sleep -Milliseconds 300
     
    if ($accessRights.accessRights -notmatch $accessRight -and $mailbox.primarysmtpaddress -notcontains $userRequiringAccess -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess
        if ($accessRights.accessRights -match $accessRight) {
            Write-Host "Successfully added $accessRight permissions on $($mailbox.primarysmtpaddress)'s calendar for $userrequiringaccess" -ForegroundColor Green
            Add-Content "$Logfile" "Successfully added $accessRight permissions on $($mailbox.primarysmtpaddress)'s calendar for $userrequiringaccess"
        }
        else {
            Write-Host "Could not add $accessRight permissions on $($mailbox.primarysmtpaddress)'s calendar for $userrequiringaccess" -ForegroundColor Red
            Add-Content "Could not add $accessRight permissions on $($mailbox.primarysmtpaddress)'s calendar for $userrequiringaccess"
        }
    $ProgressBar1.Value = 80
    }else{
        Write-Host "Permission level already exists for $($userrequiringaccess.primarysmtpaddress) on $($mailbox.primarysmtpaddress)'s calendar" -foregroundColor Green
        Add-Content "Permission level already exists for $($userrequiringaccess.primarysmtpaddress) on $($mailbox.primarysmtpaddress)'s calendar"
    }
    $ProgressBar1.Value = 80
    $LabelStatus.Text = "Updating: $mailbox"
}

$ProgressBar1.Value = 100
$LabelStatus.Text = "Completed Tasks, see log file for info"
})


$Form.ShowDialog()