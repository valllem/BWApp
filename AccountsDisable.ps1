$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
$FilePath = "C:\BWApp\logs"
$colourAlert = "Yellow"
$colourInfo = "Yellow"
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '350,648'
$Form.text                       = "Check Mailbox Permissions"
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
$ListBoxMailboxToManage.text     = "Mailbox to Check"
$ListBoxMailboxToManage.width    = 220
$ListBoxMailboxToManage.height   = 362
$ListBoxMailboxToManage.location  = New-Object System.Drawing.Point(60,81)
foreach ($mailbox in $mailboxes) {[void] $ListBoxMailboxToManage.Items.Add($mailbox.PrimarySmtpAddress)}

$LabelMailboxToManage            = New-Object system.Windows.Forms.Label
$LabelMailboxToManage.text       = "Mailbox to check permissions"
$LabelMailboxToManage.AutoSize   = $true
$LabelMailboxToManage.width      = 25
$LabelMailboxToManage.height     = 10
$LabelMailboxToManage.location   = New-Object System.Drawing.Point(60,49)
$LabelMailboxToManage.Font       = 'Microsoft Sans Serif,10,style=Bold'




$Form.controls.AddRange(@($ButtonRunScript,$ProgressBar1,$LabelStatus,$ListBoxMailboxToManage,$LabelMailboxToManage))

$ButtonRunScript.Add_Click({
$x = $ListBoxMailboxToManage.SelectedItem
   
    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Preparing..."
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    $mailbox = $ListBoxMailboxToManage.SelectedItem
    $userRequiringAccess = $ListBoxForwardToMailbox.SelectedItem



    $ProgressBar1.Value = 30
    $LabelStatus.Text = "Disabling $x"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
    Write-Host "Disabling $x" -foregroundColor Yellow
    Add-Content "$logfile" "Disabling $x"
       
    ## Disable account to block user logins
    Set-AzureADUser -objectid $x -AccountEnabled $false
    write-host -foregroundcolor $colourInfo "Disabled login"

    $ProgressBar1.Value = 40
    $LabelStatus.Text = "Revoking Tokens"
    Revoke-AzureADUserAllRefreshToken -ObjectId $x
    write-host -foregroundcolor $colourAlert "Revoked Token"
    Add-Content "$logfile" "Revoking Tokens"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
##Set-CASMailbox $x -ActiveSyncEnabled $false
#write-host -foregroundcolor $colourAlert "ActiveSync disabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
#Set-casmailbox $x -owaenabled $false
#write-host -foregroundcolor $colourAlert "OWA disabled"

    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Removing devices from exchange"
    ## TheActiveSyncAllowedDeviceIDs parameter specifies one or more Exchange ActiveSync device IDs that are allowed to synchronize with the mailbox.
    ## Setting this to $NULL clears the list of device IDs
    Set-casmailbox $x -activesyncalloweddeviceids $null
    write-host -foregroundcolor $colourAlert "Removed allowed ActiveSync devices"
    Add-Content "$logfile" "Removed Exchange Activesync Devices"

## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
#Set-casmailbox $x -mapienabled $false
#write-host -foregroundcolor $colourAlert "Disabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
#Set-casmailbox $x -OWAforDevicesEnabled $false
#write-host -foregroundcolor $colourAlert "Disabled MAPI fo devices"

    $ProgressBar1.Value = 60
    $LabelStatus.Text = "Disabling POP/IMAP"
## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
Set-casmailbox $x -popenabled $false
write-host -foregroundcolor $colourAlert "Disabled POP"
Add-Content "$logfile" "Disabled POP"

    $ProgressBar1.Value = 70
    $LabelStatus.Text = "Disabling POP/IMAP"
## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
Set-casmailbox $x -imapenabled $false
write-host -foregroundcolor $colourAlert "Disabled IMAP"
Add-Content "$logfile" "Disabled IMAP"

## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
#Set-casmailbox $x -universaloutlookenabled $false
#write-host -foregroundcolor $colourAlert "Disabled Outlook"


    $ProgressBar1.Value = 80
    $LabelStatus.Text = "Converting Mailbox to Shared"
    write-host -foregroundcolor $colourAlert "Converting Mailbox to Shared"
    #Convert mailbox to Shared
    Set-Mailbox -Identity $x -Type Shared
    Start-Sleep -Seconds 5
    Add-Content "$logfile" "Converted mailbox to Shared"


    $ProgressBar1.Value = 95
    $LabelStatus.Text = "Removing Licenses"
    write-host -foregroundcolor $colourAlert "Removing Licenses"
    #Remove user license
    $license = (Get-MsolUser -UserPrincipalName $x).licenses.accountskuid
    Set-MsolUserLicense -UserPrincipalName $x -RemoveLicenses $license
    Add-Content "$logfile" "Removed license: $license"




    write-host -foregroundcolor Green "Finished offboarding"
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Completed Tasks, see log folder for export"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
})


$Form.ShowDialog()