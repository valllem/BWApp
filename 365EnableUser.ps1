$colourAlert = "Yellow"
$colourInfo = "Green"


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,600'
$Form.text                       = "Enable User"
$Form.TopMost                    = $false
$Form.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 300
$ListBox1.height                 = 258
$ListBox1.location               = New-Object System.Drawing.Point(38,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(280,372)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $Button1
$Form.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(208,373)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $Button2
$Form.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Account to Enable"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(66,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)

$Form.controls.AddRange(@($ListBox1,$Button1,$Button2,$Identity,$ListBox3,$permission))

$Form.Controls.Add($ListBox1)



$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $mailbox.PrimarySmtpAddress
 
   
 ## Enable account to block user logins
Set-AzureADUser -objectid $x -AccountEnabled $true
write-host -foregroundcolor $colourInfo "Enable login"
Revoke-AzureADUserAllRefreshToken -ObjectId $x
write-host -foregroundcolor $colourAlert "Revoked Token"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
Set-CASMailbox $x -ActiveSyncEnabled $true
write-host -foregroundcolor $colourAlert "ActiveSync Enabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
Set-casmailbox $x -owaenabled $true
write-host -foregroundcolor $colourAlert "OWA Enabled"

## TheActiveSyncAllowedDeviceIDs parameter specifies one or more Exchange ActiveSync device IDs that are allowed to synchronize with the mailbox.
## Setting this to $NULL clears the list of device IDs
Set-casmailbox $x -activesyncalloweddeviceids $null
write-host -foregroundcolor $colourAlert "Removed allowed ActiveSync devices"

## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
Set-casmailbox $x -mapienabled $true
write-host -foregroundcolor $colourAlert "Enabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
Set-casmailbox $x -OWAforDevicesEnabled $true
write-host -foregroundcolor $colourAlert "Enabled MAPI fo devices"

## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
Set-casmailbox $x -universaloutlookenabled $true
write-host -foregroundcolor $colourAlert "Enabled Outlook"

## User will be signed out of browser, desktop and mobile applications accessing Office 365 resources across all devices.
## It can take up to an hour to sign out from all devices.
Revoke-SPOUserSession -user $x -Confirm:$false


write-host -ForegroundColor Cyan "Completed all Tasks"
 
   



}



