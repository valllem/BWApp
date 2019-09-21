## Why are you in here? Please speak with Dale if you have any errors ##

write-Host -ForegroundColor Yellow '======================================='
write-Host -ForegroundColor Yellow '    GRANT USER ALL MAILBOX ACCESS     '
write-Host -ForegroundColor Yellow '======================================='
write-Host -ForegroundColor Yellow ' '


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

try {
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
}

catch {
write-host -ForegroundColor Red "NOT LOGGED IN. Please select 'switch account' on the main menu "
}

$calendarAccess1                 = New-Object system.Windows.Forms.Form
$calendarAccess1.ClientSize      = '500,420'
$calendarAccess1.text            = "Give user access to ALL mailboxes"
$calendarAccess1.TopMost         = $false
$calendarAccess1.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 169
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
$calendarAccess1.AcceptButton = $Button1
$calendarAccess1.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(208,373)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$calendarAccess1.CancelButton = $Button2
$calendarAccess1.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Who can access ALL mailboxes?"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(40,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($Identity)

$ListBox3                        = New-Object system.Windows.Forms.ListBox
$ListBox3.text                   = "listBox"
$ListBox3.width                  = 175
$ListBox3.height                 = 260
@("Yes","No") | ForEach-Object {[void] $ListBox3.Items.Add($_)}
$ListBox3.location               = New-Object System.Drawing.Point(250,77)

$permission                      = New-Object system.Windows.Forms.Label
$permission.text                 = "Add to Outlook?"
$permission.AutoSize             = $true
$permission.width                = 25
$permission.height               = 10
$permission.location             = New-Object System.Drawing.Point(260,49)
$permission.Font                 = 'Microsoft Sans Serif,10'

$calendarAccess1.controls.AddRange(@($ListBox1,$Button1,$Button2,$Identity,$ListBox3,$permission))

$calendarAccess1.Controls.Add($ListBox1)
$calendarAccess1.Controls.Add($ListBox3)


$calendarAccess1.Topmost = $true


$result = $calendarAccess1.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $mailbox.PrimarySmtpAddress
    $z = $ListBox3.SelectedItem
    

$accessRight = "FullAccess"
$mailboxes = Get-mailbox
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess.PrimarySmtpAddress -erroraction SilentlyContinue
         
if ($z -eq "Yes") {
    Remove-MailboxPermission -Identity $y -User "$x" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
    write-host "Adding $x Permissions to... $($mailbox.displayname) "
    Add-MailboxPermission -Identity $y -User "$x" -AccessRights FullAccess -AutoMapping $true
} elseif ($z -eq "No") {
    Remove-MailboxPermission -Identity $y -User "$x" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
    write-host "Adding $x Permissions to... $($mailbox.displayname) "
    Add-MailboxPermission -Identity $y -User "$x" -AccessRights FullAccess -AutoMapping $false
} else {
write-host 'An error occurred'
}
}
Write-Host -ForegroundColor Green '======================================='
Write-Host -ForegroundColor Green '         All Tasks Complete!           '

}

