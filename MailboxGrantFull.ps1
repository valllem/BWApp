

write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '        GRANT USER FULL ACCESS         '
write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green ' '


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
try {
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
}

catch {
write-host -ForegroundColor Red "NOT LOGGED IN. Please select 'switch account' on the main menu "
}
$calendarAccess1                 = New-Object system.Windows.Forms.Form
$calendarAccess1.ClientSize      = '800,420'
$calendarAccess1.text            = "Give Mailbox Access"
$calendarAccess1.TopMost         = $false
$calendarAccess1.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 169
$ListBox1.height                 = 258
$ListBox1.location               = New-Object System.Drawing.Point(38,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$ListBox2                        = New-Object system.Windows.Forms.ListBox
$ListBox2.text                   = "listBox"
$ListBox2.width                  = 166
$ListBox2.height                 = 258
$ListBox2.location               = New-Object System.Drawing.Point(327,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox2.Items.Add($mailbox.PrimarySmtpAddress)}

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
$Identity.text                   = "Account to Access"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(62,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($Identity)

$User                            = New-Object system.Windows.Forms.Label
$User.text                       = "Person Getting Access"
$User.AutoSize                   = $true
$User.width                      = 25
$User.height                     = 10
$User.location                   = New-Object System.Drawing.Point(340,49)
$User.Font                       = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($User)

$ListBox3                        = New-Object system.Windows.Forms.ListBox
$ListBox3.text                   = "listBox"
$ListBox3.width                  = 175
$ListBox3.height                 = 260
@("Yes","No") | ForEach-Object {[void] $ListBox3.Items.Add($_)}
$ListBox3.location               = New-Object System.Drawing.Point(595,77)

$permission                      = New-Object system.Windows.Forms.Label
$permission.text                 = "Add to Outlook?"
$permission.AutoSize             = $true
$permission.width                = 25
$permission.height               = 10
$permission.location             = New-Object System.Drawing.Point(623,49)
$permission.Font                 = 'Microsoft Sans Serif,10'

$calendarAccess1.controls.AddRange(@($ListBox1,$ListBox2,$Button1,$Button2,$Identity,$User,$ListBox3,$permission))

$calendarAccess1.Controls.Add($ListBox1)
$calendarAccess1.Controls.Add($ListBox2)
$calendarAccess1.Controls.Add($ListBox3)


$calendarAccess1.Topmost = $true


$result = $calendarAccess1.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $ListBox2.SelectedItem
    $z = $ListBox3.SelectedItem

write-host 'Removing any existing permissions in place for' $x
if ($z -eq "Yes") {

try {
    write-host 'Adding Permissions...'
    Add-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -AutoMapping $true
}
catch {
    Remove-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
    write-host 'Updating Permissions...'
    Add-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -AutoMapping $true
}
    
    
} elseif ($z -eq "No") {
try {
    write-host 'Adding Permissions...'
    Add-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -AutoMapping $false
    }
    catch {
    write-host 'Updating Permissions...'
    Remove-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
    Add-MailboxPermission -Identity $x -User "$y" -AccessRights FullAccess -AutoMapping $false
    }
} else {
write-host 'An error occurred. Code: MailboxGrantFull[103-134]'
}

write-host -ForegroundColor Green 'Completed Tasks'
Start-Sleep -Seconds 4
exit
}





