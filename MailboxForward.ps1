## Why are you in here? Please speak with Dale if you have any errors ##

write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '        FORWARDING EMAILS         '
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


$Form                 = New-Object system.Windows.Forms.Form
$Form.ClientSize      = '800,420'
$Form.text            = "Give Calendar Access"
$Form.TopMost         = $false
$Form.StartPosition = 'CenterScreen'

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
$Identity.text                   = "Account to Manage"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(62,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)

$User                            = New-Object system.Windows.Forms.Label
$User.text                       = "Forward Emails To"
$User.AutoSize                   = $true
$User.width                      = 25
$User.height                     = 10
$User.location                   = New-Object System.Drawing.Point(346,49)
$User.Font                       = 'Microsoft Sans Serif,10'
$Form.Controls.Add($User)

$ListBox3                        = New-Object system.Windows.Forms.ListBox
$ListBox3.text                   = "listBox"
$ListBox3.width                  = 175
$ListBox3.height                 = 260
@("Yes","No") | ForEach-Object {[void] $ListBox3.Items.Add($_)}
$ListBox3.location               = New-Object System.Drawing.Point(595,77)

$permission                      = New-Object system.Windows.Forms.Label
$permission.text                 = "Keep Copy in original mailbox?"
$permission.AutoSize             = $true
$permission.width                = 25
$permission.height               = 10
$permission.location             = New-Object System.Drawing.Point(598,49)
$permission.Font                 = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($ListBox1,$ListBox2,$Button1,$Button2,$Identity,$User,$ListBox3,$permission))

$Form.Controls.Add($ListBox1)
$Form.Controls.Add($ListBox2)
$Form.Controls.Add($ListBox3)


$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $ListBox2.SelectedItem
    $z = $ListBox3.SelectedItem

write-host 'Updating Forwarding Rules for' $x
if ($z -eq "Yes") {
    Write-Host 'Forwarding emails and keeping a copy in' $x "'s mailbox"
    Set-Mailbox -Identity $x -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $y
    write-host -ForegroundColor Green 'Completed Task'
} elseif ($z -eq "No") {
    Write-Host 'Forwarding emails WITHOUT keeping a copy in' $x "'s mailbox"
    Set-Mailbox -Identity $x -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $y
    write-host -ForegroundColor Green 'Completed Task'
} else {
write-host 'Please try again, and choose whether to Keep a copy of emails or not in the original mailbox...'
}


Start-Sleep -Seconds 4
exit
}





