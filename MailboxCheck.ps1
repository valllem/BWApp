write-host 'Check Mailbox Access'
write-host '====================='
$colourAlert = "Yellow"
$colourInfo = "Green"

$FilePath = "C:\BWApp\Logs"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,420'
$Form.text                       = "Check Mailbox Access"
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
$Identity.text                   = "Mailbox to Check"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(38,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)

$Form.controls.AddRange(@($ListBox1,$Button1,$Button2,$Identity,$ListBox3,$permission))

$Form.Controls.Add($ListBox1)



$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $mailbox.PrimarySmtpAddress
 
$results = Get-MailboxPermission -identity $x
$results | Where-Object { ($_.accessRights -like "*") -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "*\Domain Admins")-and -not ($_.User -like "*\Organisations-Admins") -and -not ($_.User -like "*\Organization Management") -and -not ($_.User -like "*\Administrator") -and -not ($_.User -like "*\Exchange Servers") -and -not ($_.User -like "*\Exchange Trusted Subsystem") -and -not ($_.User -like "*\Enterprise Admins") -and -not ($_.User -like "*\system")} | out-gridview
$results | Where-Object { ($_.accessRights -like "*") -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "*\Domain Admins")-and -not ($_.User -like "*\Organisations-Admins") -and -not ($_.User -like "*\Organization Management") -and -not ($_.User -like "*\Administrator") -and -not ($_.User -like "*\Exchange Servers") -and -not ($_.User -like "*\Exchange Trusted Subsystem") -and -not ($_.User -like "*\Enterprise Admins") -and -not ($_.User -like "*\system")} | Export-Csv "$FilePath\MailboxPerms.csv"
Invoke-Item "$FilePath\MailboxPerms.csv"
write-host -ForegroundColor Green "Completed all Tasks"




}

