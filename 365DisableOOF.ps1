write-host -ForegroundColor Yellow 'Out of Office'
write-host -ForegroundColor Yellow '====================='
$colourAlert = "Yellow"
$colourInfo = "Green"


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '382,439'
$Form.text                       = "Out of Office"
$Form.TopMost                    = $false
$Form.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox1"
$ListBox1.width                  = 171
$ListBox1.height                 = 277
$ListBox1.location               = New-Object System.Drawing.Point(18,54)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(265,362)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $Button1
$Form.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(192,362)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $Button2
$Form.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Disable OOF for user..."
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(26,26)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)





$Form.controls.AddRange(@($Button1,$Button2,$Identity,$ListBox1))

$Form.Controls.Add($Textbox)



$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $Textbox.Text
    $y = $ListBox1.SelectedItem

try {
Write-Host "Disabling autoreply for " $y
Set-MailboxAutoReplyConfiguration -identity $y -AutoReplyState Disabled
}
catch  {
Write-Host -ForegroundColor Red " An error occured. Please report this bug. (365EnableOOF)"

}

write-host -ForegroundColor Green "Completed all Tasks"

}

