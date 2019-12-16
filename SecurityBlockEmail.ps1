write-host -ForegroundColor Yellow 'Block Email'
write-host -ForegroundColor Yellow '============'
$colourAlert = "Yellow"
$colourInfo = "Green"


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,200'
$Form.text                       = "Block Email (SPAM)"
$Form.TopMost                    = $false
$Form.StartPosition = 'CenterScreen'

$Textbox                        = New-Object system.Windows.Forms.Textbox
$Textbox.text                   = ""
$Textbox.width                  = 300
$Textbox.height                 = 258
$Textbox.location               = New-Object System.Drawing.Point(38,77)


$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(280,100)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $Button1
$Form.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(208,100)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $Button2
$Form.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Email to Block"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(66,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)

$Form.controls.AddRange(@($Textbox,$Button1,$Button2,$Identity,$ListBox3,$permission))

$Form.Controls.Add($Textbox)



$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $Textbox.Text

try {
Write-Host "Blocking " $x
Set-HostedContentFilterPolicy "BusinessWorks Policy" -BlockedSenders @{add="$x"}
Add-Content "$logfile" "$DateTime"
Add-Content "$Logfile" "Added sender to the Businessworks Policy"
}
catch  {
Write-Host -ForegroundColor Cyan "Blocking " $x
Set-HostedContentFilterPolicy default -BlockedSenders @{add="$x"}
Add-Content "$logfile" "$DateTime"
Add-Content "$Logfile" "Added sender to the Default Policy"
}

write-host -ForegroundColor Green "Completed all Tasks"

}