Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '535,137'
$Form.text                       = "Run a Command"
$Form.TopMost                    = $false

$TextBoxCommand                  = New-Object system.Windows.Forms.TextBox
$TextBoxCommand.multiline        = $false
$TextBoxCommand.text             = "Get-User"
$TextBoxCommand.width            = 495
$TextBoxCommand.height           = 20
$TextBoxCommand.location         = New-Object System.Drawing.Point(20,54)
$TextBoxCommand.Font             = 'Microsoft Sans Serif,10'

$ButtonOK                        = New-Object system.Windows.Forms.Button
$ButtonOK.text                   = "OK (Run)"
$ButtonOK.width                  = 75
$ButtonOK.height                 = 30
$ButtonOK.location               = New-Object System.Drawing.Point(441,84)
$ButtonOK.Font                   = 'Microsoft Sans Serif,10'
$ButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $ButtonOK
$Form.Controls.Add($ButtonOK)

$ButtonCancel                    = New-Object system.Windows.Forms.Button
$ButtonCancel.text               = "Cancel"
$ButtonCancel.width              = 60
$ButtonCancel.height             = 30
$ButtonCancel.location           = New-Object System.Drawing.Point(368,84)
$ButtonCancel.Font               = 'Microsoft Sans Serif,10'
$Form.CancelButton = $ButtonCancel
$Form.Controls.Add($ButtonCancel)

$Form.controls.AddRange(@($TextBoxCommand,$ButtonOK,$ButtonCancel))

$Form.Controls.Add($TextBoxCommand)
$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $TextBoxCommand.Text
    
$custom = Invoke-Expression $x | Export-Csv "C:\BWApp\Logs\Custom.csv"
Invoke-Item "C:\BWApp\Logs\Custom.csv"

write-host -ForegroundColor Green "Completed all Tasks"

Start-Sleep -Seconds 4


}