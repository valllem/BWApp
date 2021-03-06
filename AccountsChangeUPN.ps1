﻿

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = $null
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '456,377'
$Form.text                       = "Change UPN"
$Form.TopMost                    = $false

$LabelUPN                        = New-Object system.Windows.Forms.Label
$LabelUPN.text                   = "Select User to Rename"
$LabelUPN.AutoSize               = $true
$LabelUPN.width                  = 25
$LabelUPN.height                 = 10
$LabelUPN.location               = New-Object System.Drawing.Point(5,9)
$LabelUPN.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($LabelUPN)

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 185
$ListBox1.height                 = 315
$ListBox1.location               = New-Object System.Drawing.Point(5,29)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$LabelNewUPN                     = New-Object system.Windows.Forms.Label
$LabelNewUPN.text                = "New UPN (user@domain.com.au)"
$LabelNewUPN.AutoSize            = $true
$LabelNewUPN.width               = 25
$LabelNewUPN.height              = 10
$LabelNewUPN.location            = New-Object System.Drawing.Point(231,40)
$LabelNewUPN.Font                = 'Microsoft Sans Serif,10'
$Form.Controls.Add($LabelNewUPN)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$Textbox1.text                   = ""
$TextBox1.width                  = 191
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(231,64)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$ButtonOK                        = New-Object system.Windows.Forms.Button
$ButtonOK.text                   = "OK"
$ButtonOK.width                  = 60
$ButtonOK.height                 = 30
$ButtonOK.location               = New-Object System.Drawing.Point(342,313)
$ButtonOK.Font                   = 'Microsoft Sans Serif,10'
$ButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $ButtonOK
$Form.Controls.Add($ButtonOK)

$ButtonCancel                    = New-Object system.Windows.Forms.Button
$ButtonCancel.text               = "Cancel"
$ButtonCancel.width              = 60
$ButtonCancel.height             = 30
$ButtonCancel.location           = New-Object System.Drawing.Point(231,313)
$ButtonCancel.Font               = 'Microsoft Sans Serif,10'
$ButtonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $ButtonCancel
$Form.Controls.Add($ButtonCancel)


$Form.controls.AddRange(@($LabelUPN,$ListBox1,$LabelNewUPN,$TextBox1,$ButtonOK,$ButtonCancel))

$Form.Controls.Add($ListBox1)



$Form.Topmost = $true



$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $ListBox1.SelectedItem
    $y = $TextBox1.Text
    
write-host -ForegroundColor Yellow 'Changing User Principal Name to' $y    
Set-MsolUserPrincipalName -UserPrincipalName $x -NewUserPrincipalName "$y"

write-host -ForegroundColor Green 'Completed Tasks'
Start-Sleep -Seconds 2
exit
}