
$BinConsole = "C:\BWApp\Bin\settings_console.dll"
$GetBinConsole = Get-Content $BinConsole
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '259,160'
$Form.text                       = "Console Settings"
$Form.TopMost                    = $false

$CheckBoxShowConsole             = New-Object system.Windows.Forms.CheckBox
$CheckBoxShowConsole.text        = "Show Console Window"
$CheckBoxShowConsole.AutoSize    = $false
$CheckBoxShowConsole.width       = 203
$CheckBoxShowConsole.height      = 20
$CheckBoxShowConsole.location    = New-Object System.Drawing.Point(6,63)
$CheckBoxShowConsole.Font        = 'Microsoft Sans Serif,10'

$ButtonApply                     = New-Object system.Windows.Forms.Button
$ButtonApply.text                = "Apply Changes"
$ButtonApply.width               = 144
$ButtonApply.height              = 30
$ButtonApply.location            = New-Object System.Drawing.Point(6,6)
$ButtonApply.Font                = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 242
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(8,90)

$Form.controls.AddRange(@($CheckBoxShowConsole,$ButtonApply,$ProgressBar1))

#Checks
if ($GetBinConsole -like 'Enabled'){$CheckBoxShowConsole.checked = $true
write-host 'Console Window Enabled'}
if ($GetBinConsole -like 'Disabled'){$CheckBoxShowConsole.checked = $false}


$ButtonApply.Add_Click({
if ($CheckBoxShowConsole.Checked){
    Set-Content $BinConsole "Enabled" #default: Disabled = Do not show the console window
    write-host 'Console Window Enabled'
    }
    else {
    Set-Content $BinConsole "Disabled" #default: Disabled = Do not show the console window
    write-host 'Console Window Disabled'
    }
    $ProgressBar1.Refresh()
    $ProgressBar1.Value = 100
    
    Start-Sleep -milliseconds 500


})




$Form.ShowDialog()