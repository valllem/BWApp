$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
$FilePath = "C:\BWApp\logs"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '350,220'
$Form.text                       = "Check All Mailbox Forwarding"
$Form.TopMost                    = $false

$ButtonRunScript                 = New-Object system.Windows.Forms.Button
$ButtonRunScript.BackColor       = "#7ed321"
$ButtonRunScript.text            = "Run Script"
$ButtonRunScript.width           = 220
$ButtonRunScript.height          = 30
$ButtonRunScript.location        = New-Object System.Drawing.Point(60,60)
$ButtonRunScript.Font            = 'Microsoft Sans Serif,10,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 220
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(60,115)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status: Ready"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 25
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(60,100)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'



$Form.controls.AddRange(@($ButtonRunScript,$ProgressBar1,$LabelStatus))

$ButtonRunScript.Add_Click({
    
   
    $ProgressBar1.Value = 25
    $LabelStatus.Text = "Checking..."
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
    
    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Checking permissions"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()

    Write-Host -ForegroundColor Yellow "Preparing File"
    Add-Content "$FilePath\log.txt" "Getting All Forwarding"

    Get-Mailbox | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward | Export-csv C:\BWApp\Logs\MailboxForwarding.csv -NoTypeInformation 
    Get-Mailbox | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward | Out-GridView -Title "All Mailbox Forwarding | Exported to $FilePath\MailboxForwarding.csv"

    Write-Host -ForegroundColor Green "File saved to $Path\MailboxForwarding.csv"
    Add-Content "$FilePath\log.txt" "Exported to $FilePath\MailboxForwarding.csv"


    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Completed Tasks, see log folder for export"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
    Start-Sleep -Milliseconds 300
    $Form.Close()
})


$Form.ShowDialog()