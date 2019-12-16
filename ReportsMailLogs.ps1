
$hours = 48     
$FilePath = "C:\BWApp\Logs"

$logfile = "C:\BWApp\Logs\Log.txt"
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress
$FilePath = "C:\BWApp\logs"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '350,220'
$Form.text                       = "Check All Mailbox Permissions"
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
    $LabelStatus.Text = "Checking mail logs"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
    write-host -foregroundcolor Cyan "Message Trace - 48 Hours"
    write-host -foregroundcolor Cyan "==========================="
    write-host -foregroundcolor Cyan "This can take a few moments to load, depending on the logs...`n"
    Add-Content "$FilePath\log.txt" "Getting Mail logs for last 48 hours"

    $dateEnd = get-date                         ## get current time
    $dateStart = $dateEnd.AddHours(-$hours)     ## get current time less last $hours
    $results = Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID
    $results | out-gridview -Title "Mail Logs | 48 hours | Copy saved to: $FilePath\MessageTrace.csv "
    $results | Export-Csv $FilePath\MessageTrace.csv


    Write-Host -ForegroundColor Green "Saved Message Logs to: $FilePath\MessageTrace.csv"
    
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Completed Tasks. Saved Message Logs to: $FilePath\MessageTrace.csv"
    $ProgressBar1.Refresh()
    $LabelStatus.Refresh()
    Start-Sleep -Milliseconds 300
    $Form.Close()
})


$Form.ShowDialog()
