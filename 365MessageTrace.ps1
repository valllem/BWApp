
$hours = 48     
$FilePath = "C:\BWApp\Logs"



Write-Host
write-host -foregroundcolor Cyan "Message Trace - 48 Hours"
write-host -foregroundcolor Cyan "==========================="
write-host -foregroundcolor Cyan "This can take a few moments to load, depending on the logs...`n"

$dateEnd = get-date                         ## get current time
$dateStart = $dateEnd.AddHours(-$hours)     ## get current time less last $hours
$results = Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID
$results | out-gridview
$results | Export-Csv $FilePath\MessageTrace.csv


Write-Host -ForegroundColor Green "Saved Message Logs to: $FilePath\MessageTrace.csv"
Write-Host -ForegroundColor Green "Opening Logs..."
Start-Sleep -Seconds 2