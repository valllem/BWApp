
$hours = 48     
$Path = "C:\BWApp\Logs"



Write-Host
write-host -foregroundcolor Cyan "Message Trace - 48 Hours"
write-host -foregroundcolor Cyan "==========================="
write-host -foregroundcolor Cyan "This can take a few moments to load, depending on the logs...`n"

$dateEnd = get-date                         ## get current time
$dateStart = $dateEnd.AddHours(-$hours)     ## get current time less last $hours
$results = Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID
$results | out-gridview
$results | Export-Csv $Path\MessageTrace.csv


write-host -foregroundcolor Yellow "A copy of the Message Trace has been put in $Path\MessageTrace.csv "
write-host -foregroundcolor Green "Tasks completed`n"