Clear-Host

$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"
$hours = 48     




write-host -foregroundcolor $systemmessagecolor "Script started`n"

$dateEnd = get-date                         ## get current time
$dateStart = $dateEnd.AddHours(-$hours)     ## get current time less last $hours
$results = Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID
$results | out-gridview

write-host -foregroundcolor $systemmessagecolor "Script completed`n"