$FilePath = "C:\BWApp\Logs"
write-host -ForegroundColor Yellow "Getting email addresses, check the logs for csv file"
write-host -ForegroundColor Yellow "Opening results"
$result = Get-Recipient -ResultSize Unlimited | select DisplayName,RecipientType,EmailAddresses
$result | Out-GridView -Title "All email addresses - Check the logs folder for a csv export"
$result | Export-Csv "$FilePath\emailaddresses.csv"
