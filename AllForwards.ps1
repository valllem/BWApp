$FilePath = "C:\BWApp\Logs"
Write-Host -ForegroundColor Yellow "Preparing File"

Get-Mailbox | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward | Export-csv C:\BWApp\Logs\Office365Forwards.csv -NoTypeInformation 
Get-Mailbox | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward | Out-GridView
Write-Host
Write-Host
Write-Host
Write-Host -ForegroundColor Green "File saved to $Path\Office365Forwards.csv"
Write-Host -ForegroundColor Green "Opening File..."
Invoke-Item "$FilePath\Office365Forwards.csv"
Write-Host -ForegroundColor Green "TASK COMPLETE"