Clear-Host
write-host -ForegroundColor Yellow 'Signing Out'
Exit-PSSession
Write-Host -ForegroundColor Yellow "Signed Out"
write-host 'Exiting App'
Start-Sleep -Seconds 2
Exit