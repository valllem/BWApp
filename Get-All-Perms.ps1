

## SCRIPT

$FilePath = "C:\BWApp\Logs"
Write-Host -ForegroundColor Yellow "Preparing File"
Get-Mailbox | Get-mailboxPermission | Where-Object { ($_.accessRights -like "*fullaccess*") -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "*\Domain Admins")-and -not ($_.User -like "*\Organisations-Admins") -and -not ($_.User -like "*\Organization Management") -and -not ($_.User -like "*\Administrator") -and -not ($_.User -like "*\Exchange Servers") -and -not ($_.User -like "*\Exchange Trusted Subsystem") -and -not ($_.User -like "*\Enterprise Admins") -and -not ($_.User -like "*\system")} | Export-CSV "$FilePath\permission.CSV"
Write-Host -ForegroundColor Green "Saved list of Permissions to: $FilePath\Permission.csv"
Write-Host -ForegroundColor Green "Opening File..."
Start-Sleep -Seconds 2
Invoke-Item "$FilePath\Permission.csv"

Exit-PSSession