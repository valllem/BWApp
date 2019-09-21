

## SCRIPT
cd $HOME
cd downloads
$Path = "C:\BWApp\Logs"

Get-Mailbox | Get-mailboxPermission | Where-Object { ($_.accessRights -like "*fullaccess*") -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "*\Domain Admins")-and -not ($_.User -like "*\Organisations-Admins") -and -not ($_.User -like "*\Organization Management") -and -not ($_.User -like "*\Administrator") -and -not ($_.User -like "*\Exchange Servers") -and -not ($_.User -like "*\Exchange Trusted Subsystem") -and -not ($_.User -like "*\Enterprise Admins") -and -not ($_.User -like "*\system")} | Export-CSV "$Path\permission.CSV"
Write-Host -ForegroundColor Green "Check your downloads folder for the Permission.csv file"


Exit-PSSession