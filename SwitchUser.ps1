Clear-host 
write-host -ForegroundColor Yellow 'Signing Out'
Disconnect-PSSession $EXOSession
Exit-PSSession
Write-Host -ForegroundColor Yellow "Signed Out, Switching User"
Start-Sleep -Seconds 1

Write-Host -ForegroundColor Yellow "Signing In`n"

Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
-Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"} `
|select -First 1)
write-host -foregroundcolor Yellow "Exchange Online MFA module loaded"
$EXOSession = New-ExoPSSession
Import-PSSession $EXOSession
write-host -foregroundcolor Yellow "Connected to Exchange Online MFA`n"

write-host 'Session active'

Write-Host -ForegroundColor Green 'Sign In Procedure Complete...'
Write-Host -ForegroundColor Green 'APP READY'