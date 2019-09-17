
Write-Host -ForegroundColor Yellow '    RENAME USERS UPN  '
Write-Host -ForegroundColor Yellow '  ===================='
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''

Write-Host -ForegroundColor Yellow 'Specify the Users Current Principal Name  '
$userToRename = Read-Host 'Current User Principal Name E.g. user@company.onmicrosoft.com '

Start-Sleep -Seconds 2

Write-Host -ForegroundColor Yellow 'Specify the new User Principal Name '
$userToReplace = Read-Host 'Enter new User Principal Name E.g. user@company.com.au '

Start-Sleep -Seconds 2

Write-Host -ForegroundColor Yellow 'Changing User Principal Name '

Start-Sleep -Seconds 1



Set-MsolUserPrincipalName -UserPrincipalName $userToRename -NewUserPrincipalName $userToReplace

Write-Host -ForegroundColor Green 'Tasks Completed. '
Start-Sleep -Seconds 2

Write-Host -ForegroundColor Green 'Press any key to return to previous menu'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
Exit