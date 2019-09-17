Write-Host -ForegroundColor Yellow '    SET USERS FIRSTNAME  '
Write-Host -ForegroundColor Yellow '  ======================='
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''

Write-Host -ForegroundColor Yellow 'User changing name...'
$userToRename = Read-Host 'Enter users email... '


Start-Sleep -Seconds 2

Write-Host -ForegroundColor Yellow 'Specify the new First Name '
$FirstName = Read-Host 'Enter new First Name'
Start-Sleep -Seconds 1
Write-Host -ForegroundColor Yellow 'Specify the new Last Name '
$LastName = Read-Host 'Enter new Last Name'

Write-Host -ForegroundColor Yellow 'Changing First Name '
Set-MsolUser -UserPrincipalName $userToRename -FirstName $FirstName

Write-Host -ForegroundColor Yellow 'Changing Last Name '
Set-MsolUser -UserPrincipalName $userToRename -LastName $LastName

Write-Host -ForegroundColor Yellow 'Processing... '
Start-Sleep -Seconds 2

Set-MsolUser -UserPrincipalName $userToRename -DisplayName " $FirstName $LastName "


Write-Host -ForegroundColor Green 'Tasks Completed. '
Start-Sleep -Seconds 1

Write-Host -ForegroundColor Green 'Press any key to return to previous menu'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
Exit