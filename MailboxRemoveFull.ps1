$message  = ''
$question = 'Do you want to remove this users access?'
$choices  = '&Yes', '&No'


write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '        REMOVE USER FULL ACCESS         '
write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green ' '
##-- Prompt for mailbox the user will access--##
$UserIdentity = Read-Host 'Enter email of account to Manage'

##-- Prompt to enter user email that needs access--##
$UserEmail = Read-Host 'Who access do you want to revoke from' $UserIdentity '?'

##-- Shows Details --##

Write-Host -ForegroundColor Yellow ' This will revoke' $UserEmail +"'s Full Access to" $UserIdentity ''
##--Run Script--##

    Write-Host 'Removing Full Access'
    Remove-MailboxPermission –Identity $UserIdentity -User $UserEmail -AccessRights FullAccess -InheritanceType All



##--Show result--##

write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'



$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()






