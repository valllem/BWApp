$message  = ''
$question = 'Do you want to automatically add this account to Outlook?'
$choices  = '&Yes', '&No'


write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '        GRANT USER FULL ACCESS         '
write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green ' '
##-- Prompt for mailbox the user will access--##
$UserIdentity = Read-Host 'Enter email of account to Manage'

##-- Prompt to enter user email that needs access--##
$UserEmail = Read-Host 'Who are you going to give access to' $UserIdentity '?'

##-- Shows Details --##

Write-Host -ForegroundColor Yellow ' The account' $UserEmail 'will be granted Full Access to' $UserIdentity"'s mailbox"
##--Run Script--##

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
if ($decision -eq 0) {
    Write-Host 'Granting Full Access and adding to Outlook'
    Add-MailboxPermission –Identity $UserIdentity -User $UserEmail -AccessRights FullAccess -AutoMapping $true
} else {
    Write-Host 'Granting Full Access'
    Add-MailboxPermission –Identity $UserIdentity -User $UserEmail -AccessRights FullAccess -AutoMapping $false
}


##--Show result--##

write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'



$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()






