$message  = ''
$question = 'Do you want to remove this users access?'
$choices  = '&Yes', '&No'


write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '        Manage Send Permissions        '
write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green ' '
##-- Prompt for mailbox the user will access--##
$UserIdentity = Read-Host 'Enter email of account to Manage'

##-- Prompt to enter user email that needs access--##
$UserEmail = Read-Host 'Who do you want to allow to send mail for' $UserIdentity '?'


##--Run Script--##

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
if ($decision -eq 0) {
    Write-Host 'Removing Full Access'
    Remove-MailboxPermission –Identity $UserIdentity -User $UserEmail -InheritanceType All
} else {
    Write-Host 'Leaving current access as is'
    
}


##--Show result--##

write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'



$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()






