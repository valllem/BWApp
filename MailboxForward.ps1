$message  = ''
$question = 'Do you want to keep a copy in the original mailbox? (Recommend: Yes)'
$choices  = '&Yes', '&No'


write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green '             Forward Emails            '
write-Host -ForegroundColor Green '======================================='
write-Host -ForegroundColor Green ' '
##-- Prompt for mailbox the user will access--##
$UserIdentity = Read-Host 'Enter email of account to forward...'

##-- Prompt to enter user email that needs access--##
$UserEmail = Read-Host 'Enter email you are forwarding to...'

##--Run Script--##

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
if ($decision -eq 0) {
    Write-Host 'Forwarding emails and keeping a copy in' $UserIdentity"'s mailbox"
    Set-Mailbox -Identity $UserIdentity -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $UserEmail
} else {
    Write-Host 'Forwarding emails without keeping a copy in' $UserIdentity+"'s mailbox"
    Set-Mailbox -Identity $UserIdentity -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $UserEmail
}


##--Show result--##

write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()