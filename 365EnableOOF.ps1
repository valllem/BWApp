Clear-Host

Write-Host -ForegroundColor Yellow 'Enabling Out of Office Reply'
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''

$UserIdentity = Read-Host 'Enter the Email of account to manage...'

$AutoReplyMessage = Read-Host 'Please type your message and press enter...'

Set-MailboxAutoReplyConfiguration -identity $UserIdentity -AutoReplyState Enabled -InternalMessage $AutoReplyMessage  -ExternalMessage $AutoReplyMessage



Write-Host -ForegroundColor Green 'Out of Office Reply has been set'
Write-Host -ForegroundColor Green 'Press any key to return to previous menu'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
Exit