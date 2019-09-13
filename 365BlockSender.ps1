Clear-Host

Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow 'Block an email address...'

$SenderEmail = Read-Host 'Email Address to block'

Set-HostedContentFilterPolicy default -BlockedSenders @{add="$SenderEmail"} 

Write-Host -ForegroundColor Green "$SenderEmail has been added to the blocked senders list."

Write-Host -ForegroundColor Green '======================================='
Write-Host -ForegroundColor Green '         All Tasks Complete!           '


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()