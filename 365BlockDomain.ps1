Clear-Host

Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow 'Block a domain...'

$SenderDomain = Read-Host 'Domain to block'

Set-HostedContentFilterPolicy default -BlockedSenderDomains @{add="$SenderDomain"} 

Write-Host -ForegroundColor Green " $SenderDomain has been added to the blocked senders list."

Write-Host -ForegroundColor Green '======================================='
Write-Host -ForegroundColor Green '         All Tasks Complete!           '


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()