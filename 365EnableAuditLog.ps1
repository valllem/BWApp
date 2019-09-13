Clear-Host
##--Info Text
Write-Host -ForegroundColor Yellow 'Enabling Audit Logging in 365...'
Write-Host -ForegroundColor Yellow 'This can take several seconds'

##--Script
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true

##--Info Text
Write-Host -ForegroundColor Green 'Task Completed!'
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'

##--Wait for key press to continue
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()