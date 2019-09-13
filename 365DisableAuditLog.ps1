Clear-Host

##--Info Text
Write-Host -ForegroundColor Yellow 'DISABLING Audit Logging in 365...'

##--Script
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $False -confirm



##--Info Text
Write-Host -ForegroundColor Green 'Task Completed!'
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green ''
write-Host -ForegroundColor Green 'Please allow a few minutes for any changes to take effect.'
write-Host -ForegroundColor Green 'Press any key to return to previous menu.'

##--Wait for key press to continue
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()