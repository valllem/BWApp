Clear-Host
write-Host -ForegroundColor Yellow 'Enabled status = UnifiedAuditLogIngestionEnabled : True'
write-Host -ForegroundColor Yellow 'Disabled status = UnifiedAuditLogIngestionEnabled : False'
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Green 'STATUS'
write-Host -ForegroundColor Green '==========================================='
Get-AdminAuditLogConfig | FL UnifiedAuditLogIngestionEnabled
write-Host -ForegroundColor Green '==========================================='


write-Host -ForegroundColor Green 'Press any key to return to previous menu.'

##--Wait for key press to continue
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()