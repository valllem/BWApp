write-Host -ForegroundColor Green '                AUDIT STATUS               '
write-Host -ForegroundColor Green '==========================================='
Get-AdminAuditLogConfig | FL UnifiedAuditLogIngestionEnabled
write-Host -ForegroundColor Green '==========================================='
write-Host -ForegroundColor Yellow 'Enabled status = UnifiedAuditLogIngestionEnabled : True'
write-Host -ForegroundColor Yellow 'Disabled status = UnifiedAuditLogIngestionEnabled : False'
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow '------------------------------------------'

