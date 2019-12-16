Write-Host -ForegroundColor Yellow 'Enabling Audit Logging in 365...'
Write-Host -ForegroundColor Yellow '=============================='
Write-Host -ForegroundColor Yellow 'This can take several seconds'
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Write-Host -ForegroundColor Green 'Task Completed!'
Write-Host -ForegroundColor Yellow '=============================='
