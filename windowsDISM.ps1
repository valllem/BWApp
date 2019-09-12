clear-host
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''


Write-Host -ForegroundColor Yellow '  ██████╗ ██╗███████╗███╗   ███╗'
Write-Host -ForegroundColor Yellow '  ██╔══██╗██║██╔════╝████╗ ████║'
Write-Host -ForegroundColor Yellow '  ██║  ██║██║███████╗██╔████╔██║'
Write-Host -ForegroundColor Yellow '  ██║  ██║██║╚════██║██║╚██╔╝██║'
Write-Host -ForegroundColor Yellow '  ██████╔╝██║███████║██║ ╚═╝ ██║'
Write-Host -ForegroundColor Yellow '  ╚═════╝ ╚═╝╚══════╝╚═╝     ╚═╝'
Write-Host -ForegroundColor Yellow ' '
Write-Host -ForegroundColor Yellow ''


Write-host -ForegroundColor Yellow '============================================='
Write-host -ForegroundColor Yellow 'This will download some files before running.'
Write-host -ForegroundColor Yellow 'This may take a while, please be patient.'
Write-host -ForegroundColor Yellow '============================================='
Dism.exe /Online /Cleanup-Image /Restorehealth



Write-Host -ForegroundColor Green 'Task Finished, press any key to return to previous menu...'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
