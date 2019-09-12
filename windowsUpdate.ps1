Install-Module PSWindowsUpdate -Force
Get-WindowsUpdate
Install-WindowsUpdate -Force

Write-Host -ForegroundColor Yellow 'Task Finished, press any key to return to previous menu...'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()