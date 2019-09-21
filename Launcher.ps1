$Path = "C:\BWApp\BWApp-master"
New-Item -ItemType Directory -Force -Path C:\TEMP\BWApp
New-Item -ItemType Directory -Force -Path C:\BWApp\Logs
$TempPath = "C:\Temp\BWApp\Path.txt"
Set-Content -Path $TempPath -value $Path
Write-Host "$Path"
write-host 'Loading App...'


C:\BWApp\BWApp-master\MENU.ps1
Start-Sleep -Seconds 2
exit