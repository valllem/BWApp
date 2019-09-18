$Path = Get-Location
New-Item -ItemType Directory -Force -Path C:\TEMP\BWApp
$TempPath = "C:\Temp\BWApp\Path.txt"
Set-Content -Path $TempPath -value $Path





.\MENU.ps1

exit