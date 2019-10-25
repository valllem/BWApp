$logfile = "C:\BWApp\Logs\Log.txt"

##Create App, Temp and Log directory if needed
New-Item -ItemType Directory -Force -Path C:\TEMP\BWApp
New-Item -ItemType Directory -Force -Path C:\BWApp\Logs

$Path = "C:\BWApp\BWApp-master"
$TempPath = "C:\Temp\BWApp\Path.txt"
$RunningUser = whoami
$DateTime = Get-Date

##Clear log file if it still exists
##It should be renamed to the date and time when exiting app cleanly

if (Test-Path -Path C:\BWApp\Logs\Log.txt) 
    {
    Write-Host -ForegroundColor Yellow "Launching..."
    Clear-Content "C:\BWApp\Logs\Log.txt"
    Add-Content "$logfile" "====================="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "$RunningUser launched BWApp"
    }
else 
{
    Write-Host -ForegroundColor Green "Launching..."
    Add-Content "$logfile" "====================="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "$RunningUser launched BWApp"
    
}

   


##End of log tidy up



Set-Content -Path $TempPath -value $Path

powershell.exe "C:\BWApp\BWApp-master\MENU.ps1"

exit