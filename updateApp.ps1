$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp.zip”)

Write-Progress -Activity "Updating BWApp" -Status "Getting Ready" -PercentComplete 15    

Write-Progress -Activity "Updating BWApp" -Status "Downloading Components" -PercentComplete 30
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Clear-Host
Start-Sleep -Seconds 2    

Write-Progress -Activity "Updating BWApp" -Status "Extracting..." -PercentComplete 45
Clear-Host
Start-Sleep -Seconds 2  
# Unzip the Archive
Expand-Archive $output -DestinationPath $Path -Force
Clear-Host
Write-Progress -Activity "Updating BWApp" -Status "Tidying Up" -PercentComplete 60
## I need to add the tidy up script here.
Clear-Host
Start-Sleep -Seconds 2     

Write-Progress -Activity "Updating BWApp" -Status "Updating shortcuts..." -PercentComplete 75
Start-Sleep -Seconds 2 
cd $HOME
cd desktop
$ShortCutDir = Get-Location
Clear-Host
Start-Sleep -Seconds 2

function set-shortcut {
param ( [string]$SourceLnk, [string]$DestinationPath )
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
    $Shortcut.Save()
    }
set-shortcut "$ShortcutDir\BWApp.lnk" "$Path\BWApp-master\Launcher.ps1"
Clear-Host
Start-Sleep -Seconds 2
cd "C:\BWApp\BWApp-master"
Write-Progress -Activity "Updating BWApp" -Status "Finishing Update..." -PercentComplete 90
Clear-Host
Start-Sleep -Seconds 2 

Write-Progress -Activity "Updating BWApp" -Status "FINISHED UPDATE" -PercentComplete 100
Start-Sleep -Seconds 1 