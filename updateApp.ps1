$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
    
Write-Host "Downloading BWApp Components" -ForegroundColor Yellow 
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Sleep -Seconds 2    
# Unzip the Archive
write-host -foregroundcolor Yellow "Extracting..."
Expand-Archive $output -DestinationPath $Path -Force
Start-Sleep -Seconds 2    
Exit-PSSession
$BWApp.Close()
#Set the environment variable
##$Home = [IO.Path]::Combine($Path, "BWApp")
    
##[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")

Start-Sleep -Seconds 2
write-host -foregroundcolor Yellow "Creating Shortcuts"
cd $HOME
cd desktop
$ShortCutDir = Get-Location
write-host $ShortCutDir
Start-Sleep -Seconds 2
function set-shortcut {
param ( [string]$SourceLnk, [string]$DestinationPath )
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
    $Shortcut.Save()
    }
set-shortcut "$ShortcutDir\BWApp.lnk" "$Path\BWApp-master\Launcher.ps1"
Start-Sleep -Seconds 2
Clear-Host
cd "C:\BWApp\BWApp-master"

Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host -ForegroundColor Green "INSTALLATION COMPLETE.          "
write-host 'Restarting App...'
.\launcher.ps1
Exit
