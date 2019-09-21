if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}


New-Item -ItemType Directory -Force -Path "C:\BWApp"
$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
    
Write-Host "Downloading BWApp from $url to path " $Path -ForegroundColor Yellow 

Start-Sleep -Seconds 2
    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Sleep -Seconds 2    
# Unzip the Archive
Expand-Archive $output -DestinationPath $Path -Force
Start-Sleep -Seconds 2    
#Set the environment variable
##$Home = [IO.Path]::Combine($Path, "BWApp")
    
##[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")

Start-Sleep -Seconds 2
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
read-host -Prompt 'Press Enter to Close Installer'
