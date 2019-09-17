param(
    [string] $Version = "0.1.3"
)

$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = Get-Location
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
    
Write-Host "Downloading BWApp from $url to path " $Path -ForegroundColor Green 

Start-Sleep -Seconds 2
    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
    
# Unzip the Archive
Expand-Archive $output -DestinationPath $Path
    
#Set the environment variable
##$Home = [IO.Path]::Combine($Path, "BWApp")
    
##[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")

Start-Sleep -Seconds 2