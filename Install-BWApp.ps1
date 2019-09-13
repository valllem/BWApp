param(
    [string] $Version = "0.1.0"
)

$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = Get-Location
$output = [IO.Path]::Combine($Path, "BWApp$Version.zip”)
    
Write-Host "Downloading BWApp from $url to path " $Path -ForegroundColor Green 
    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
    
# Unzip the Archive
Expand-Archive $output -DestinationPath $Path
    
#Set the environment variable
$Home = [IO.Path]::Combine($Path, "BWApp")
    
[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")