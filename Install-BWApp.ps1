if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}

Write-Host "Preparing to Install"
Write-Host "Please allow the script to run if prompted. `n"
Set-ExecutionPolicy -Scope CurrentUser Unrestricted -Force
Read-Host -Prompt "Press Enter to begin Installation"



New-Item -ItemType Directory -Force -Path "C:\BWApp"
$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
    
Write-Host "Downloading BWApp Components" -ForegroundColor Yellow 


Write-Host -ForegroundColor Yellow 'CHECKING AND INSTALLING MODULES'
Write-Host -ForegroundColor Yellow '==============================='
                                                           

Write-Host -ForegroundColor Yellow 'Checking if required Modules are installed.'
Install-PackageProvider -Name NuGet -Force
##--Security Principal Check for If statement--##
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

Clear-Host

If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    
if (Get-Module -ListAvailable -Name AADRM) {
    Write-Host -ForegroundColor Green "Azure AD Right Management Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD Right Management module"
    Install-Module -Name AADRM -force
}

if (Get-Module -ListAvailable -Name AzureAD) {
    Write-Host -ForegroundColor Green "Azure AD Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD module"
    Install-Module -Name AzureAD -force
}
  
if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    Write-Host -ForegroundColor Green "Teams Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Teams Module"
    Install-Module -Name MicrosoftTeams -Force
}

if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    Write-Host -ForegroundColor Green "SharePoint Online Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing SharePoint Online module"
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force
}
    
if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host -ForegroundColor Green "Microsoft Online Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Microsoft Online module"
    Install-Module -Name MSOnline -force
}

if (Get-Module -ListAvailable -Name AzureRM) {
    Write-Host -ForegroundColor Green "Azure Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure module... This may take a few minutes"
    Install-Module -name AzureRM -Force
}    

if (Get-Module -ListAvailable -Name CreateExoPsSession) {
    Write-Host -ForegroundColor Green "Exchange MFA Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Exchange MFA module... This may take a few minutes"
    Install-Script -Name CreateExoPsSession -force 
}



  
    
}
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}

write-host -foregroundcolor Green "All Modules Installed"



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

Clear-Host


Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host -ForegroundColor Green "                  INSTALLATION COMPLETE.          "
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
read-host -Prompt 'Press Enter to Close Installer'
