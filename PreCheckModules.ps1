
Write-Host -ForegroundColor Yellow '██████╗ ██████╗ ███████╗     ██████╗██╗  ██╗███████╗ ██████╗██╗  ██╗'
Write-Host -ForegroundColor Yellow '██╔══██╗██╔══██╗██╔════╝    ██╔════╝██║  ██║██╔════╝██╔════╝██║ ██╔╝'
Write-Host -ForegroundColor Yellow '██████╔╝██████╔╝█████╗█████╗██║     ███████║█████╗  ██║     █████╔╝ '
Write-Host -ForegroundColor Yellow '██╔═══╝ ██╔══██╗██╔══╝╚════╝██║     ██╔══██║██╔══╝  ██║     ██╔═██╗ '
Write-Host -ForegroundColor Yellow '██║     ██║  ██║███████╗    ╚██████╗██║  ██║███████╗╚██████╗██║  ██╗'
Write-Host -ForegroundColor Yellow '╚═╝     ╚═╝  ╚═╝╚══════╝     ╚═════╝╚═╝  ╚═╝╚══════╝ ╚═════╝╚═╝  ╚═╝'
Write-Host -ForegroundColor Yellow ''
                                                           

Write-Host -ForegroundColor Yellow 'Checking if required Modules are installed.'
Install-PackageProvider -Name NuGet -Force
##--Security Principal Check for If statement--##
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())


If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    
if (Get-Module -ListAvailable -Name AADRM) {
    Write-Host -ForegroundColor Green "Azure AD Right Management Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD Right Management module"
    Install-Module -Name AADRM -force
}
Start-Sleep -Seconds 2
if (Get-Module -ListAvailable -Name AzureAD) {
    Write-Host -ForegroundColor Green "Azure AD Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD module"
    Install-Module -Name AzureAD -force
}
Start-Sleep -Seconds 2    
if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    Write-Host -ForegroundColor Green "Teams Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Teams Module"
    Install-Module -Name MicrosoftTeams -Force
}
Start-Sleep -Seconds 2
if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    Write-Host -ForegroundColor Green "SharePoint Online Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing SharePoint Online module"
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force
}
Start-Sleep -Seconds 2    
if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host -ForegroundColor Green "Microsoft Online Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Microsoft Online module"
    Install-Module -Name MSOnline -force
}
Start-Sleep -Seconds 2
if (Get-Module -ListAvailable -Name AzureRM) {
    Write-Host -ForegroundColor Green "Azure Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure module... This may take a few minutes"
    Install-Module -name AzureRM -Force
}    
   
    
}
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}

write-host -foregroundcolor Green "All Modules Installed"
Write-host 'Press any Key to return to the Menu'
##-- Waits for user input before continuing --##
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
}
}