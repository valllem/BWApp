if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}

Start-Sleep -Seconds 2
Write-Host "Preparing to Install"
Set-ExecutionPolicy -Scope CurrentUser Unrestricted -Force
Start-Sleep -Seconds 2





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
    Write-Host -ForegroundColor Green "Exchange PS Module exists"
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Exchange PS module"
    Install-Script -Name CreateExoPsSession -force 
}


    write-host -foregroundcolor Yellow "Installing latest Exchange Management module..."
    Install-Module -Name Microsoft.Exchange.Management.ExoPowershellModule -Force -AllowClobber



    write-host -foregroundcolor Yellow "Installing latest Exchange MFA module"
    Install-Script -Name Load-ExchangeMFA -Force 
    Install-Module -Name ExchangeOnlineShell -Force


    

}


else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}

write-host -foregroundcolor Green "All Modules Installed"


write-host -foregroundcolor Yellow "Downloading required components..."
Start-Sleep -Seconds 2



<# Simple Install Check
#>
function Get-ClickOnce {
[CmdletBinding()]  
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    $InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
    return $InstalledApplicationNotMSI | ? { $_.displayname -match $ApplicationName } | Select-Object -First 1
}


Function Test-ClickOnce {
[CmdletBinding()] 
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    return ( (Get-ClickOnce -ApplicationName $ApplicationName) -ne $null) 
}



Function Load-ExchangeMFAModule { 
[CmdletBinding()] 
Param ()
    $Modules = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
    if ($Modules.Count -ne 1 ) {
        throw "No or Multiple Modules found : Count = $($Modules.Count )"  
    }  else {
        $ModuleName =  Join-path $Modules[0].Directory.FullName "Microsoft.Exchange.Management.ExoPowershellModule.dll"
        Write-Verbose "Start Importing MFA Module"
        if ($PSVersionTable.PSVersion -ge "5.0")  { 
            Import-Module -FullyQualifiedName $ModuleName  -Force 
        } else { 
            #in case -FullyQualifiedName is not supported
            Import-Module $ModuleName  -Force 
        }

        $ScriptName =  Join-path $Modules[0].Directory.FullName "CreateExoPSSession.ps1"
        if (Test-Path $ScriptName) {
            return $ScriptName
<#
            # Load the script to add the additional commandlets (Connect-EXOPSSession)
            # DotSourcing does not work from inside a function (. $ScriptName)
            #Therefore load the script as a dynamic module instead
 
            $content = Get-Content -Path $ScriptName -Raw -ErrorAction Stop
            #BugBug >> $PSScriptRoot is Blank :-(
<#
            $PipeLine = $Host.Runspace.CreatePipeline()
            $PipeLine.Commands.AddScript(". $scriptName")
            $r = $PipeLine.Invoke()
#Err : Pipelines cannot be run concurrently.
 
            $scriptBlock = [scriptblock]::Create($content)
            New-Module -ScriptBlock $scriptBlock -Name "Microsoft.Exchange.Management.CreateExoPSSession.ps1" -ReturnResult -ErrorAction SilentlyContinue
#>

        } else {
            throw "Script not found"
            return $null
        }
    }
}


if ((Test-ClickOnce -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false)  {
   Install-ClickOnce -Manifest "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
}
#Load the Module
$script = Load-ExchangeMFAModule -Verbose
#Dot Source the associated script
. $Script

#make sure the Exchange session uses the same proxy settings as IE/Edge
$ProxySetting = New-PSSessionOption -ProxyAccessType IEConfig
 



    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Sleep -Seconds 2    
# Unzip the Archive
write-host -foregroundcolor Yellow "Extracting..."
Expand-Archive $output -DestinationPath $Path -Force
Start-Sleep -Seconds 2    
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
read-host -Prompt '               Press Enter to Close Installer'
