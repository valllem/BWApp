## CHECK IF EXECUTION POLICY IS REMOTESIGNED ##
If ((Get-ExecutionPolicy) -ne "RemoteSigned") {    
    If ((Get-ExecutionPolicy) -ne "unrestricted") {   
    
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""

Write-Host "                                ██████╗ ██╗    ██╗ █████╗ ██████╗ ██████╗             "
Write-Host "                                ██╔══██╗██║    ██║██╔══██╗██╔══██╗██╔══██╗            "
Write-Host "                                ██████╔╝██║ █╗ ██║███████║██████╔╝██████╔╝            "
Write-Host "                                ██╔══██╗██║███╗██║██╔══██║██╔═══╝ ██╔═══╝             "
Write-Host "                                ██████╔╝╚███╔███╔╝██║  ██║██║     ██║                 "
Write-Host "                                ╚═════╝  ╚══╝╚══╝ ╚═╝  ╚═╝╚═╝     ╚═╝                 "
Write-Host "                                                                                      "
Write-Host "                ██╗███╗   ██╗███████╗████████╗ █████╗ ██╗     ██╗     ███████╗██████╗ "
Write-Host "                ██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║     ██║     ██╔════╝██╔══██╗"
Write-Host "                ██║██╔██╗ ██║███████╗   ██║   ███████║██║     ██║     █████╗  ██████╔╝"
Write-Host "                ██║██║╚██╗██║╚════██║   ██║   ██╔══██║██║     ██║     ██╔══╝  ██╔══██╗"
Write-Host "                ██║██║ ╚████║███████║   ██║   ██║  ██║███████╗███████╗███████╗██║  ██║"
Write-Host "                ╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝╚═╝  ╚═╝"
Write-Host ""
Write-Host ""                                                                                      
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "     ==========================================================================================="
Write-Host "        Please open another powershell window as administrator and type the following command..."
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "                             Set-ExecutionPolicy RemoteSigned -Force                             " -ForegroundColor Red
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "             Once done or if you have already done this, press Enter to begin installation" -ForegroundColor Green



    ## PAUSE SCRIPT UNTIL KEY PRESSED ##
    $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
    $HOST.UI.RawUI.Flushinputbuffer()


}
}
else {

}

## RUN AS ELEVATED WINDOW ##
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}






Write-Progress -Activity "Installing BWAPP" -Status "Getting Ready" -PercentComplete 1
 


New-Item -ItemType Directory -Force -Path "C:\BWApp"
$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
    


Write-Progress -Activity "Installing BWAPP" -Status "Checking Required Modules" -PercentComplete 5

                                                           


Install-PackageProvider -Name NuGet -Force
##--Security Principal Check for If statement--##
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

Clear-Host
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    
if (Get-Module -ListAvailable -Name AADRM) {
    
    
} 
else {
    Clear-Host
    Install-Module -Name AADRM -force
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 8
if (Get-Module -ListAvailable -Name AzureAD) {
    

} 
else {
    Clear-Host
    Install-Module -Name AzureAD -force
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 10  
if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    
} 
else {
    Clear-Host
    Install-Module -Name MicrosoftTeams -Force
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 15
if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    
} 
else {
    Clear-Host
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 20    
if (Get-Module -ListAvailable -Name MSOnline) {
    
} 
else {
    Clear-Host
    Install-Module -Name MSOnline -force
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 25
if (Get-Module -ListAvailable -Name AzureRM) {
    
} 
else {
    Clear-Host
    Install-Module -name AzureRM -Force
}    
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 30
if (Get-Module -ListAvailable -Name CreateExoPsSession) {
    
} 
else {
    Clear-Host
    Install-Script -Name CreateExoPsSession -force 
}
Write-Progress -Activity "Installing BWAPP" -Status "Installing Required Modules - This can take several minutes" -PercentComplete 35

    
    Install-Module -Name Microsoft.Exchange.Management.ExoPowershellModule -Force -AllowClobber
    Install-Script -Name Load-ExchangeMFA -Force 
    Install-Module -Name ExchangeOnlineShell -Force
    winrm qc -transport:https -Force
    Clear-Host

}
Write-Progress -Activity "Installing BWAPP" -Status "Installed Required Modules" -PercentComplete 40
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Importing needed scripts" -PercentComplete 41
function Install-ClickOnce {
[CmdletBinding()] 
Param(
    $Manifest = "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application",
    #AssertApplicationRequirements
    $ElevatePermissions = $true
)
    Try { 
        Add-Type -AssemblyName System.Deployment
        
        Write-Verbose "Start installation of ClockOnce Application $Manifest "

        $RemoteURI = [URI]::New( $Manifest , [UriKind]::Absolute)
        if (-not  $Manifest)
        {
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }

        $HostingManager = New-Object System.Deployment.Application.InPlaceHostingManager -ArgumentList $RemoteURI , $False
    
        #register an event to trigger custom event (yep, its a hack)
        Register-ObjectEvent -InputObject $HostingManager -EventName GetManifestCompleted -Action { 
            new-event -SourceIdentifier "ManifestDownloadComplete"
        } | Out-Null
        #register an event to trigger custom event (yep, its a hack)
        Register-ObjectEvent -InputObject $HostingManager -EventName DownloadApplicationCompleted -Action { 
            new-event -SourceIdentifier "DownloadApplicationCompleted"
        } | Out-Null

        #get the Manifest
        $HostingManager.GetManifestAsync()

        #Waitfor up to 5s for our custom event
        $event = Wait-Event -SourceIdentifier "ManifestDownloadComplete" -Timeout 5
        if ($event ) {
            $event | Remove-Event
            Write-Verbose "ClickOnce Manifest Download Completed"

            $HostingManager.AssertApplicationRequirements($ElevatePermissions)
            #todo :: can this fail ?
            
            #Download Application
            $HostingManager.DownloadApplicationAsync()
            #register and wait for completion event
            # $HostingManager.DownloadApplicationCompleted
            $event = Wait-Event -SourceIdentifier "DownloadApplicationCompleted" -Timeout 15
            if ($event ) {
                $event | Remove-Event
                Write-Verbose "ClickOnce Application Download Completed"
            } else {
                Write-error "ClickOnce Application Download did not complete in time (15s)"
            }
        } else {
           Write-error "ClickOnce Manifest Download did not complete in time (5s)"
        }

        #Clean Up
    } finally {
        #get rid of our eventhandlers
        Get-EventSubscriber|? {$_.SourceObject.ToString() -eq 'System.Deployment.Application.InPlaceHostingManager'} | Unregister-Event
    }
}
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Importing needed scripts" -PercentComplete 60
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
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Testing Scripts" -PercentComplete 65
Function Test-ClickOnce {
[CmdletBinding()] 
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    return ( (Get-ClickOnce -ApplicationName $ApplicationName) -ne $null) 
}
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Creating UnInstall" -PercentComplete 70
<# Simple UnInstall
#>
function Uninstall-ClickOnce {
[CmdletBinding()] 
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    $app=Get-ClickOnce -ApplicationName $ApplicationName

    #Deinstall One to remove all instances
    if ($App) { 
        $selectedUninstallString = $App.UninstallString 
        #Seperate cmd from parameters (First Space)
        $parts = $selectedUninstallString.Split(' ', 2)
        Start-Process -FilePath $parts[0] -ArgumentList $parts[1] -Wait 
        #ToDo : Automatic press of OK
        #Start-Sleep 5
        #$wshell = new-object -com wscript.shell
        #$wshell.sendkeys("`"OK`"~")

        $app=Get-ClickOnce -ApplicationName $ApplicationName
        if ($app) {
            Write-verbose 'De-installation aborted'
            #return $false
        } else {
            Write-verbose 'De-installation completed'
            #return $true
        } 
        
    } else {
        #return $null
    }
}
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Loading Exchange MFA Module" -PercentComplete 75
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
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Preparing ClickOnce Script" -PercentComplete 80
if ((Test-ClickOnce -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false)  {
   Install-ClickOnce -Manifest "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
}
#Load the Module
$script = Load-ExchangeMFAModule -Verbose
#Dot Source the associated script
. $Script
Clear-Host
Write-Progress -Activity "Installing BWAPP" -Status "Downloading Components" -PercentComplete 90
    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Sleep -Seconds 2    
# Unzip the Archive
Write-Progress -Activity "Installing BWAPP" -Status "Extracting Components" -PercentComplete 95
Expand-Archive $output -DestinationPath $Path -Force
    
#Set the environment variable
##$Home = [IO.Path]::Combine($Path, "BWApp")
Clear-Host    
##[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")
Write-Progress -Activity "Installing BWAPP" -Status "Creating Shortcuts" -PercentComplete 99
Start-Sleep -Seconds 4
write-host -foregroundcolor Yellow "Creating Shortcuts"
cd $HOME
cd desktop
$ShortCutDir = Get-Location
write-host $ShortCutDir
Start-Sleep -Seconds 2


$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$ShortCutDir\BWApp.lnk")
$Shortcut.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
$argA = """C:\BWApp\BWApp-master\Launcher.ps1"""
##$argB = """/S:Search Card"""
$Shortcut.Arguments = $argA
$Shortcut.Save()

Write-Progress -Activity "Installing BWAPP" -Status "Ready" -Completed
Clear-Host

    
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host "     ██╗███╗   ██╗███████╗████████╗ █████╗ ██╗     ██╗      █████╗ ████████╗██╗ ██████╗ ███╗   ██╗"
Write-Host "     ██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║     ██║     ██╔══██╗╚══██╔══╝██║██╔═══██╗████╗  ██║"
Write-Host "     ██║██╔██╗ ██║███████╗   ██║   ███████║██║     ██║     ███████║   ██║   ██║██║   ██║██╔██╗ ██║"
Write-Host "     ██║██║╚██╗██║╚════██║   ██║   ██╔══██║██║     ██║     ██╔══██║   ██║   ██║██║   ██║██║╚██╗██║"
Write-Host "     ██║██║ ╚████║███████║   ██║   ██║  ██║███████╗███████╗██║  ██║   ██║   ██║╚██████╔╝██║ ╚████║"
Write-Host "     ╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝"
Write-Host "                                                                                             "
Write-Host "      ██████╗ ██████╗ ███╗   ███╗██████╗ ██╗     ███████╗████████╗███████╗                        "
Write-Host "     ██╔════╝██╔═══██╗████╗ ████║██╔══██╗██║     ██╔════╝╚══██╔══╝██╔════╝                        "
Write-Host "     ██║     ██║   ██║██╔████╔██║██████╔╝██║     █████╗     ██║   █████╗                          "
Write-Host "     ██║     ██║   ██║██║╚██╔╝██║██╔═══╝ ██║     ██╔══╝     ██║   ██╔══╝                          "
Write-Host "     ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║     ███████╗███████╗   ██║   ███████╗                        "
Write-Host "      ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚══════╝╚══════╝   ╚═╝   ╚══════╝                        "
Write-Host "                                                                                             "
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "                             Press any key to Close         "


## WAIT FOR ANY KEY PRESS (OR IF CONSOLE HIDDEN, CLOSE WINDOW)

    $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
    $HOST.UI.RawUI.Flushinputbuffer()