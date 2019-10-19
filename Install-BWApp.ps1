## CHECK IF EXECUTION POLICY IS REMOTESIGNED ##
If ((Get-ExecutionPolicy) -ne "RemoteSigned") {    
    If ((Get-ExecutionPolicy) -ne "unrestricted") {   
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host "                ==========================================================================================="
    Write-Host "                 Please open another powershell window as administrator and type the following command..."
    Write-Host "                "
    Write-Host -ForegroundColor red "                Set-ExecutionPolicy RemoteSigned -Force"
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host -ForegroundColor Green "                If you have already done this, press Enter to begin installation"
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

## HIDE THE CONSOLE WINDOW BEFORE GUI LAUNCHES ##
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
##################

## ENABLE THE GUI ##
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

## PROGRESS BAR INSTALLING APP ##
    $ObjForm = New-Object System.Windows.Forms.Form
	$ObjForm.Text = "Installing"
	$ObjForm.Height = 100
	$ObjForm.Width = 500
	$ObjForm.BackColor = "White"

	$ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

	## -- Create The Label
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Text = "Installing App. Please wait ... "
	$ObjLabel.Left = 5
	$ObjLabel.Top = 10
	$ObjLabel.Width = 500 - 20
	$ObjLabel.Height = 15
	$ObjLabel.Font = "Tahoma"
	## -- Add the label to the Form
	$ObjForm.Controls.Add($ObjLabel)

	$PB = New-Object System.Windows.Forms.ProgressBar
	$PB.Name = "PowerShellProgressBar"
	$PB.Value = 10
	$PB.Style="Continuous"

	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Width = 500 - 40
	$System_Drawing_Size.Height = 20
	$PB.Size = $System_Drawing_Size
	$PB.Left = 5
	$PB.Top = 40
	$ObjForm.Controls.Add($PB)

	## -- Show the Progress-Bar and Start The PowerShell Script
	$ObjForm.Show() | Out-Null
	$ObjForm.Focus() | Out-NUll
	$ObjLabel.Text = "Installing App. Please wait ..."
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 1
	$ObjLabel.Text = "Installing App. Getting Ready ..."
	Start-Sleep -Milliseconds 300
 

## GETTING FILES FROM GITHUB ##
New-Item -ItemType Directory -Force -Path "C:\BWApp"
$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp_$Version.zip”)
###########    



    $ObjForm.Refresh()
    $PB.Value = 20
	$ObjLabel.Text = "Checking Required Modules ..."
	Start-Sleep -Seconds 2
 
################################## SECOND PROGRESS BAR ############################

    ## PROGRESS BAR INSTALLING APP ##
    $ObjForm2 = New-Object System.Windows.Forms.Form
	$ObjForm2.Text = "Installing Modules"
	$ObjForm2.Height = 100
	$ObjForm2.Width = 500
	$ObjForm2.BackColor = "White"

	$ObjForm2.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	##$ObjForm2.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ObjForm2.Location.X = 500
    $ObjForm2.Location.Y = 500

	## -- Create The Label
	$ObjLabel2 = New-Object System.Windows.Forms.Label
	$ObjLabel2.Text = "Installing Modules ..."
	$ObjLabel2.Left = 5
	$ObjLabel2.Top = 10
	$ObjLabel2.Width = 500 - 20
	$ObjLabel2.Height = 15
	$ObjLabel2.Font = "Tahoma"
	## -- Add the label to the Form
	$ObjForm2.Controls.Add($ObjLabel2)

	$PB2 = New-Object System.Windows.Forms.ProgressBar
	$PB2.Name = "PowerShellProgressBar"
	$PB2.Value = 0
	$PB2.Style="Continuous"

	$System_Drawing_Size2 = New-Object System.Drawing.Size
	$System_Drawing_Size2.Width = 500 - 40
	$System_Drawing_Size2.Height = 20
	$PB2.Size = $System_Drawing_Size2
	$PB2.Left = 5
	$PB2.Top = 40
	$ObjForm2.Controls.Add($PB2)

	## -- Show the Progress-Bar and Start The PowerShell Script
	$ObjForm2.Show() | Out-Null
	$ObjForm2.Focus() | Out-NUll
	$ObjLabel2.Text = "Installing Modules ..."
	$ObjForm2.Refresh()
	Start-Sleep -Milliseconds 300
    
    $ObjForm2.Refresh()
    $PB2.Value = 1
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

####################################################################################
 
                                                           

## INSTALL REQUIRED MODULES ##
Install-PackageProvider -Name NuGet -Force
##--Security Principal Check for If statement--##
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())


If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    
    $ObjForm2.Refresh()
    $PB2.Value = 10
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name AADRM) {
    
    
} 
else {
    
    Install-Module -Name AADRM -force
}

    $ObjForm2.Refresh()
    $PB2.Value = 20
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name AzureAD) {
    

} 
else {
    
    Install-Module -Name AzureAD -force
}
    $ObjForm2.Refresh()
    $PB2.Value = 30
	$ObjLabel2.Text = "Installing Modules, this may take several minutes..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    
} 
else {
    
    Install-Module -Name MicrosoftTeams -Force
}

    $ObjForm2.Refresh()
    $PB2.Value = 60
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    
} 
else {
   
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force
}

    $ObjForm2.Refresh()
    $PB2.Value = 70
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300
  
if (Get-Module -ListAvailable -Name MSOnline) {
    
} 
else {
    
    Install-Module -Name MSOnline -force
}

    $ObjForm2.Refresh()
    $PB2.Value = 80
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name AzureRM) {
    
} 
else {
    
    Install-Module -name AzureRM -Force
}    

    $ObjForm2.Refresh()
    $PB2.Value = 90
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300

if (Get-Module -ListAvailable -Name CreateExoPsSession) {
    
} 
else {
    
    Install-Script -Name CreateExoPsSession -force 
}

    $ObjForm2.Refresh()
    $PB2.Value = 100
	$ObjLabel2.Text = "Installing Modules ..."
	Start-Sleep -Milliseconds 300
    
    Install-Module -Name Microsoft.Exchange.Management.ExoPowershellModule -Force -AllowClobber
    Install-Script -Name Load-ExchangeMFA -Force 
    Install-Module -Name ExchangeOnlineShell -Force
    winrm qc -transport:https -Force
    Clear-Host

}

    $ObjForm2.Refresh()
    $PB2.Value = 100
	$ObjLabel2.Text = "All Modules Installed"
	Start-Sleep -Seconds 2
    
    $ObjForm2.Close()




    $ObjForm.Refresh()
    $PB.Value = 40
	$ObjLabel.Text = "Importing Required Scripts..."
	Start-Sleep -Milliseconds 300

## INSTALL CLICKONCE ##
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

    $ObjForm.Refresh()
    $PB.Value = 60
	$ObjLabel.Text = "Importing Required Scripts..."
	Start-Sleep -Milliseconds 300


## CLICKONECE INSTALL CHECK ##
function Get-ClickOnce {
[CmdletBinding()]  
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    $InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
    return $InstalledApplicationNotMSI | ? { $_.displayname -match $ApplicationName } | Select-Object -First 1
}

    $ObjForm.Refresh()
    $PB.Value = 65
	$ObjLabel.Text = "Testing Scripts..."
	Start-Sleep -Milliseconds 300

Function Test-ClickOnce {
[CmdletBinding()] 
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    return ( (Get-ClickOnce -ApplicationName $ApplicationName) -ne $null) 
}

    $ObjForm.Refresh()
    $PB.Value = 70
	$ObjLabel.Text = "Creating ClickOnce Uninstall"
	Start-Sleep -Milliseconds 300


## CREATE UNINSTALL CLICKONCE ##
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


    $ObjForm.Refresh()
    $PB.Value = 75
	$ObjLabel.Text = "Loading Exchange MFA Module"
	Start-Sleep -Milliseconds 300

## LOAD EXCHANGE MFA MODULE ##
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


    $ObjForm.Refresh()
    $PB.Value = 80
	$ObjLabel.Text = "Preparing ClickOnce Script ..."
	Start-Sleep -Milliseconds 300


if ((Test-ClickOnce -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false)  {
   Install-ClickOnce -Manifest "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
}
#Load the Module
$script = Load-ExchangeMFAModule -Verbose
#Dot Source the associated script
. $Script


    $ObjForm.Refresh()
    $PB.Value = 90
	$ObjLabel.Text = "Downloading Components ..."
	Start-Sleep -Milliseconds 300
    
#Download file
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Sleep -Seconds 2    
# Unzip the Archive

    $ObjForm.Refresh()
    $PB.Value = 95
	$ObjLabel.Text = "Extracting Components ..."
	Start-Sleep -Milliseconds 300

Expand-Archive $output -DestinationPath $Path -Force
    
#Set the environment variable
##$Home = [IO.Path]::Combine($Path, "BWApp")
   
##[Environment]::SetEnvironmentVariable("HOME", "$Home", "User")

    $ObjForm.Refresh()
    $PB.Value = 99
	$ObjLabel.Text = "Creating Shortcuts..."
	Start-Sleep -Milliseconds 300


write-host -foregroundcolor Yellow "Creating Shortcuts"
cd $HOME
cd desktop
$ShortCutDir = Get-Location

Start-Sleep -Seconds 2
function set-shortcut {
param ( [string]$SourceLnk, [string]$DestinationPath )
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
    $Shortcut.Save()
    }
set-shortcut "$ShortcutDir\BWApp.lnk" "$Path\BWApp-master\Launcher.ps1"


    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Installation Complete"
	Start-Sleep -Seconds 3

    Clear-Host

    [Console.Window]::ShowWindow($consolePtr, 1)
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