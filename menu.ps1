## START NEW WINDOW AS ADMINISTRATOR ELEVATED ##

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}
#CHECK THE CONSOLE SETTINGS AND TURN ON THE CONSOLE WINDOW IF APPLICABLE
$BinConsole = "C:\BWApp\Bin\settings_console.dll"
$GetBinConsole = Get-Content $BinConsole
if ($GetBinConsole -like 'Enabled'){
                                    $ShowConsole = 1
                                    write-host 'Console Window Enabled'
                                    }
if ($GetBinConsole -like 'Disabled'){
                                    ## HIDE THE CONSOLE WINDOW BEFORE GUI LAUNCHES ##
                                    Add-Type -Name Window -Namespace Console -MemberDefinition '
                                    [DllImport("Kernel32.dll")]
                                    public static extern IntPtr GetConsoleWindow();
                                    [DllImport("user32.dll")]
                                    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
                                    '
                                    $consolePtr = [Console.Window]::GetConsoleWindow()
                                    [Console.Window]::ShowWindow($consolePtr, $ShowConsole)
                                    ##################
                                    }



    



#GET APP LOCATION FROM LAUNCHER TO LOCATE SCRIPTS#
$Path = Get-Content -Path "C:\Temp\BWApp\Path.txt"
Set-Location -Path $Path


#Check Version
$versioncheckurl = "https://raw.githubusercontent.com/valllem/BWApp/master/version.dll"
$versionoutput = "C:\BWapp\bin\versioncheck.dll"
(New-Object System.Net.WebClient).DownloadFile($versioncheckurl, $versionoutput)
$versionavailable = Get-Content $versionoutput
$versioncurrent = Get-Content "C:\BWApp\BWApp-master\version.dll"

#Login
$UserCredential = Get-Credential
Connect-AzureAD -Credential $UserCredential
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking
    Start-Sleep -seconds 1
    Connect-MsolService -Credential $UserCredential
}
catch {
    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
    -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"} `
    |select -First 1)
    $EXOSession = New-ExoPSSession
    Import-PSSession $EXOSession
}

$answer = [System.Windows.Forms.MessageBox]::Show(
	"A new app is available to download.
The new app is in development and currently is only usable if you use Microsoft's Partner Network.
Support for single tenant use is likely to be added sometime soon.

Did you want to install the new version? (Runs along side the current)", "A New App is available!", "YesNo", "Question", "Button1")
if ($answer -eq 'Yes')
{
	$Installerurl = "https://github.com/valllem/My365Partner/raw/main/Installer/BWApp-Installer.msi"
	$TempInstaller = "C:\Temp\BWApp-Installer.msi"
	(New-Object System.Net.WebClient).DownloadFile($Installerurl, $TempInstaller)
	Invoke-Item -Path $TempInstaller
	Start-Sleep -Seconds 5
	ii "C:\Program Files\DCNetworks\BWApp\"
}


 $logfile = "C:\BWApp\Logs\Log.txt"   
    
#[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.drawing

#Form Setup
$form1 = New-Object System.Windows.Forms.Form
$LabelStatus = New-Object system.Windows.Forms.Label
$tooltip1 = New-Object System.Windows.Forms.ToolTip
$customCommand = New-Object System.Windows.Forms.Button
$OpenLogs = New-Object System.Windows.Forms.Button



#Mailbox
$mailboxGrantFullAccess = New-Object System.Windows.Forms.Button
$mailboxRemoveFullAccess = New-Object System.Windows.Forms.Button
$mailboxGrantSendAs = New-Object System.Windows.Forms.Button
$mailboxRemoveSendAs = New-Object System.Windows.Forms.Button
$mailboxGrantForward = New-Object System.Windows.Forms.Button
$mailboxRemoveForward = New-Object System.Windows.Forms.Button
$mailboxGrantFullAccessAll = New-Object System.Windows.Forms.Button
$mailboxRemoveFullAccessAll = New-Object System.Windows.Forms.Button
$mailboxEnableOOF = New-Object System.Windows.Forms.Button
$mailboxDisableOOF = New-Object System.Windows.Forms.Button

#Calendar
$CalendarGrantAccess = New-Object System.Windows.Forms.Button
$CalendarRemoveAccess = New-Object System.Windows.Forms.Button
$CalendarGrantAccessAll = New-Object System.Windows.Forms.Button
$CalendarRemoveAccessAll = New-Object System.Windows.Forms.Button

#Security
$SecurityBlockEmail = New-Object System.Windows.Forms.Button
$SecurityBlockDomain = New-Object System.Windows.Forms.Button

#Reports
$ReportsMailboxPerms = New-Object System.Windows.Forms.Button
$ReportsAllMailboxPerms = New-Object System.Windows.Forms.Button
$ReportsAllMailboxForwards = New-Object System.Windows.Forms.Button
$ReportsAllDistGroupMembers = New-Object System.Windows.Forms.Button
$ReportsMailLogs = New-Object System.Windows.Forms.Button
$ReportsOverview = New-Object System.Windows.Forms.Button
$ReportsLicensed = New-Object System.Windows.Forms.Button
$ReportsEmailAddresses = New-Object System.Windows.Forms.Button

#Accounts
$AccountsChangeUPN = New-Object System.Windows.Forms.Button
$AccountsChangeNames = New-Object System.Windows.Forms.Button
$AccountsDisable = New-Object System.Windows.Forms.Button
$AccountsEnable = New-Object System.Windows.Forms.Button

#Tenant
$SecurityPrepareTenancy = New-Object System.Windows.Forms.Button
$SecurityEnableAuditLogging = New-Object System.Windows.Forms.Button


#Settings
$ConsoleWindow = New-Object system.Windows.Forms.CheckBox
$CheckUpdates = New-Object System.Windows.Forms.Button
$UpdateModules = New-Object System.Windows.Forms.Button


#Tabs
$TabControl = New-object System.Windows.Forms.TabControl
$Mailbox = New-Object System.Windows.Forms.TabPage
$Calendar = New-Object System.Windows.Forms.TabPage
$Security = New-Object System.Windows.Forms.TabPage
$Reports = New-Object System.Windows.Forms.TabPage
$Accounts = New-Object System.Windows.Forms.TabPage
$Tenant = New-Object System.Windows.Forms.TabPage
$Settings = New-Object System.Windows.Forms.TabPage


$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState


#Form Parameter
$form1.Text = "BWApp"
$form1.Name = "form1"
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 650
$System_Drawing_Size.Height = 400
$form1.ClientSize = $System_Drawing_Size

$LabelStatus.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 9
$System_Drawing_Point.Y = 380
$LabelStatus.Location = $System_Drawing_Point
$LabelStatus.Name = "LabelStatus"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 10
$System_Drawing_Size.Width = 300
$LabelStatus.Size = $System_Drawing_Size
$LabelStatus.text                = "Status: READY"
$LabelStatus.AutoSize            = $true
$LabelStatus.Font                = 'Microsoft Sans Serif,10'
$LabelStatus.ForeColor           = "#7ed321"
$form1.Controls.Add($LabelStatus)


#customCommand
$customCommand                = New-Object system.Windows.Forms.Button
$customCommand.text           = "Custom Commands"
$customCommand.width          = 160
$customCommand.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 5
$customCommand.location       = $System_Drawing_Point
$customCommand.Font           = 'Microsoft Sans Serif,10,style=Bold'
$customCommand.ForeColor      = "#9013fe"
$customCommand.add_Click({write-host 'Custom Command'})
$form1.Controls.Add($customCommand)
$customCommand.Add_Click({
$pi = New-Object system.Diagnostics.ProcessStartInfo
$pi.FileName = "powershell.exe"
$pi.Arguments = "-NoExit -noprofile -command C:\BWApp\BWApp-master\CustomCommand.ps1"
[system.Diagnostics.Process]::Start($pi)

})

#OpenLogs
$OpenLogs                = New-Object system.Windows.Forms.Button
$OpenLogs.text           = "Open Log Folder"
$OpenLogs.width          = 160
$OpenLogs.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 5
$OpenLogs.location       = $System_Drawing_Point
$OpenLogs.Font           = 'Microsoft Sans Serif,10,style=Bold'
$OpenLogs.ForeColor      = "#9013fe"
$OpenLogs.add_Click({
write-host 'Log Folder'
ii C:\BWApp\Logs\})
$form1.Controls.Add($OpenLogs)

#SignOut
$SignOut                = New-Object system.Windows.Forms.Button
$SignOut.text           = "Sign Out & Exit"
$SignOut.width          = 160
$SignOut.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 440
$System_Drawing_Point.Y = 5
$SignOut.location       = $System_Drawing_Point
$SignOut.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SignOut.ForeColor      = "#d0021b"
$SignOut.add_Click({
.\SignOut.ps1
$form1.Close()
})
$form1.Controls.Add($SignOut)



#Tab Control 
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 85
$tabControl.Location = $System_Drawing_Point
$tabControl.Name = "tabControl"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 275
$System_Drawing_Size.Width = 600
$tabControl.Size = $System_Drawing_Size
$form1.Controls.Add($tabControl)

#Mailbox Page
$Mailbox.DataBindings.DefaultDataSourceUpdateMode = 0
$Mailbox.UseVisualStyleBackColor = $True
$Mailbox.Name = "Mailbox Tasks"
$Mailbox.Text = "Mailbox Tasks”
$tabControl.Controls.Add($Mailbox)

#Calendar Page
$Calendar.DataBindings.DefaultDataSourceUpdateMode = 0
$Calendar.UseVisualStyleBackColor = $True
$Calendar.Name = "Calendar Tasks"
$Calendar.Text = "Calendar Tasks”
$tabControl.Controls.Add($Calendar)

#Security Page
$Security.DataBindings.DefaultDataSourceUpdateMode = 0
$Security.UseVisualStyleBackColor = $True
$Security.Name = "Security Tasks"
$Security.Text = "Security Tasks”
$tabControl.Controls.Add($Security)

#Reports Page
$Reports.DataBindings.DefaultDataSourceUpdateMode = 0
$Reports.UseVisualStyleBackColor = $True
$Reports.Name = "Reports"
$Reports.Text = "Reports Tasks”
$tabControl.Controls.Add($Reports)

#Accounts Page
$Accounts.DataBindings.DefaultDataSourceUpdateMode = 0
$Accounts.UseVisualStyleBackColor = $True
$Accounts.Name = "Accounts"
$Accounts.Text = "Accounts Tasks”
$tabControl.Controls.Add($Accounts)

#Tenant Page
$Tenant.DataBindings.DefaultDataSourceUpdateMode = 0
$Tenant.UseVisualStyleBackColor = $True
$Tenant.Name = "Tenant Tasks"
$Tenant.Text = "Tenant Tasks”
$tabControl.Controls.Add($Tenant)

#Settings Page
$Settings.DataBindings.DefaultDataSourceUpdateMode = 0
$Settings.UseVisualStyleBackColor = $True
$Settings.Name = "Settings"
$Settings.Text = "App Settings”
$tabControl.Controls.Add($Settings)













$OnLoadForm_StateCorrection=
{
    $form1.WindowState = $InitialFormWindowState
}


########################MAILBOX########################
$mailboxGrantFullAccess_RunOnClick={write-host 'Grant Mailbox Full Access'}
$mailboxRemoveFullAccess_RunOnClick={write-host 'Remove Mailbox Full Access'}
$mailboxGrantSendAs_RunOnClick={write-host 'Grant Mailbox Send As'}
$mailboxRemoveSendAs_RunOnClick={write-host 'Remove Mailbox Send As'}
$mailboxGrantForward_RunOnClick={write-host 'Create Forward'}
$mailboxRemoveForward_RunOnClick={write-host 'Remove Forward'}
$mailboxGrantFullAccessAll_RunOnClick={write-host 'Grant Full Access to All'}
$mailboxRemoveFullAccessAll_RunOnClick={write-host 'Remove Full Access to All'}
$mailboxEnableOOF_RunOnClick={write-host 'Enable Out of Office'}
$mailboxDisableOOF_RunOnClick={write-host 'Disable Out of Office'}


#mailboxGrantFullAccess
$mailboxGrantFullAccess                = New-Object system.Windows.Forms.Button
$mailboxGrantFullAccess.text           = "Grant Full Access"
$mailboxGrantFullAccess.width          = 160
$mailboxGrantFullAccess.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$mailboxGrantFullAccess.location       = $System_Drawing_Point
$mailboxGrantFullAccess.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxGrantFullAccess.ForeColor      = "#7ed321"
$mailboxGrantFullAccess.add_Click({
$LabelStatus.text = "Status: Granting Mailbox Access"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Grant Mailbox Full Access'
.\mailboxGrantFullAccess.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxGrantFullAccess,'Grant Full Access Permission on a mailbox for a user')
$Mailbox.Controls.Add($mailboxGrantFullAccess)

#mailboxRemoveFullAccess
$mailboxRemoveFullAccess                = New-Object system.Windows.Forms.Button
$mailboxRemoveFullAccess.text           = "Remove Full Access"
$mailboxRemoveFullAccess.width          = 160
$mailboxRemoveFullAccess.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$mailboxRemoveFullAccess.location       = $System_Drawing_Point
$mailboxRemoveFullAccess.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxRemoveFullAccess.ForeColor      = "#d0021b"
$mailboxRemoveFullAccess.add_Click({
$LabelStatus.text = "Status: Removing Mailbox Access"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Mailbox Full Access'
.\mailboxRemoveFullAccess.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxRemoveFullAccess,'Remove Full Access Permission on a mailbox for a user')
$Mailbox.Controls.Add($mailboxRemoveFullAccess)

#mailboxGrantSendAs
$mailboxGrantSendAs                = New-Object system.Windows.Forms.Button
$mailboxGrantSendAs.text           = "Grant Send As"
$mailboxGrantSendAs.width          = 160
$mailboxGrantSendAs.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 80
$mailboxGrantSendAs.location       = $System_Drawing_Point
$mailboxGrantSendAs.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxGrantSendAs.ForeColor      = "#7ed321"
$mailboxGrantSendAs.add_Click({
$LabelStatus.text = "Status: Granting Send As Permission"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Grant Send As'
.\mailboxGrantSendAs.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxGrantSendAs,'Grant Send As Permission on a mailbox for a user')
$Mailbox.Controls.Add($mailboxGrantSendAs)

#mailboxRemoveSendAs
$mailboxRemoveSendAs                = New-Object system.Windows.Forms.Button
$mailboxRemoveSendAs.text           = "Remove Send As"
$mailboxRemoveSendAs.width          = 160
$mailboxRemoveSendAs.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 80
$mailboxRemoveSendAs.location       = $System_Drawing_Point
$mailboxRemoveSendAs.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxRemoveSendAs.ForeColor      = "#d0021b"
$mailboxRemoveSendAs.add_Click({
$LabelStatus.text = "Status: Removing Send As Permission"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Mailbox Send As'
.\mailboxRemoveSendAs.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxRemoveSendAs,'Remove Send As Permission on a mailbox for a user')
$Mailbox.Controls.Add($mailboxRemoveSendAs)

#mailboxGrantForward
$mailboxGrantForward                = New-Object system.Windows.Forms.Button
$mailboxGrantForward.text           = "Create Forward"
$mailboxGrantForward.width          = 160
$mailboxGrantForward.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 115
$mailboxGrantForward.location       = $System_Drawing_Point
$mailboxGrantForward.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxGrantForward.ForeColor      = "#7ed321"
$mailboxGrantForward.add_Click({
$LabelStatus.text = "Status: Forwarding Mail"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Forward Mail'
.\mailboxGrantForward.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxGrantForward,'Create mail forwarding on a mailbox')
$Mailbox.Controls.Add($mailboxGrantForward)

#mailboxRemoveForward
$mailboxRemoveForward                = New-Object system.Windows.Forms.Button
$mailboxRemoveForward.text           = "Remove Forward"
$mailboxRemoveForward.width          = 160
$mailboxRemoveForward.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 115
$mailboxRemoveForward.location       = $System_Drawing_Point
$mailboxRemoveForward.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxRemoveForward.ForeColor      = "#d0021b"
$mailboxRemoveForward.add_Click({
$LabelStatus.text = "Status: Removing Forwarding"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Forwarding'
.\mailboxRemoveForward.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxRemoveForward,'Remove mail forwarding on a mailbox')
$Mailbox.Controls.Add($mailboxRemoveForward)

#mailboxGrantFullAccessAll
$mailboxGrantFullAccessAll                = New-Object system.Windows.Forms.Button
$mailboxGrantFullAccessAll.text           = "Full Access All"
$mailboxGrantFullAccessAll.width          = 160
$mailboxGrantFullAccessAll.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 150
$mailboxGrantFullAccessAll.location       = $System_Drawing_Point
$mailboxGrantFullAccessAll.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxGrantFullAccessAll.ForeColor      = "#7ed321"
$mailboxGrantFullAccessAll.add_Click({
$LabelStatus.text = "Status: Granting Full Access to All"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Grant Full Access to All'
.\mailboxGrantFullAccessAll.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxGrantFullAccessAll,'Grant a user permission to All Mailboxes')
$Mailbox.Controls.Add($mailboxGrantFullAccessAll)

#mailboxRemoveFullAccessAll
$mailboxRemoveFullAccessAll                = New-Object system.Windows.Forms.Button
$mailboxRemoveFullAccessAll.text           = "Remove Access All"
$mailboxRemoveFullAccessAll.width          = 160
$mailboxRemoveFullAccessAll.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 150
$mailboxRemoveFullAccessAll.location       = $System_Drawing_Point
$mailboxRemoveFullAccessAll.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxRemoveFullAccessAll.ForeColor      = "#d0021b"
$mailboxRemoveFullAccessAll.add_Click({
$LabelStatus.text = "Status: Removing Full Access to All"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Full Access to All'
.\mailboxRemoveFullAccessAll.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxRemoveFullAccessAll,'Remove a users permissions to All Mailboxes')
$Mailbox.Controls.Add($mailboxRemoveFullAccessAll)

#mailboxEnableOOF
$mailboxEnableOOF                = New-Object system.Windows.Forms.Button
$mailboxEnableOOF.text           = "Enable OOF"
$mailboxEnableOOF.width          = 160
$mailboxEnableOOF.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 185
$mailboxEnableOOF.location       = $System_Drawing_Point
$mailboxEnableOOF.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxEnableOOF.ForeColor      = "#7ed321"
$mailboxEnableOOF.add_Click({
$LabelStatus.text = "Status: Enabling Out of Office Message"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Enable Out of Office Message'
.\mailboxEnableOOF.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxEnableOOF,'Enable Out of Office Message for a user')
$Mailbox.Controls.Add($mailboxEnableOOF)

#mailboxDisableOOF
$mailboxDisableOOF                = New-Object system.Windows.Forms.Button
$mailboxDisableOOF.text           = "Disable OOF"
$mailboxDisableOOF.width          = 160
$mailboxDisableOOF.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 185
$mailboxDisableOOF.location       = $System_Drawing_Point
$mailboxDisableOOF.Font           = 'Microsoft Sans Serif,10,style=Bold'
$mailboxDisableOOF.ForeColor      = "#d0021b"
$mailboxDisableOOF.add_Click({
$LabelStatus.text = "Status: Disabling Out of Office Message"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Disable Out of Office Message'
.\mailboxDisableOOF.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($mailboxDisableOOF,'Disable Out of Office Message for a user')
$Mailbox.Controls.Add($mailboxDisableOOF)




########################CALENDAR########################
####CALENDAR####
$CalendarGrantAccess_RunOnClick={write-host 'Grant Calendar Access'}
$CalendarRemoveAccess_RunOnClick={write-host 'Remove Calendar Access'}
$CalendarGrantAccessAll_RunOnClick={write-host 'Grant Calendar Access to ALL'}
$CalendarRemoveAccessAll_RunOnClick={write-host 'Remove Calendar Access to ALL'}


#CalendarGrantAccess
$CalendarGrantAccess                = New-Object system.Windows.Forms.Button
$CalendarGrantAccess.text           = "Grant Access"
$CalendarGrantAccess.width          = 160
$CalendarGrantAccess.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$CalendarGrantAccess.location       = $System_Drawing_Point
$CalendarGrantAccess.Font           = 'Microsoft Sans Serif,10,style=Bold'
$CalendarGrantAccess.ForeColor      = "#7ed321"
$CalendarGrantAccess.add_Click({
$LabelStatus.text = "Status: Granting Calendar Access"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Grant Calendar Access'
.\CalendarGrantAccess.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($CalendarGrantAccess,'Grant a user Access to another users calendar.')
$Calendar.Controls.Add($CalendarGrantAccess)

#CalendarRemoveAccess
$CalendarRemoveAccess                = New-Object system.Windows.Forms.Button
$CalendarRemoveAccess.text           = "Remove Access"
$CalendarRemoveAccess.width          = 160
$CalendarRemoveAccess.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$CalendarRemoveAccess.location       = $System_Drawing_Point
$CalendarRemoveAccess.Font           = 'Microsoft Sans Serif,10,style=Bold'
$CalendarRemoveAccess.ForeColor      = "#d0021b"
$CalendarRemoveAccess.add_Click({
$LabelStatus.text = "Status: Removing Calendar Access"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Calendar Access'
.\CalendarRemoveAccess.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($CalendarRemoveAccess,'Remove a users Access to another users calendar.')
$Calendar.Controls.Add($CalendarRemoveAccess)

#CalendarGrantAccessAll
$CalendarGrantAccessAll                = New-Object system.Windows.Forms.Button
$CalendarGrantAccessAll.text           = "Grant Access All"
$CalendarGrantAccessAll.width          = 160
$CalendarGrantAccessAll.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 80
$CalendarGrantAccessAll.location       = $System_Drawing_Point
$CalendarGrantAccessAll.Font           = 'Microsoft Sans Serif,10,style=Bold'
$CalendarGrantAccessAll.ForeColor      = "#7ed321"
$CalendarGrantAccessAll.add_Click({
$LabelStatus.text = "Status: Granting Calendar Access to All"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Grant Calendar Access to All'
.\CalendarGrantAccessAll.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($CalendarGrantAccessAll,'Grant a user access to All other Calendars')
$Calendar.Controls.Add($CalendarGrantAccessAll)

#CalendarRemoveAccessAll
$CalendarRemoveAccessAll                = New-Object system.Windows.Forms.Button
$CalendarRemoveAccessAll.text           = "Remove Access All"
$CalendarRemoveAccessAll.width          = 160
$CalendarRemoveAccessAll.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 80
$CalendarRemoveAccessAll.location       = $System_Drawing_Point
$CalendarRemoveAccessAll.Font           = 'Microsoft Sans Serif,10,style=Bold'
$CalendarRemoveAccessAll.ForeColor      = "#d0021b"
$CalendarRemoveAccessAll.add_Click({
$LabelStatus.text = "Status: Removing Calendar Access to All"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Remove Calendar Access to All'
.\CalendarRemoveAccessAll.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($CalendarRemoveAccessAll,'Remove a users access to All other Calendars')
$Calendar.Controls.Add($CalendarRemoveAccessAll)


########################SECURITY########################
####SECURITY####
$SecurityBlockEmail_RunOnClick={write-host 'Block Email'}
$SecurityBlockDomain_RunOnClick={write-host 'Block Domain'}
$SecurityPrepareTenancy_RunOnClick={write-host 'Prepare Tenancy'}
$SecurityEnableAuditLogging_RunOnClick={write-host 'Enable Audit Logs'}


#SecurityBlockEmail
$SecurityBlockEmail                = New-Object system.Windows.Forms.Button
$SecurityBlockEmail.text           = "Block Email"
$SecurityBlockEmail.width          = 160
$SecurityBlockEmail.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$SecurityBlockEmail.location       = $System_Drawing_Point
$SecurityBlockEmail.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SecurityBlockEmail.ForeColor      = "#7ed321"
$SecurityBlockEmail.add_Click({
$LabelStatus.text = "Status: Blocking Email Address"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Block Email Address'
.\SecurityBlockEmail.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($SecurityBlockEmail,'Block an Email Address in the Spam Policy')
$Security.Controls.Add($SecurityBlockEmail)

#SecurityBlockDomain
$SecurityBlockDomain                = New-Object system.Windows.Forms.Button
$SecurityBlockDomain.text           = "Block Domain"
$SecurityBlockDomain.width          = 160
$SecurityBlockDomain.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$SecurityBlockDomain.location       = $System_Drawing_Point
$SecurityBlockDomain.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SecurityBlockDomain.ForeColor      = "#d0021b"
$SecurityBlockDomain.add_Click({
$LabelStatus.text = "Status: Blocking Domain"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Block Domain'
.\SecurityBlockDomain.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($SecurityBlockDomain,'Block a Domain address in the Spam Policy')
$Security.Controls.Add($SecurityBlockDomain)



########################REPORTS########################
####REPORTS####
$ReportsMailboxPerms_RunOnClick={write-host 'Check a Mailbox Permissions'}
$ReportsAllMailboxPerms_RunOnClick={write-host 'Check All Perms'}
$ReportsAllMailboxForwards_RunOnClick={write-host 'Check all Forwards'}
$ReportsAllDistGroupMembers_RunOnClick={write-host 'All Dist Members'}
$ReportsMailLogs_RunOnClick={write-host 'Mail Log (48hrs)'}
$ReportsOverview_RunOnClick={write-host 'Mail Log (48hrs)'}

#ReportsMailboxPerms
$ReportsMailboxPerms                = New-Object system.Windows.Forms.Button
$ReportsMailboxPerms.text           = "Check Mailbox Perms"
$ReportsMailboxPerms.width          = 160
$ReportsMailboxPerms.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$ReportsMailboxPerms.location       = $System_Drawing_Point
$ReportsMailboxPerms.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsMailboxPerms.ForeColor      = "#7ed321"
$ReportsMailboxPerms.add_Click({
$LabelStatus.text = "Status: Getting Mailbox Permissions"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Mailbox Permissions'
.\ReportsMailboxPerms.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsMailboxPerms,'Get list of users that have access to a specific mailbox')
$Reports.Controls.Add($ReportsMailboxPerms)

#ReportsAllMailboxPerms
$ReportsAllMailboxPerms                = New-Object system.Windows.Forms.Button
$ReportsAllMailboxPerms.text           = "Check All Perms"
$ReportsAllMailboxPerms.width          = 160
$ReportsAllMailboxPerms.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$ReportsAllMailboxPerms.location       = $System_Drawing_Point
$ReportsAllMailboxPerms.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsAllMailboxPerms.ForeColor      = "#d0021b"
$ReportsAllMailboxPerms.add_Click({
$LabelStatus.text = "Status: Getting All Mailbox Permissions"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'All Mailbox Permissions'
.\ReportsAllMailboxPerms.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsAllMailboxPerms,'Check all mailbox permissions ')
$Reports.Controls.Add($ReportsAllMailboxPerms)

#ReportsAllMailboxForwards
$ReportsAllMailboxForwards                = New-Object system.Windows.Forms.Button
$ReportsAllMailboxForwards.text           = "Check all Forwards"
$ReportsAllMailboxForwards.width          = 160
$ReportsAllMailboxForwards.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 80
$ReportsAllMailboxForwards.location       = $System_Drawing_Point
$ReportsAllMailboxForwards.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsAllMailboxForwards.ForeColor      = "#7ed321"
$ReportsAllMailboxForwards.add_Click({
$LabelStatus.text = "Status: Getting All Mailbox Permissions"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'All Mailbox Permissions'
.\ReportsAllMailboxForwards.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsAllMailboxForwards,'Check all mailbox forwarding rules being applied and produces a list')
$Reports.Controls.Add($ReportsAllMailboxForwards)

#ReportsAllDistGroupMembers
$ReportsAllDistGroupMembers                = New-Object system.Windows.Forms.Button
$ReportsAllDistGroupMembers.text           = "All Dist Members"
$ReportsAllDistGroupMembers.width          = 160
$ReportsAllDistGroupMembers.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 80
$ReportsAllDistGroupMembers.location       = $System_Drawing_Point
$ReportsAllDistGroupMembers.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsAllDistGroupMembers.ForeColor      = "#d0021b"
$ReportsAllDistGroupMembers.add_Click({
$LabelStatus.text = "Status: Getting Distribution Groups and Members"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Get Distribution Groups and Members'
.\ReportsAllDistGroupMembers.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsAllDistGroupMembers,'List all Distribution Groups and the members in them')
$Reports.Controls.Add($ReportsAllDistGroupMembers)

#ReportsMailLogs
$ReportsMailLogs                = New-Object system.Windows.Forms.Button
$ReportsMailLogs.text           = "Mail Log (48hrs)"
$ReportsMailLogs.width          = 160
$ReportsMailLogs.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 115
$ReportsMailLogs.location       = $System_Drawing_Point
$ReportsMailLogs.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsMailLogs.ForeColor      = "#7ed321"
$ReportsMailLogs.add_Click({
$LabelStatus.text = "Status: Getting Mail Logs (last 48hrs)"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Get Mail Logs (last 48hrs)'
.\ReportsMailLogs.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsMailLogs,'Shows the last 48 hours of mail logs for all accounts')
$Reports.Controls.Add($ReportsMailLogs)

#ReportsOverview
$ReportsOverview                = New-Object system.Windows.Forms.Button
$ReportsOverview.text           = "Overview Report"
$ReportsOverview.width          = 160
$ReportsOverview.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 115
$ReportsOverview.location       = $System_Drawing_Point
$ReportsOverview.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsOverview.ForeColor      = "#d0021b"
$ReportsOverview.add_Click({
$LabelStatus.text = "Status: Generating Tenant Overview - will launch webpage (This can take a while, please wait)"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Generating Tenant Overview'
.\ReportsOverview.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsOverview,'Generates an HTML file report showing various Tenancy information such as Global admins, licensing, mailboxes, groups etc')
$Reports.Controls.Add($ReportsOverview)

#ReportsLicensed
$ReportsLicensed                = New-Object system.Windows.Forms.Button
$ReportsLicensed.text           = "Show Licensed Users"
$ReportsLicensed.width          = 160
$ReportsLicensed.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 150
$ReportsLicensed.location       = $System_Drawing_Point
$ReportsLicensed.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsLicensed.ForeColor      = "#7ed321"
$ReportsLicensed.add_Click({
$LabelStatus.text = "Status: Getting Licensed Users"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Getting Licensed Users'
.\ReportsLicensed.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsLicensed,'Get list of licensed users')
$Reports.Controls.Add($ReportsLicensed)

#ReportsEmailAddresses
$ReportsEmailAddresses                = New-Object system.Windows.Forms.Button
$ReportsEmailAddresses.text           = "Show Email Addresses"
$ReportsEmailAddresses.width          = 160
$ReportsEmailAddresses.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 150
$ReportsEmailAddresses.location       = $System_Drawing_Point
$ReportsEmailAddresses.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ReportsEmailAddresses.ForeColor      = "#7ed321"
$ReportsEmailAddresses.add_Click({
$LabelStatus.text = "Status: Getting Email Addresses"
$LabelStatus.ForeColor = "#7ed321"
write-host -ForegroundColor Cyan 'Getting Email Addresses'
.\ReportsEmailAddresses.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ReportsEmailAddresses,'Get list of Email Addresses')
$Reports.Controls.Add($ReportsEmailAddresses)

########################ACCOUNTS########################
####ACCOUNTS####
$AccountsChangeUPN_RunOnClick={write-host 'Change UPN'}
$AccountsChangeNames_RunOnClick={write-host 'Update Names'}
$AccountsDisable_RunOnClick={write-host 'Disable Account'}
$AccountsEnable_RunOnClick={write-host 'Enable Account'}


#AccountsChangeUPN
$AccountsChangeUPN                = New-Object system.Windows.Forms.Button
$AccountsChangeUPN.text           = "Change UPN"
$AccountsChangeUPN.width          = 160
$AccountsChangeUPN.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$AccountsChangeUPN.location       = $System_Drawing_Point
$AccountsChangeUPN.Font           = 'Microsoft Sans Serif,10,style=Bold'
$AccountsChangeUPN.ForeColor      = "#7ed321"
$AccountsChangeUPN.add_Click({
$LabelStatus.text = "Status: Changing user sign-in name"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Change user sign-in name'
.\AccountsChangeUPN.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($AccountsChangeUPN,'Change a users UPN (Sign in name)')
$Accounts.Controls.Add($AccountsChangeUPN)

#AccountsChangeNames
$AccountsChangeNames                = New-Object system.Windows.Forms.Button
$AccountsChangeNames.text           = "Update Names"
$AccountsChangeNames.width          = 160
$AccountsChangeNames.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$AccountsChangeNames.location       = $System_Drawing_Point
$AccountsChangeNames.Font           = 'Microsoft Sans Serif,10,style=Bold'
$AccountsChangeNames.ForeColor      = "#d0021b"
$AccountsChangeNames.add_Click({
$LabelStatus.text = "Status: Changing users Name"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Change users Name'
.\AccountsChangeNames.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($AccountsChangeNames,'Update a users First and Last name')
$Accounts.Controls.Add($AccountsChangeNames)

#AccountsDisable


#AccountsEnable
$AccountsEnable                = New-Object system.Windows.Forms.Button
$AccountsEnable.text           = "Enable Account"
$AccountsEnable.width          = 160
$AccountsEnable.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 80
$AccountsEnable.location       = $System_Drawing_Point
$AccountsEnable.Font           = 'Microsoft Sans Serif,10,style=Bold'
$AccountsEnable.ForeColor      = "#d0021b"
$AccountsEnable.add_Click({
$LabelStatus.text = "Status: Enabling Account"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Enable Account'
.\AccountsEnable.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($AccountsEnable,'Enable a user account')
$Accounts.Controls.Add($AccountsEnable)





############TENANT#####################################

#SecurityPrepareTenancy
$SecurityPrepareTenancy                = New-Object system.Windows.Forms.Button
$SecurityPrepareTenancy.text           = "Prepare Tenancy"
$SecurityPrepareTenancy.width          = 160
$SecurityPrepareTenancy.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 45
$SecurityPrepareTenancy.location       = $System_Drawing_Point
$SecurityPrepareTenancy.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SecurityPrepareTenancy.ForeColor      = "#7ed321"
$SecurityPrepareTenancy.add_Click({
$LabelStatus.text = "Status: Preparing Tenant"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Prepare Tenant'
.\SecurityPrepareTenancy.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($SecurityPrepareTenancy,'Prepares the tenancy by applying various security standards, such as Disable POP/IMAP, enabling audit logging etc')
$Tenant.Controls.Add($SecurityPrepareTenancy)

#SecurityEnableAuditLogging
$SecurityEnableAuditLogging                = New-Object system.Windows.Forms.Button
$SecurityEnableAuditLogging.text           = "Enable Audit Logs"
$SecurityEnableAuditLogging.width          = 160
$SecurityEnableAuditLogging.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 45
$SecurityEnableAuditLogging.location       = $System_Drawing_Point
$SecurityEnableAuditLogging.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SecurityEnableAuditLogging.ForeColor      = "#d0021b"
$SecurityEnableAuditLogging.add_Click({
$LabelStatus.text = "Status: Enabling Audit Logging"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Enable Audit Logging'
.\SecurityEnableAuditLogging.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($SecurityEnableAuditLogging,'Turns on audit logging in the Tenancy')
$Tenant.Controls.Add($SecurityEnableAuditLogging)

#SecurityOutsiderImpersinating
$SecurityOutsiderImpersinating                = New-Object system.Windows.Forms.Button
$SecurityOutsiderImpersinating.text           = "Outside Senders"
$SecurityOutsiderImpersinating.width          = 160
$SecurityOutsiderImpersinating.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 80
$SecurityOutsiderImpersinating.location       = $System_Drawing_Point
$SecurityOutsiderImpersinating.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SecurityOutsiderImpersinating.ForeColor      = "#d0021b"
$SecurityOutsiderImpersinating.add_Click({
$LabelStatus.text = "Status: Creating Transport Rule, please wait..."
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Creating Transport Rule, please wait...'
.\Transport-OutsideSender.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($SecurityOutsiderImpersinating,'Creates a rule that alerts the recipient if the outside sender has the same name as an internal user')
$Tenant.Controls.Add($SecurityOutsiderImpersinating)


############SETTINGS

#ConsoleWindow
$ConsoleWindow                = New-Object system.Windows.Forms.Button
$ConsoleWindow.text           = "Show Console Settings"
$ConsoleWindow.width          = 300
$ConsoleWindow.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 45
$ConsoleWindow.location       = $System_Drawing_Point
$ConsoleWindow.Font           = 'Microsoft Sans Serif,10'
$ConsoleWindow.add_Click({
$LabelStatus.text = "Status: Changing Settings"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Change Console Window'
.\ConsoleWindow.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($ConsoleWindow,'Open the Console Settings to turn on or off showing the console window')
$Settings.Controls.Add($ConsoleWindow)

#CheckUpdates
$CheckUpdates                = New-Object system.Windows.Forms.Button
$CheckUpdates.text           = "Install Updates"
$CheckUpdates.width          = 300
$CheckUpdates.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 80
$CheckUpdates.location       = $System_Drawing_Point
$CheckUpdates.Font           = 'Microsoft Sans Serif,10,style=Bold'
$CheckUpdates.ForeColor      = "#d0021b"
$CheckUpdates.add_Click({
$LabelStatus.text = "Status: Checking for Updates"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Check for Updates'
.\UpdateApp.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($CheckUpdates,'Install BWApp Updates')
$settings.Controls.Add($CheckUpdates)

#UpdateModules
$UpdateModules                = New-Object system.Windows.Forms.Button
$UpdateModules.text           = "Update Modules"
$UpdateModules.width          = 300
$UpdateModules.height         = 30
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 115
$UpdateModules.location       = $System_Drawing_Point
$UpdateModules.Font           = 'Microsoft Sans Serif,10,style=Bold'
$UpdateModules.ForeColor      = "#d0021b"
$UpdateModules.add_Click({
$LabelStatus.text = "Status: Checking for Module Updates"
$LabelStatus.ForeColor = "#f5a623"
write-host -ForegroundColor Cyan 'Check for Module Updates'
.\UpdateModules.ps1
$LabelStatus.text = "Status: Ready"
$LabelStatus.ForeColor = "#7ed321"})
$tooltip1.SetToolTip($UpdateModules,'Install and Update Modules')
$settings.Controls.Add($UpdateModules)

if ($versionavailable -notmatch $versioncurrent){
$LabelStatus.text                = "Status: UPDATE AVAILABLE"
$LabelStatus.ForeColor      = "#d0021b"
}
elseif ($versionavailable -match $versioncurrent) {
$LabelStatus.text                = "Status: Ready"
}


#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null
 #End function CreateForm
 



#Call the Function

#CreateForm