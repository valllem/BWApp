﻿## START NEW WINDOW AS ADMINISTRATOR ELEVATED ##

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -WindowStyle Minimized -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}
###################

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

#################

## PROGRESS BAR LOADING APP ##
    $ObjForm = New-Object System.Windows.Forms.Form
	$ObjForm.Text = "Starting App"
	$ObjForm.Height = 100
	$ObjForm.Width = 500
	$ObjForm.BackColor = "White"

	$ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

	## -- Create The Label
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Text = "Starting App. Please wait ... "
	$ObjLabel.Left = 5
	$ObjLabel.Top = 10
	$ObjLabel.Width = 500 - 20
	$ObjLabel.Height = 15
	$ObjLabel.Font = "Tahoma"
	## -- Add the label to the Form
	$ObjForm.Controls.Add($ObjLabel)

	$PB = New-Object System.Windows.Forms.ProgressBar
	$PB.Name = "PowerShellProgressBar"
	$PB.Value = 0
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
	$ObjLabel.Text = "Please Log In"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 25
	$ObjLabel.Text = "Logging In"
	Start-Sleep -Milliseconds 300
    

#################

## GET APP LOCATION FROM LAUNCHER TO LOCATE SCRIPTS ##
    $Path = Get-Content -Path "C:\Temp\BWApp\Path.txt"
    Set-Location -Path $Path
#################

## LOGIN WITH BASIC AUTH ##
    $UserCredential = Get-Credential

#################

## PROGRESS BAR CONNECTING TO AZURE AD ##
    $ObjForm.Refresh()
    $PB.Value = 50
	$ObjLabel.Text = "Connecting to AzureAD"
	Start-Sleep -Milliseconds 300

################

## CONNECT TO AZUREAD ##
    Connect-AzureAD -Credential $UserCredential

###############

## START THE LOGIN PROCESS AND CHECK IF MODERN AUTH IS REQUIRED ##
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking

    Start-Sleep -seconds 1

    $ObjForm.Refresh()
    $PB.Value = 75
	$ObjLabel.Text = "Connecting to MsolService"
	Start-Sleep -Milliseconds 300

    Connect-MsolService -Credential $UserCredential
}
catch {

    $ObjForm.Refresh()
    $PB.Value = 75
	$ObjLabel.Text = "Connecting using Modern Auth..."
	Start-Sleep -Milliseconds 300

    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
    -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"} `
    |select -First 1)
    $EXOSession = New-ExoPSSession
    Import-PSSession $EXOSession
}

    $ObjForm.Refresh()
    $PB.Value = 99
	$ObjLabel.Text = "Starting..."
	Start-Sleep -Milliseconds 300

## CLOSE PROGRESS BAR ##
$ObjForm.Close()



## LOAD MAIN FORM ##

$BWApp                           = New-Object system.Windows.Forms.Form
$BWApp.ClientSize                = '530,722'
$BWApp.text                      = "BWApp - Main Menu"
$BWApp.TopMost                   = $false


$ButtonCalendarAccess1           = New-Object system.Windows.Forms.Button
$ButtonCalendarAccess1.text      = "Give Access"
$ButtonCalendarAccess1.width     = 116
$ButtonCalendarAccess1.height    = 30
$ButtonCalendarAccess1.location  = New-Object System.Drawing.Point(32,301)
$ButtonCalendarAccess1.Font      = 'Microsoft Sans Serif,10'

$ButtonGiveAccessAll             = New-Object system.Windows.Forms.Button
$ButtonGiveAccessAll.text        = "Access All"
$ButtonGiveAccessAll.width       = 114
$ButtonGiveAccessAll.height      = 30
$ButtonGiveAccessAll.location    = New-Object System.Drawing.Point(32,341)
$ButtonGiveAccessAll.Font        = 'Microsoft Sans Serif,10'

$ButtonRemoveAccess              = New-Object system.Windows.Forms.Button
$ButtonRemoveAccess.text         = "Remove Access"
$ButtonRemoveAccess.width        = 114
$ButtonRemoveAccess.height       = 30
$ButtonRemoveAccess.location     = New-Object System.Drawing.Point(32,383)
$ButtonRemoveAccess.Font         = 'Microsoft Sans Serif,10'

$ButtonCheckAccess               = New-Object system.Windows.Forms.Button
$ButtonCheckAccess.text          = "Check Access"
$ButtonCheckAccess.width         = 116
$ButtonCheckAccess.height        = 30
$ButtonCheckAccess.location      = New-Object System.Drawing.Point(32,427)
$ButtonCheckAccess.Font          = 'Microsoft Sans Serif,10'

$ButtonFullAccess1               = New-Object system.Windows.Forms.Button
$ButtonFullAccess1.text          = "Full Access"
$ButtonFullAccess1.width         = 124
$ButtonFullAccess1.height        = 30
$ButtonFullAccess1.location      = New-Object System.Drawing.Point(17,83)
$ButtonFullAccess1.Font          = 'Microsoft Sans Serif,10'
$ButtonFullAccess1.ForeColor     = "#7ed321"

$ButtonRemoveFull                = New-Object system.Windows.Forms.Button
$ButtonRemoveFull.text           = "Remove Access"
$ButtonRemoveFull.width          = 127
$ButtonRemoveFull.height         = 30
$ButtonRemoveFull.location       = New-Object System.Drawing.Point(17,166)
$ButtonRemoveFull.Font           = 'Microsoft Sans Serif,10'
$ButtonRemoveFull.ForeColor      = "#d0021b"

$ButtonSendAs                    = New-Object system.Windows.Forms.Button
$ButtonSendAs.text               = "Send As"
$ButtonSendAs.width              = 125
$ButtonSendAs.height             = 30
$ButtonSendAs.location           = New-Object System.Drawing.Point(17,124)
$ButtonSendAs.Font               = 'Microsoft Sans Serif,10'
$ButtonSendAs.ForeColor          = "#7ed321"

$ButtonForward                   = New-Object system.Windows.Forms.Button
$ButtonForward.text              = "Forwarding"
$ButtonForward.width             = 128
$ButtonForward.height            = 30
$ButtonForward.location          = New-Object System.Drawing.Point(17,208)
$ButtonForward.Font              = 'Microsoft Sans Serif,10'
$ButtonForward.ForeColor         = "#f5a623"

$ButtonFullAll                   = New-Object system.Windows.Forms.Button
$ButtonFullAll.text              = "Full Access ALL"
$ButtonFullAll.width             = 147
$ButtonFullAll.height            = 30
$ButtonFullAll.location          = New-Object System.Drawing.Point(160,165)
$ButtonFullAll.Font              = 'Microsoft Sans Serif,10'
$ButtonFullAll.ForeColor         = "#f5a623"

$ButtonRemoveAll                 = New-Object system.Windows.Forms.Button
$ButtonRemoveAll.text            = "Remove Access ALL"
$ButtonRemoveAll.width           = 148
$ButtonRemoveAll.height          = 30
$ButtonRemoveAll.location        = New-Object System.Drawing.Point(160,208)
$ButtonRemoveAll.Font            = 'Microsoft Sans Serif,10'
$ButtonRemoveAll.ForeColor       = "#d0021b"

$ButtonCheckLogs                 = New-Object system.Windows.Forms.Button
$ButtonCheckLogs.text            = "Check Mail Logs (48 Hrs)"
$ButtonCheckLogs.width           = 179
$ButtonCheckLogs.height          = 30
$ButtonCheckLogs.location        = New-Object System.Drawing.Point(197,517)
$ButtonCheckLogs.Font            = 'Microsoft Sans Serif,10'
$ButtonCheckLogs.ForeColor       = "#4a90e2"

$ButtonRenameUPN                 = New-Object system.Windows.Forms.Button
$ButtonRenameUPN.text            = "Change Primary UPN"
$ButtonRenameUPN.width           = 149
$ButtonRenameUPN.height          = 30
$ButtonRenameUPN.location        = New-Object System.Drawing.Point(196,301)
$ButtonRenameUPN.Font            = 'Microsoft Sans Serif,10'

$ButtonRenameUser                = New-Object system.Windows.Forms.Button
$ButtonRenameUser.text           = "Update Names"
$ButtonRenameUser.width          = 149
$ButtonRenameUser.height         = 30
$ButtonRenameUser.location       = New-Object System.Drawing.Point(196,341)
$ButtonRenameUser.Font           = 'Microsoft Sans Serif,10'

$ButtonDisableUser               = New-Object system.Windows.Forms.Button
$ButtonDisableUser.text          = "Disable Account"
$ButtonDisableUser.width         = 150
$ButtonDisableUser.height        = 30
$ButtonDisableUser.location      = New-Object System.Drawing.Point(196,383)
$ButtonDisableUser.Font          = 'Microsoft Sans Serif,10'
$ButtonDisableUser.ForeColor     = "#d0021b"

$ButtonEnableUser                = New-Object system.Windows.Forms.Button
$ButtonEnableUser.text           = "Enable Account"
$ButtonEnableUser.width          = 150
$ButtonEnableUser.height         = 30
$ButtonEnableUser.location       = New-Object System.Drawing.Point(197,427)
$ButtonEnableUser.Font           = 'Microsoft Sans Serif,10'
$ButtonEnableUser.ForeColor      = "#7ed321"

$ButtonBlockEmail                = New-Object system.Windows.Forms.Button
$ButtonBlockEmail.text           = "Block Email"
$ButtonBlockEmail.width          = 148
$ButtonBlockEmail.height         = 30
$ButtonBlockEmail.location       = New-Object System.Drawing.Point(21,517)
$ButtonBlockEmail.Font           = 'Microsoft Sans Serif,10'

$ButtonBlockDomain               = New-Object system.Windows.Forms.Button
$ButtonBlockDomain.text          = "Block Domain"
$ButtonBlockDomain.width         = 147
$ButtonBlockDomain.height        = 30
$ButtonBlockDomain.location      = New-Object System.Drawing.Point(21,557)
$ButtonBlockDomain.Font          = 'Microsoft Sans Serif,10'

$ButtonPrepareTenancy            = New-Object system.Windows.Forms.Button
$ButtonPrepareTenancy.text       = "Prepare Tenancy"
$ButtonPrepareTenancy.width      = 147
$ButtonPrepareTenancy.height     = 30
$ButtonPrepareTenancy.location   = New-Object System.Drawing.Point(21,599)
$ButtonPrepareTenancy.Font       = 'Microsoft Sans Serif,10'
$ButtonPrepareTenancy.ForeColor  = "#d0021b"

$ButtonEnableAuditLog            = New-Object system.Windows.Forms.Button
$ButtonEnableAuditLog.text       = "Enable Audit Logging"
$ButtonEnableAuditLog.width      = 149
$ButtonEnableAuditLog.height     = 30
$ButtonEnableAuditLog.location   = New-Object System.Drawing.Point(21,643)
$ButtonEnableAuditLog.Font       = 'Microsoft Sans Serif,10'
$ButtonEnableAuditLog.ForeColor  = "#7ed321"

$LabelClickInstallPowershell     = New-Object system.Windows.Forms.Label
$LabelClickInstallPowershell.text  = "Check for Updates"
$LabelClickInstallPowershell.AutoSize  = $true
$LabelClickInstallPowershell.width  = 25
$LabelClickInstallPowershell.height  = 10
$LabelClickInstallPowershell.location  = New-Object System.Drawing.Point(24,28)
$LabelClickInstallPowershell.Font  = 'Microsoft Sans Serif,10'

$LabelWiki                       = New-Object system.Windows.Forms.Label
$LabelWiki.text                  = "HELP WIKI"
$LabelWiki.AutoSize              = $true
$LabelWiki.width                 = 25
$LabelWiki.height                = 10
$LabelWiki.location              = New-Object System.Drawing.Point(213,28)
$LabelWiki.Font                  = 'Microsoft Sans Serif,10,style=Bold'
$LabelWiki.ForeColor             = "#f5a623"

$LabelSignOutClose               = New-Object system.Windows.Forms.Label
$LabelSignOutClose.text          = "Sign Out & Exit"
$LabelSignOutClose.AutoSize      = $true
$LabelSignOutClose.width         = 25
$LabelSignOutClose.height        = 10
$LabelSignOutClose.location      = New-Object System.Drawing.Point(370,28)
$LabelSignOutClose.Font          = 'Microsoft Sans Serif,10,style=Bold'
$LabelSignOutClose.ForeColor     = "#d0021b"

$ButtonAllDistMembers            = New-Object system.Windows.Forms.Button
$ButtonAllDistMembers.text       = "All Dist Members"
$ButtonAllDistMembers.width      = 179
$ButtonAllDistMembers.height     = 30
$ButtonAllDistMembers.location   = New-Object System.Drawing.Point(197,557)
$ButtonAllDistMembers.Font       = 'Microsoft Sans Serif,10'
$ButtonAllDistMembers.ForeColor  = "#4a90e2"

$ButtonAllPerms                  = New-Object system.Windows.Forms.Button
$ButtonAllPerms.text             = "All User Permissions"
$ButtonAllPerms.width            = 180
$ButtonAllPerms.height           = 30
$ButtonAllPerms.location         = New-Object System.Drawing.Point(197,599)
$ButtonAllPerms.Font             = 'Microsoft Sans Serif,10'
$ButtonAllPerms.ForeColor        = "#4a90e2"

$ButtonEnableOOF                 = New-Object system.Windows.Forms.Button
$ButtonEnableOOF.text            = "Enable OOF"
$ButtonEnableOOF.width           = 145
$ButtonEnableOOF.height          = 30
$ButtonEnableOOF.location        = New-Object System.Drawing.Point(161,83)
$ButtonEnableOOF.Font            = 'Microsoft Sans Serif,10'
$ButtonEnableOOF.ForeColor       = "#f5a623"

$ButtonDisableOOF                = New-Object system.Windows.Forms.Button
$ButtonDisableOOF.text           = "Disable OOF"
$ButtonDisableOOF.width          = 146
$ButtonDisableOOF.height         = 30
$ButtonDisableOOF.location       = New-Object System.Drawing.Point(161,124)
$ButtonDisableOOF.Font           = 'Microsoft Sans Serif,10'
$ButtonDisableOOF.ForeColor      = "#d0021b"

$GroupBoxMenu                    = New-Object system.Windows.Forms.Groupbox
$GroupBoxMenu.height             = 35
$GroupBoxMenu.width              = 464
$GroupBoxMenu.text               = "Options"
$GroupBoxMenu.location           = New-Object System.Drawing.Point(9,15)

$ButtonAllForwards               = New-Object system.Windows.Forms.Button
$ButtonAllForwards.text          = "All Mail Forwarding"
$ButtonAllForwards.width         = 182
$ButtonAllForwards.height        = 30
$ButtonAllForwards.location      = New-Object System.Drawing.Point(196,643)
$ButtonAllForwards.Font          = 'Microsoft Sans Serif,10'
$ButtonAllForwards.ForeColor     = "#4a90e2"

$GroupboxMailBox                 = New-Object system.Windows.Forms.Groupbox
$GroupboxMailBox.height          = 191
$GroupboxMailBox.width           = 308
$GroupboxMailBox.text            = "Mailbox Tasks"
$GroupboxMailBox.location        = New-Object System.Drawing.Point(9,65)

$GroupboxCalendar                = New-Object system.Windows.Forms.Groupbox
$GroupboxCalendar.height         = 196
$GroupboxCalendar.width          = 166
$GroupboxCalendar.text           = "Calendar Tasks"
$GroupboxCalendar.location       = New-Object System.Drawing.Point(9,284)

$GroupboxUsers                   = New-Object system.Windows.Forms.Groupbox
$GroupboxUsers.height            = 197
$GroupboxUsers.width             = 178
$GroupboxUsers.text              = "User Tasks"
$GroupboxUsers.location          = New-Object System.Drawing.Point(183,284)

$GroupboxSecurity                = New-Object system.Windows.Forms.Groupbox
$GroupboxSecurity.height         = 189
$GroupboxSecurity.width          = 172
$GroupboxSecurity.text           = "Security"
$GroupboxSecurity.location       = New-Object System.Drawing.Point(9,499)

$CheckMailboxPerms               = New-Object system.Windows.Forms.Button
$CheckMailboxPerms.text          = "Check Mailbox Perms"
$CheckMailboxPerms.width         = 196
$CheckMailboxPerms.height        = 30
$CheckMailboxPerms.location      = New-Object System.Drawing.Point(328,208)
$CheckMailboxPerms.Font          = 'Microsoft Sans Serif,10'

$LabelCustomCommand              = New-Object system.Windows.Forms.Label
$LabelCustomCommand.text         = "Run Custom Command"
$LabelCustomCommand.AutoSize     = $true
$LabelCustomCommand.width        = 25
$LabelCustomCommand.height       = 10
$LabelCustomCommand.location     = New-Object System.Drawing.Point(304,696)
$LabelCustomCommand.Font         = 'Microsoft Sans Serif,10,style=Bold'
$LabelCustomCommand.ForeColor    = "#9013fe"


$LabelLogs                       = New-Object system.Windows.Forms.Label
$LabelLogs.text                  = "Open Log Folder"
$LabelLogs.AutoSize              = $true
$LabelLogs.width                 = 25
$LabelLogs.height                = 10
$LabelLogs.location              = New-Object System.Drawing.Point(40,696)
$LabelLogs.Font                  = 'Microsoft Sans Serif,10,style=Bold'
$LabelLogs.ForeColor             = "#9013fe"


$LabelReport                     = New-Object system.Windows.Forms.Label
$LabelReport.text                = "Generate 365 Report"
$LabelReport.AutoSize            = $true
$LabelReport.width               = 25
$LabelReport.height              = 10
$LabelReport.location            = New-Object System.Drawing.Point(212,696)
$LabelReport.Font                = 'Microsoft Sans Serif,10,style=Bold'
$LabelReport.ForeColor           = "#f5a623"

$BWApp.controls.AddRange(@($ButtonCalendarAccess1,$ButtonGiveAccessAll,$ButtonRemoveAccess,$ButtonCheckAccess,$ButtonFullAccess1,$ButtonRemoveFull,$ButtonSendAs,$ButtonForward,$ButtonFullAll,$ButtonRemoveAll,$ButtonCheckLogs,$ButtonRenameUPN,$ButtonRenameUser,$ButtonDisableUser,$ButtonEnableUser,$ButtonBlockEmail,$ButtonBlockDomain,$ButtonPrepareTenancy,$ButtonEnableAuditLog,$LabelClickInstallPowershell,$LabelWiki,$LabelSignOutClose,$ButtonAllDistMembers,$ButtonAllPerms,$ButtonEnableOOF,$ButtonDisableOOF,$GroupBoxMenu,$ButtonAllForwards,$GroupboxMailBox,$GroupboxCalendar,$GroupboxUsers,$GroupboxSecurity,$CheckMailboxPerms,$LabelCustomCommand,$LabelLogs,$LabelReport))

$ButtonFullAccess1.Add_Click({.\MailboxGrantFull.ps1})
$ButtonRemoveFull.Add_Click({.\MailboxRemoveFull.ps1})
$ButtonSendAs.Add_Click({.\MailboxSend.ps1})
$ButtonForward.Add_Click({.\MailboxForward.ps1})
$ButtonCalendarAccess1.Add_Click({.\calendarAccess1.ps1})
$ButtonRemoveAccess.Add_Click({.\CalendarAccessRemove.ps1})
$ButtonGiveAccessAll.Add_Click({.\CalendarAccessAll.ps1})
$ButtonCheckAccess.Add_Click({.\CalendarCheck.ps1})
$ButtonFullAll.Add_Click({.\MailboxGrantALL.ps1})
$ButtonRemoveAll.Add_Click({.\MailboxRemoveALL.ps1})
$ButtonCheckLogs.Add_Click({ .\365MessageTrace.ps1})
$ButtonRenameUPN.Add_Click({.\RenameSyncedUsers.ps1})
$ButtonRenameUser.Add_Click({.\RenameFirstLastName.ps1})
$ButtonDisableUser.Add_Click({.\365BlockUser.ps1})
$ButtonEnableUser.Add_Click({.\365EnableUser.ps1})
$ButtonBlockEmail.Add_Click({.\365BlockSender.ps1})
$ButtonBlockDomain.Add_Click({.\365BlockDomain.ps1})
$ButtonPrepareTenancy.Add_Click({.\prepareTenant.ps1})
$ButtonEnableAuditLog.Add_Click({.\365EnableAuditLog.ps1})
$LabelClickInstallPowershell.Add_Click({.\updateApp.ps1})
$LabelSignOutClose.Add_Click({.\SignOutClose.ps1;$BWApp.Close()})
$LabelWiki.Add_Click({Start "https://github.com/valllem/BWApp/wiki"})
$ButtonEnableOOF.Add_Click({.\365EnableOOF.ps1})
$ButtonDisableOOF.Add_Click({.\365DisableOOF.ps1})
$ButtonAllDistMembers.Add_Click({.\Get-All-DistMembers.ps1})
$ButtonAllPerms.Add_Click({.\Get-All-Perms.ps1})
$ButtonAllForwards.Add_Click({.\AllForwards.ps1})
$CheckMailboxPerms.Add_Click({.\MailboxCheck.ps1})
$LabelCustomCommand.Add_Click({.\CustomCommand.ps1})
$LabelLogs.Add_Click({ii C:\BWApp\Logs\})
$LabelReport.Add_Click({.\Report.ps1})

$result = $BWApp.ShowDialog()
