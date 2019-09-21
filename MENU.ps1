if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}



$Path = Get-Content -Path "C:\Temp\BWApp\Path.txt"
write-host $Path
Set-Location -Path $Path

Write-Host "`n"
Write-Host -ForegroundColor Yellow "Signing In`n"

$UserCredential = Get-Credential
try {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Start-Sleep -Seconds 1
Connect-AzureAD -Credential $UserCredential
Connect-MsolService -Credential $UserCredential
}
catch {
Clear-Host
Write-Host -ForegroundColor Yellow "Using MFA...Switching to Modern Auth."
Write-Host -ForegroundColor Yellow "Please Sign In again"
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
-Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"} `
|select -First 1)
$EXOSession = New-ExoPSSession
Import-PSSession $EXOSession
}






Clear-Host
write-host "`n"
write-host -foregroundcolor Green "Connected to 365`n"







Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$BWApp                           = New-Object system.Windows.Forms.Form
$BWApp.ClientSize                = '659,649'
$BWApp.text                      = "BWApp - Main Menu"
$BWApp.TopMost                   = $false

$LabelCalendar                  = New-Object system.Windows.Forms.Label
$LabelCalendar.text             = "Calendars"
$LabelCalendar.AutoSize         = $true
$LabelCalendar.width            = 25
$LabelCalendar.height           = 10
$LabelCalendar.location         = New-Object System.Drawing.Point(350,11)
$LabelCalendar.Font             = 'Microsoft Sans Serif,14,style=Underline'

$ButtonCalendarAccess1           = New-Object system.Windows.Forms.Button
$ButtonCalendarAccess1.text      = "Give Access"
$ButtonCalendarAccess1.width     = 96
$ButtonCalendarAccess1.height    = 30
$ButtonCalendarAccess1.location  = New-Object System.Drawing.Point(350,50)
$ButtonCalendarAccess1.Font      = 'Microsoft Sans Serif,10'

$ButtonGiveAccessAll             = New-Object system.Windows.Forms.Button
$ButtonGiveAccessAll.text        = "Access All"
$ButtonGiveAccessAll.width       = 97
$ButtonGiveAccessAll.height      = 30
$ButtonGiveAccessAll.location    = New-Object System.Drawing.Point(350,96)
$ButtonGiveAccessAll.Font        = 'Microsoft Sans Serif,10'

$ButtonRemoveAccess              = New-Object system.Windows.Forms.Button
$ButtonRemoveAccess.text         = "Remove Access"
$ButtonRemoveAccess.width        = 114
$ButtonRemoveAccess.height       = 30
$ButtonRemoveAccess.location     = New-Object System.Drawing.Point(477,50)
$ButtonRemoveAccess.Font         = 'Microsoft Sans Serif,10'

$ButtonCheckAccess               = New-Object system.Windows.Forms.Button
$ButtonCheckAccess.text          = "Check Access"
$ButtonCheckAccess.width         = 116
$ButtonCheckAccess.height        = 30
$ButtonCheckAccess.location      = New-Object System.Drawing.Point(477,96)
$ButtonCheckAccess.Font          = 'Microsoft Sans Serif,10'

$LabelMailboxes                  = New-Object system.Windows.Forms.Label
$LabelMailboxes.text             = "Mailboxes"
$LabelMailboxes.AutoSize         = $true
$LabelMailboxes.width            = 25
$LabelMailboxes.height           = 10
$LabelMailboxes.location         = New-Object System.Drawing.Point(30,11)
$LabelMailboxes.Font             = 'Microsoft Sans Serif,14,style=Underline'

$ButtonFullAccess1               = New-Object system.Windows.Forms.Button
$ButtonFullAccess1.text          = "Full Access"
$ButtonFullAccess1.width         = 117
$ButtonFullAccess1.height        = 30
$ButtonFullAccess1.location      = New-Object System.Drawing.Point(20,52)
$ButtonFullAccess1.Font          = 'Microsoft Sans Serif,10'

$ButtonRemoveFull                = New-Object system.Windows.Forms.Button
$ButtonRemoveFull.text           = "Remove Access"
$ButtonRemoveFull.width          = 136
$ButtonRemoveFull.height         = 30
$ButtonRemoveFull.location       = New-Object System.Drawing.Point(148,52)
$ButtonRemoveFull.Font           = 'Microsoft Sans Serif,10'

$ButtonSendAs                    = New-Object system.Windows.Forms.Button
$ButtonSendAs.text               = "Send As"
$ButtonSendAs.width              = 117
$ButtonSendAs.height             = 30
$ButtonSendAs.location           = New-Object System.Drawing.Point(20,101)
$ButtonSendAs.Font               = 'Microsoft Sans Serif,10'

$ButtonForward                   = New-Object system.Windows.Forms.Button
$ButtonForward.text              = "Forwarding"
$ButtonForward.width             = 135
$ButtonForward.height            = 30
$ButtonForward.location          = New-Object System.Drawing.Point(149,101)
$ButtonForward.Font              = 'Microsoft Sans Serif,10'

$ButtonFullAll                   = New-Object system.Windows.Forms.Button
$ButtonFullAll.text              = "Full Access ALL"
$ButtonFullAll.width             = 118
$ButtonFullAll.height            = 30
$ButtonFullAll.location          = New-Object System.Drawing.Point(20,151)
$ButtonFullAll.Font              = 'Microsoft Sans Serif,10'

$ButtonRemoveAll                 = New-Object system.Windows.Forms.Button
$ButtonRemoveAll.text            = "Remove Access ALL"
$ButtonRemoveAll.width           = 140
$ButtonRemoveAll.height          = 30
$ButtonRemoveAll.location        = New-Object System.Drawing.Point(147,151)
$ButtonRemoveAll.Font            = 'Microsoft Sans Serif,10'

$ButtonCheckLogs                 = New-Object system.Windows.Forms.Button
$ButtonCheckLogs.text            = "Check Logs (Last 48 hours)"
$ButtonCheckLogs.width           = 266
$ButtonCheckLogs.height          = 30
$ButtonCheckLogs.location        = New-Object System.Drawing.Point(20,192)
$ButtonCheckLogs.Font            = 'Microsoft Sans Serif,10'

$LabelUsers                      = New-Object system.Windows.Forms.Label
$LabelUsers.text                 = "Users"
$LabelUsers.AutoSize             = $true
$LabelUsers.width                = 25
$LabelUsers.height               = 10
$LabelUsers.location             = New-Object System.Drawing.Point(55,338)
$LabelUsers.Font                 = 'Microsoft Sans Serif,14,style=Underline'

$ButtonRenameUPN                 = New-Object system.Windows.Forms.Button
$ButtonRenameUPN.text            = "Change Primary UPN"
$ButtonRenameUPN.width           = 145
$ButtonRenameUPN.height          = 30
$ButtonRenameUPN.location        = New-Object System.Drawing.Point(30,375)
$ButtonRenameUPN.Font            = 'Microsoft Sans Serif,10'

$ButtonRenameUser                = New-Object system.Windows.Forms.Button
$ButtonRenameUser.text           = "Update Names"
$ButtonRenameUser.width          = 146
$ButtonRenameUser.height         = 30
$ButtonRenameUser.location       = New-Object System.Drawing.Point(30,418)
$ButtonRenameUser.Font           = 'Microsoft Sans Serif,10'

$ButtonDisableUser               = New-Object system.Windows.Forms.Button
$ButtonDisableUser.text          = "Disable Account"
$ButtonDisableUser.width         = 147
$ButtonDisableUser.height        = 30
$ButtonDisableUser.location      = New-Object System.Drawing.Point(30,465)
$ButtonDisableUser.Font          = 'Microsoft Sans Serif,10'
$ButtonDisableUser.ForeColor     = "#d0021b"

$ButtonEnableUser                = New-Object system.Windows.Forms.Button
$ButtonEnableUser.text           = "Enable Account"
$ButtonEnableUser.width          = 149
$ButtonEnableUser.height         = 30
$ButtonEnableUser.location       = New-Object System.Drawing.Point(30,515)
$ButtonEnableUser.Font           = 'Microsoft Sans Serif,10'
$ButtonEnableUser.ForeColor      = "#7ed321"

$LabelSecurity                   = New-Object system.Windows.Forms.Label
$LabelSecurity.text              = "Security"
$LabelSecurity.AutoSize          = $true
$LabelSecurity.width             = 25
$LabelSecurity.height            = 10
$LabelSecurity.location          = New-Object System.Drawing.Point(351,338)
$LabelSecurity.Font              = 'Microsoft Sans Serif,14,style=Underline'

$ButtonBlockEmail                = New-Object system.Windows.Forms.Button
$ButtonBlockEmail.text           = "Block Email"
$ButtonBlockEmail.width          = 100
$ButtonBlockEmail.height         = 30
$ButtonBlockEmail.location       = New-Object System.Drawing.Point(348,377)
$ButtonBlockEmail.Font           = 'Microsoft Sans Serif,10'

$ButtonBlockDomain               = New-Object system.Windows.Forms.Button
$ButtonBlockDomain.text          = "Block Domain"
$ButtonBlockDomain.width         = 101
$ButtonBlockDomain.height        = 30
$ButtonBlockDomain.location      = New-Object System.Drawing.Point(349,421)
$ButtonBlockDomain.Font          = 'Microsoft Sans Serif,10'

$ButtonPrepareTenancy            = New-Object system.Windows.Forms.Button
$ButtonPrepareTenancy.text       = "Prepare Tenancy"
$ButtonPrepareTenancy.width      = 127
$ButtonPrepareTenancy.height     = 30
$ButtonPrepareTenancy.location   = New-Object System.Drawing.Point(349,469)
$ButtonPrepareTenancy.Font       = 'Microsoft Sans Serif,10'
$ButtonPrepareTenancy.ForeColor  = "#d0021b"

$ButtonEnableAuditLog            = New-Object system.Windows.Forms.Button
$ButtonEnableAuditLog.text       = "Enable Audit Logging"
$ButtonEnableAuditLog.width      = 149
$ButtonEnableAuditLog.height     = 30
$ButtonEnableAuditLog.location   = New-Object System.Drawing.Point(350,516)
$ButtonEnableAuditLog.Font       = 'Microsoft Sans Serif,10'
$ButtonEnableAuditLog.ForeColor  = "#7ed321"

$LabelClickInstallPowershell     = New-Object system.Windows.Forms.Label
$LabelClickInstallPowershell.text  = "Check for Updates"
$LabelClickInstallPowershell.AutoSize  = $true
$LabelClickInstallPowershell.width  = 25
$LabelClickInstallPowershell.height  = 10
$LabelClickInstallPowershell.location  = New-Object System.Drawing.Point(26,623)
$LabelClickInstallPowershell.Font  = 'Microsoft Sans Serif,10'

$LabelLogin                      = New-Object system.Windows.Forms.Label
$LabelLogin.text                 = "Switch Account"
$LabelLogin.AutoSize             = $true
$LabelLogin.width                = 25
$LabelLogin.height               = 10
$LabelLogin.location             = New-Object System.Drawing.Point(301,600)
$LabelLogin.Font                 = 'Microsoft Sans Serif,10,style=Bold'
$LabelLogin.ForeColor            = "#f5a623"

$LabelSignOutClose               = New-Object system.Windows.Forms.Label
$LabelSignOutClose.text          = "Sign Out & Exit"
$LabelSignOutClose.AutoSize      = $true
$LabelSignOutClose.width         = 25
$LabelSignOutClose.height        = 10
$LabelSignOutClose.location      = New-Object System.Drawing.Point(445,600)
$LabelSignOutClose.Font          = 'Microsoft Sans Serif,10,style=Bold'
$LabelSignOutClose.ForeColor     = "#d0021b"

$ButtonAllDistMembers            = New-Object system.Windows.Forms.Button
$ButtonAllDistMembers.text       = "All Dist Members"
$ButtonAllDistMembers.width      = 120
$ButtonAllDistMembers.height     = 30
$ButtonAllDistMembers.location   = New-Object System.Drawing.Point(20,237)
$ButtonAllDistMembers.Font       = 'Microsoft Sans Serif,10'

$ButtonAllPerms                  = New-Object system.Windows.Forms.Button
$ButtonAllPerms.text             = "All User Permissions"
$ButtonAllPerms.width            = 140
$ButtonAllPerms.height           = 30
$ButtonAllPerms.location         = New-Object System.Drawing.Point(148,236)
$ButtonAllPerms.Font             = 'Microsoft Sans Serif,10'

$ButtonEnableOOF                 = New-Object system.Windows.Forms.Button
$ButtonEnableOOF.text            = "Enable OOF"
$ButtonEnableOOF.width           = 121
$ButtonEnableOOF.height          = 30
$ButtonEnableOOF.location        = New-Object System.Drawing.Point(20,281)
$ButtonEnableOOF.Font            = 'Microsoft Sans Serif,10'

$ButtonDisableOOF                = New-Object system.Windows.Forms.Button
$ButtonDisableOOF.text           = "Disable OOF"
$ButtonDisableOOF.width          = 130
$ButtonDisableOOF.height         = 30
$ButtonDisableOOF.location       = New-Object System.Drawing.Point(154,281)
$ButtonDisableOOF.Font           = 'Microsoft Sans Serif,10'

$BWApp.controls.AddRange(@($LabelCalendar,$ButtonCalendarAccess1,$ButtonGiveAccessAll,$ButtonRemoveAccess,$ButtonCheckAccess,$LabelMailboxes,$ButtonFullAccess1,$ButtonRemoveFull,$ButtonSendAs,$ButtonForward,$ButtonFullAll,$ButtonRemoveAll,$ButtonCheckLogs,$LabelUsers,$ButtonRenameUPN,$ButtonRenameUser,$ButtonDisableUser,$ButtonEnableUser,$LabelSecurity,$ButtonBlockEmail,$ButtonBlockDomain,$ButtonPrepareTenancy,$ButtonEnableAuditLog,$LabelClickInstallPowershell,$LabelLogin,$LabelSignOutClose,$ButtonEnableOOF,$ButtonDisableOOF,$ButtonAllDistMembers,$ButtonAllPerms))

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
$LabelClickInstallPowershell.Add_Click({.\updateApp.ps1;$BWApp.Close()})
$LabelSignOutClose.Add_Click({.\SignOutClose.ps1;$BWApp.Close()})
$LabelLogin.Add_Click({.\SwitchUser.ps1})
$ButtonEnableOOF.Add_Click({.\365EnableOOF.ps1})
$ButtonDisableOOF.Add_Click({.\365DisableOOF.ps1})
$ButtonAllDistMembers.Add_Click({.\Get-All-DistMembers.ps1})
$ButtonAllPerms.Add_Click({.\Get-All-Perms.ps1})

$result = $BWApp.ShowDialog()
