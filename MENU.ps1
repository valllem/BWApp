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
$BWApp.ClientSize                = '487,722'
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

$LabelLogin                      = New-Object system.Windows.Forms.Label
$LabelLogin.text                 = "Switch Account"
$LabelLogin.AutoSize             = $true
$LabelLogin.width                = 25
$LabelLogin.height               = 10
$LabelLogin.location             = New-Object System.Drawing.Point(232,28)
$LabelLogin.Font                 = 'Microsoft Sans Serif,10,style=Bold'
$LabelLogin.ForeColor            = "#f5a623"

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
$CheckMailboxPerms.width         = 138
$CheckMailboxPerms.height        = 30
$CheckMailboxPerms.location      = New-Object System.Drawing.Point(328,208)
$CheckMailboxPerms.Font          = 'Microsoft Sans Serif,10'

$BWApp.controls.AddRange(@($ButtonCalendarAccess1,$ButtonGiveAccessAll,$ButtonRemoveAccess,$ButtonCheckAccess,$ButtonFullAccess1,$ButtonRemoveFull,$ButtonSendAs,$ButtonForward,$ButtonFullAll,$ButtonRemoveAll,$ButtonCheckLogs,$ButtonRenameUPN,$ButtonRenameUser,$ButtonDisableUser,$ButtonEnableUser,$ButtonBlockEmail,$ButtonBlockDomain,$ButtonPrepareTenancy,$ButtonEnableAuditLog,$LabelClickInstallPowershell,$LabelLogin,$LabelSignOutClose,$ButtonAllDistMembers,$ButtonAllPerms,$ButtonEnableOOF,$ButtonDisableOOF,$GroupBoxMenu,$ButtonAllForwards,$GroupboxMailBox,$GroupboxCalendar,$GroupboxUsers,$GroupboxSecurity,$CheckMailboxPerms))

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
$LabelLogin.Add_Click({.\SwitchUser.ps1})
$ButtonEnableOOF.Add_Click({.\365EnableOOF.ps1})
$ButtonDisableOOF.Add_Click({.\365DisableOOF.ps1})
$ButtonAllDistMembers.Add_Click({.\Get-All-DistMembers.ps1})
$ButtonAllPerms.Add_Click({.\Get-All-Perms.ps1})
$ButtonAllForwards.Add_Click({.\AllForwards.ps1})
$CheckMailboxPerms.Add_Click({.\MailboxCheck.ps1})

$result = $BWApp.ShowDialog()
