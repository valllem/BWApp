<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    MENU
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$BWApp                           = New-Object system.Windows.Forms.Form
$BWApp.ClientSize                = '626,462'
$BWApp.text                      = "BWApp - Main Menu"
$BWApp.TopMost                   = $false
$BWApp.StartPosition = 'CenterScreen'


$calendars                       = New-Object system.Windows.Forms.Button
$calendars.text                  = "CALENDARS"
$calendars.width                 = 150
$calendars.height                = 56
$calendars.location              = New-Object System.Drawing.Point(9,10)
$calendars.Font                  = 'Microsoft Sans Serif,10'

$mailboxes                       = New-Object system.Windows.Forms.Button
$mailboxes.text                  = "MAILBOXES"
$mailboxes.width                 = 154
$mailboxes.height                = 60
$mailboxes.location              = New-Object System.Drawing.Point(8,78)
$mailboxes.Font                  = 'Microsoft Sans Serif,10'

$windows                         = New-Object system.Windows.Forms.Button
$windows.text                    = "WINDOWS"
$windows.width                   = 156
$windows.height                  = 54
$windows.location                = New-Object System.Drawing.Point(7,147)
$windows.Font                    = 'Microsoft Sans Serif,10'

$configuration                   = New-Object system.Windows.Forms.Button
$configuration.text              = "Configuration"
$configuration.width             = 157
$configuration.height            = 49
$configuration.location          = New-Object System.Drawing.Point(8,216)
$configuration.Font              = 'Microsoft Sans Serif,10'

$security                        = New-Object system.Windows.Forms.Button
$security.text                   = "SECURITY"
$security.width                  = 156
$security.height                 = 50
$security.location               = New-Object System.Drawing.Point(8,279)
$security.Font                   = 'Microsoft Sans Serif,10'

$users                           = New-Object system.Windows.Forms.Button
$users.text                      = "USERS"
$users.width                     = 158
$users.height                    = 53
$users.location                  = New-Object System.Drawing.Point(6,344)
$users.Font                      = 'Microsoft Sans Serif,10'

$signout                         = New-Object system.Windows.Forms.Button
$signout.text                    = "Sign Out"
$signout.width                   = 60
$signout.height                  = 30
$signout.location                = New-Object System.Drawing.Point(525,416)
$signout.Font                    = 'Microsoft Sans Serif,10'

$login                           = New-Object system.Windows.Forms.Button
$login.text                      = "Switch User"
$login.width                     = 84
$login.height                    = 30
$login.location                  = New-Object System.Drawing.Point(432,416)
$login.Font                      = 'Microsoft Sans Serif,10'

$BWApp.controls.AddRange(@($calendars,$mailboxes,$windows,$configuration,$security,$users,$signout,$login))

$calendars.Add_Click({ .\calendarAccess1-2.ps1 })
$mailboxes.Add_Click({ .\menuMailbox.ps1 })
$windows.Add_Click({ .\menuWindows.ps1 })
$configuration.Add_Click({ .\menuConfiguration.ps1 })
$security.Add_Click({ .\menu365Security.ps1 })
$users.Add_Click({ .\menuUsers.ps1 })

$BWApp.ShowDialog()