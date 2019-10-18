
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$calendarAccess1                 = New-Object system.Windows.Forms.Form
$calendarAccess1.ClientSize      = '600,420'
$calendarAccess1.text            = "Give All Calendar Access"
$calendarAccess1.TopMost         = $false
$calendarAccess1.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 169
$ListBox1.height                 = 258
$ListBox1.location               = New-Object System.Drawing.Point(38,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(280,372)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$calendarAccess1.AcceptButton = $Button1
$calendarAccess1.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(208,373)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$calendarAccess1.CancelButton = $Button2
$calendarAccess1.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Account to Grant Access to ALL Calendars"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(66,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($Identity)

$ListBox3                        = New-Object system.Windows.Forms.ListBox
$ListBox3.text                   = "listBox"
$ListBox3.width                  = 175
$ListBox3.height                 = 260
@("Owner","PublishingAuthor","LimitedDetails") | ForEach-Object {[void] $ListBox3.Items.Add($_)}
$ListBox3.location               = New-Object System.Drawing.Point(360,77)

$permission                      = New-Object system.Windows.Forms.Label
$permission.text                 = "Which Permission"
$permission.AutoSize             = $true
$permission.width                = 25
$permission.height               = 10
$permission.location             = New-Object System.Drawing.Point(360,49)
$permission.Font                 = 'Microsoft Sans Serif,10'

$calendarAccess1.controls.AddRange(@($ListBox1,$Button1,$Button2,$Identity,$ListBox3,$permission))

$calendarAccess1.Controls.Add($ListBox1)
$calendarAccess1.Controls.Add($ListBox3)


$calendarAccess1.Topmost = $true


$result = $calendarAccess1.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $mailbox.PrimarySmtpAddress
    $z = $ListBox3.SelectedItem
    
    ## -- Create The Progress-Bar
	$ObjForm = New-Object System.Windows.Forms.Form
	$ObjForm.Text = "Running Task"
	$ObjForm.Height = 100
	$ObjForm.Width = 500
	$ObjForm.BackColor = "White"

	$ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

	## -- Create The Label
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Text = "Starting Task. Please wait ... "
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
	$ObjLabel.Text = "Preparing Script. Please wait ... "
	$ObjForm.Refresh()

	Start-Sleep -Milliseconds 300
#####
    $ObjForm.Refresh()
    $PB.Value = 10
	$ObjLabel.Text = "Adjusting Permissions"
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 10
	$ObjLabel.Text = "Giving $x Full Access to all Calendars"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300





if ($z -eq "LimitedDetails") {
    $mailboxes = Get-mailbox
    $userRequiringAccess = Get-mailbox $userRequiringAccess
    $Counter = 0
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -erroraction SilentlyContinue

    $Counter++
	[Int]$Percentage = ($Counter/$accessRights.Count)*100
	$PB.Value = $Percentage
	$ObjLabel.Text = "Granting $x Full Access to all Calendars"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 150

    Remove-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User "$x" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Green     
    if ($accessRights.accessRights -notmatch $z -and $mailbox.primarysmtpaddress -notcontains $x.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue

            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue 
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"   
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress
        if ($accessRights.accessRights -match $z) {
            
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Successfully added $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
        }
        else {
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Could not add $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
            
        }
    }else{
        Write-Host "Permission level already exists for $($x.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
    }
}
} elseif ($z -eq "PublishingAuthor") {
    $mailboxes = Get-mailbox
    $userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -erroraction SilentlyContinue

    $Counter++
	[Int]$Percentage = ($Counter/$accessRights.Count)*100
	$PB.Value = $Percentage
	$ObjLabel.Text = "Granting $x Full Access to all Calendars"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 150

    Remove-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User "$x" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Green
    if ($accessRights.accessRights -notmatch $z -and $mailbox.primarysmtpaddress -notcontains $x.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue

            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue 
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"   
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress
        if ($accessRights.accessRights -match $z) {
            
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Successfully added $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
        }
        else {
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Could not add $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
            
        }
    }else{
        Write-Host "Permission level already exists for $($x.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
    }
}
} elseif ($z -eq "Owner") {
    $mailboxes = Get-mailbox
    $userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -erroraction SilentlyContinue

    $Counter++
	[Int]$Percentage = ($Counter/$accessRights.Count)*100
	$PB.Value = $Percentage
	$ObjLabel.Text = "Granting $x Full Access to all Calendars"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 150

    Remove-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User "$x" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Green
    if ($accessRights.accessRights -notmatch $z -and $mailbox.primarysmtpaddress -notcontains $x.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue

            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress -AccessRights $z -ErrorAction SilentlyContinue 
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Added $x Full Access to $($mailbox.primarysmtpaddress) mailbox"   
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $x.PrimarySmtpAddress
        if ($accessRights.accessRights -match $z) {
            
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Successfully added $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
        }
        else {
            Add-Content "$logfile" "====================="
            Add-Content "$logfile" "$DateTime"
            Add-Content "$Logfile" "$RunningUser"
            Add-Content "$logfile" "Could not add $z permissions on $($mailbox.displayname)'s calendar for $($x.displayname)"
            
        }
    }else{
        Write-Host "Permission level already exists for $($x.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
    }
}

} else {

}
    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Completed Tasks"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 150

    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Completed Tasks"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 150
    
    
Start-Sleep -Seconds 3

}
$ObjForm.Close()
Start-Sleep -Seconds 1
Exit