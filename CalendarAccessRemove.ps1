$logfile = "C:\BWApp\Logs\Log.txt"
$jobstatus = "C:\Temp\BWApp\JobStatus.txt"

$RunningUser = whoami
$DateTime = Get-Date

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$calendarAccess1                 = New-Object system.Windows.Forms.Form
$calendarAccess1.ClientSize      = '600,420'
$calendarAccess1.text            = "Remove Calendar Access"
$calendarAccess1.TopMost         = $false
$calendarAccess1.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 169
$ListBox1.height                 = 258
$ListBox1.location               = New-Object System.Drawing.Point(38,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox1.Items.Add($mailbox.PrimarySmtpAddress)}

$ListBox2                        = New-Object system.Windows.Forms.ListBox
$ListBox2.text                   = "listBox"
$ListBox2.width                  = 166
$ListBox2.height                 = 258
$ListBox2.location               = New-Object System.Drawing.Point(327,77)
foreach ($mailbox in $mailboxes) {[void] $ListBox2.Items.Add($mailbox.PrimarySmtpAddress)}

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
$Identity.text                   = "Account to Remove Access"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(40,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($Identity)

$User                            = New-Object system.Windows.Forms.Label
$User.text                       = "Person being removed"
$User.AutoSize                   = $true
$User.width                      = 25
$User.height                     = 10
$User.location                   = New-Object System.Drawing.Point(340,49)
$User.Font                       = 'Microsoft Sans Serif,10'
$calendarAccess1.Controls.Add($User)


$calendarAccess1.controls.AddRange(@($ListBox1,$ListBox2,$Button1,$Button2,$Identity,$User))

$calendarAccess1.Controls.Add($ListBox1)
$calendarAccess1.Controls.Add($ListBox2)
$calendarAccess1.Controls.Add($ListBox3)


$calendarAccess1.Topmost = $true


$result = $calendarAccess1.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $ListBox1.SelectedItem
    $y = $ListBox2.SelectedItem
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
	$PB.Style = Continuous
    

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
    $PB.Value = 50
	$ObjLabel.Text = "Adjusting Permissions"
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 75
	$ObjLabel.Text = "Removing $y's Access to $x's Calendar"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300
    
$permission = Get-MailboxFolderPermission -identity $x':\calendar' -User "$y"
if($permission -eq $null){
    Add-Content "$logfile" "========INFO========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Tried to remove $y from $x's Calendar. $y did not have access to $x and could not be removed"
}else{
    Remove-MailboxFolderPermission $x':\calendar' -User "$y" -Confirm:$false
    Add-Content "$logfile" "========SUCCESS========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Removed $y from $x's Calendar."
}
 
    $ObjForm.Refresh()
    $PB.Value = 90
    Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 95
    Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 99
    Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Task Completed Successfully"
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 100
    Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 100
    Start-Sleep -Milliseconds 300





    $ObjForm.Close()
    exit



Start-Sleep -Seconds 2
exit
}