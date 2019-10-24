$logfile = "C:\BWApp\Logs\Log.txt"
$jobstatus = "C:\Temp\BWApp\JobStatus.txt"

$RunningUser = whoami
$DateTime = Get-Date


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$mailboxes = Get-Mailbox |Select-Object PrimarySmtpAddress

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,420'
$Form.text                       = "Check Calendar Access"
$Form.TopMost                    = $false
$Form.StartPosition = 'CenterScreen'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 300
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
$Form.AcceptButton = $Button1
$Form.Controls.Add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(208,373)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $Button2
$Form.Controls.Add($Button2)

$Identity                        = New-Object system.Windows.Forms.Label
$Identity.text                   = "Calendar to Check"
$Identity.AutoSize               = $true
$Identity.width                  = 25
$Identity.height                 = 10
$Identity.location               = New-Object System.Drawing.Point(38,49)
$Identity.Font                   = 'Microsoft Sans Serif,10'
$Form.Controls.Add($Identity)

$Form.controls.AddRange(@($ListBox1,$Button1,$Button2,$Identity,$ListBox3,$permission))

$Form.Controls.Add($ListBox1)



$Form.Topmost = $true


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $ListBox1.SelectedItem
    $y = $mailbox.PrimarySmtpAddress
 
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
	$ObjLabel.Text = "Getting Permissions"
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 75
	$ObjLabel.Text = "Exporting to Logs..."
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300

Get-MailboxFolderPermission -identity $x':\calendar' | Export-Csv "C:\BWApp\Logs\CalendarPerms.csv"

    Add-Content "$logfile" "========INFO========="
    Add-Content "$logfile" "$DateTime"
    Add-Content "$Logfile" "Checked $x's Calendar Permissions."

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
    $ObjLabel.Text = "Displaying Results"
    Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 100
    Start-Sleep -Milliseconds 300


Get-MailboxFolderPermission -identity $x':\calendar' | out-gridview

$ObjForm.Close()

}

