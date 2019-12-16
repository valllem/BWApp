## PROGRESS BAR ##
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()


    $ObjForm = New-Object System.Windows.Forms.Form
	$ObjForm.Text = "Closing App"
	$ObjForm.Height = 100
	$ObjForm.Width = 500
	$ObjForm.BackColor = "White"

	$ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

	## -- Create The Label
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Text = "Closing App"
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
	$ObjLabel.Text = "Closing App"
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300
#####


    $ObjForm.Refresh()
    $PB.Value = 30
	$ObjLabel.Text = "Disconnecting from 365"
	Start-Sleep -Milliseconds 300
    
    
    Exit-PSSession


	$ObjLabel.Text = "Disconnected"
	Start-Sleep -Milliseconds 300


    $ObjForm.Refresh()
    $PB.Value = 60
	$ObjLabel.Text = "Tidying Log Files"
	Start-Sleep -Milliseconds 500

## TIDY LOGS ##

    $DateTime = Get-Date -Format "dd-MM-yyyy_HHmm"
    Rename-Item -Path "c:\BWApp\Logs\Log.txt" -NewName "c:\BWApp\Logs\Log_$DateTime.txt"

    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Closing BWApp"
	Start-Sleep -Milliseconds 500
    
    $ObjForm.Refresh()
    $PB.Value = 100
    Start-Sleep -Milliseconds 500

    $ObjForm.Close()

Exit
