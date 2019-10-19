$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp.zip”)

## ENABLE THE GUI ##
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

## PROGRESS BAR INSTALLING APP ##
    $ObjForm = New-Object System.Windows.Forms.Form
	$ObjForm.Text = "Updating"
	$ObjForm.Height = 100
	$ObjForm.Width = 500
	$ObjForm.BackColor = "White"

	$ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    ## -- Create The Label
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Text = "Updating App. Please wait ... "
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
	$ObjLabel.Text = "Updating App. Please wait ..."
	$ObjForm.Refresh()
	Start-Sleep -Milliseconds 300
    
    $ObjForm.Refresh()
    $PB.Value = 15
	$ObjLabel.Text = "Updating App. Getting Ready ..."
	Start-Sleep -Milliseconds 300


    $ObjForm.Refresh()
    $PB.Value = 30
	$ObjLabel.Text = "Updating App. Downloading Components ..."
	Start-Sleep -Milliseconds 300


(New-Object System.Net.WebClient).DownloadFile($url, $output)

    $ObjForm.Refresh()
    $PB.Value = 40
	$ObjLabel.Text = "Updating App. File Downloaded ..."
	Start-Sleep -Milliseconds 300
 
    $ObjForm.Refresh()
    $PB.Value = 50
	$ObjLabel.Text = "Updating App. Extracting Files ..."
	Start-Sleep -Milliseconds 300

 
# Unzip the Archive
Expand-Archive $output -DestinationPath $Path -Force
    
    $ObjForm.Refresh()
    $PB.Value = 70
	$ObjLabel.Text = "Updating App. Files Extracted ..."
	Start-Sleep -Milliseconds 300

    $ObjForm.Refresh()
    $PB.Value = 80
	$ObjLabel.Text = "Updating App. Updating Shortcuts ..."
	Start-Sleep -Milliseconds 300

Start-Sleep -Seconds 2 
cd $HOME
cd desktop
$ShortCutDir = Get-Location

##function set-shortcut {
##param ( [string]$SourceLnk, [string]$DestinationPath )
  ##  $WshShell = New-Object -comObject WScript.Shell
    ##$Shortcut = $WshShell.CreateShortcut($SourceLnk)
    ##$Shortcut.TargetPath = $DestinationPath
    ##$Shortcut.Save()
    ##}
##set-shortcut "$ShortcutDir\BWApp.lnk" "`"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`" C:\BWApp\BWApp-master\Launcher.ps1"

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$ShortCutDir\BWApp.lnk")
$Shortcut.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
$argA = """C:\BWApp\BWApp-master\Launcher.ps1"""
##$argB = """/S:Search Card"""
$Shortcut.Arguments = $argA
##+ " " + $argB
$Shortcut.Save()



    $ObjForm.Refresh()
    $PB.Value = 90
	$ObjLabel.Text = "Updating App. Finishing Update ..."
	Start-Sleep -Milliseconds 300


cd "C:\BWApp\BWApp-master"


Write-Progress -Activity "Updating BWApp" -Status "FINISHED UPDATE" -PercentComplete 100
    $ObjForm.Refresh()
    $PB.Value = 100
	$ObjLabel.Text = "Updating App. FINISHED!"
	Start-Sleep -Milliseconds 300



