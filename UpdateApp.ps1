$versioncheckurl = "https://raw.githubusercontent.com/valllem/BWApp/master/version.dll"
$versionoutput = "C:\BWapp\bin\versioncheck.dll"
(New-Object System.Net.WebClient).DownloadFile($versioncheckurl, $versionoutput)
$versionavailable = Get-Content $versionoutput
$versioncurrent = Get-Content "C:\BWApp\BWApp-master\version.dll"
$url = "https://github.com/valllem/BWApp/archive/master.zip"
$Path = "C:\BWApp"
$output = [IO.Path]::Combine($Path, "BWApp.zip”)




Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '519,245'
$Form.text                       = "Update BWApp"
$Form.TopMost                    = $false

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 481
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(16,82)

$LabelStatus                     = New-Object system.Windows.Forms.Label
$LabelStatus.text                = "Status:"
$LabelStatus.AutoSize            = $true
$LabelStatus.width               = 350
$LabelStatus.height              = 10
$LabelStatus.location            = New-Object System.Drawing.Point(16,32)
$LabelStatus.Font                = 'Microsoft Sans Serif,10'

$ButtonUpdate                    = New-Object system.Windows.Forms.Button
$ButtonUpdate.text               = "Update Now"
$ButtonUpdate.width              = 126
$ButtonUpdate.height             = 30
$ButtonUpdate.location           = New-Object System.Drawing.Point(367,200)
$ButtonUpdate.Font               = 'Microsoft Sans Serif,10,style=Bold'
$ButtonUpdate.ForeColor          = "#7ed321"

$Form.controls.AddRange(@($ProgressBar1,$LabelStatus,$ButtonUpdate))


if ($versionavailable -notmatch $versioncurrent){
$LabelStatus.text                = "Status: UPDATE AVAILABLE"
}
elseif ($versionavailable -match $versioncurrent) {
$LabelStatus.text                = "Status: No updates available"
}

$ButtonUpdate.Add_Click({  


#Download Files
$ProgressBar1.Value = 10
$LabelStatus.Text = "Status: Downloading latest copy of files"

(New-Object System.Net.WebClient).DownloadFile($url, $output)



#Extract Files
$ProgressBar1.Value = 10
$LabelStatus.Text = "Status: Extracting files"

Expand-Archive $output -DestinationPath $Path -Force



#Update Shortcut
$ProgressBar1.Value = 10
$LabelStatus.Text = "Status: Updating Shortcut"
Start-Sleep -Seconds 2 
cd $HOME
cd desktop
$ShortCutDir = Get-Location
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$ShortCutDir\BWApp.lnk")
$Shortcut.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
$argA = """C:\BWApp\BWApp-master\Launcher.ps1"""
##$argB = """/S:Search Card"""
$Shortcut.Arguments = $argA
##+ " " + $argB
$Shortcut.Save()

#Change DIR
cd "C:\BWApp\BWApp-master"

#Install or update Modules
$ProgressBar1.Value = 20
$LabelStatus.Text = "Status: Checking Modules"

####    
if (Get-Module -ListAvailable -Name AADRM) {
    Write-Host -ForegroundColor Green "Azure AD Right Management Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD Right Management module"
    $ProgressBar1.Value = 20
    $LabelStatus.Text = "Status: Installing AADRM"
    Install-Module -Name AADRM -force
}
####
if (Get-Module -ListAvailable -Name AzureAD) {
    Write-Host -ForegroundColor Green "Azure AD Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure AD module"
    $ProgressBar1.Value = 30
    $LabelStatus.Text = "Status: Installing AzureAD"
    Install-Module -Name AzureAD -force -AllowClobber
}
####  
if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    Write-Host -ForegroundColor Green "Teams Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Teams Module"
    $ProgressBar1.Value = 40
    $LabelStatus.Text = "Status: Installing MicrosoftTeams"
    Install-Module -Name MicrosoftTeams -Force
}
####
Start-Sleep -Milliseconds 300
#micro pause to prevent 365 forcing connection pause

if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    Write-Host -ForegroundColor Green "SharePoint Online Module exists, updating"
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing SharePoint Online module"
    $ProgressBar1.Value = 50
    $LabelStatus.Text = "Status: Installing SharePoint Online Module"
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force
}
####    
if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host -ForegroundColor Green "Microsoft Online Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Microsoft Online module"
    $ProgressBar1.Value = 60
    $LabelStatus.Text = "Status: Installing MSOnline"
    Install-Module -Name MSOnline -force
}
####
Start-Sleep -Milliseconds 300
#micro pause to prevent 365 forcing connection pause
if (Get-Module -ListAvailable -Name AzureRM) {
    Write-Host -ForegroundColor Green "Azure Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Azure module... This may take a few minutes"
    $ProgressBar1.Value = 70
    $LabelStatus.Text = "Status: Installing AzureRM"
    Install-Module -name AzureRM -Force
}    
####
if (Get-Module -ListAvailable -Name CreateExoPsSession) {
    Write-Host -ForegroundColor Green "Exchange MFA Module exists, updating."
    
} 
else {
    Write-Host -foregroundcolor Yellow "Module does not exist..."
    write-host -foregroundcolor Yellow "Installing Exchange MFA module... This may take a few minutes"
    $ProgressBar1.Value = 80
    $LabelStatus.Text = "Status: Installing CreateExoPsSession"
    Install-Script -Name CreateExoPsSession -force 
}
####
    $ProgressBar1.Value = 90
    $LabelStatus.Text = "Status: Installing Load-ExchangeMFA"
    Install-Script -Name Load-ExchangeMFA

    
    $ProgressBar1.Value = 100
    $LabelStatus.Text = "Status: Completed Update!"


})
$Form.ShowDialog()
