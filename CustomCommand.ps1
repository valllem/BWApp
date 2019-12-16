


if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
  
 }
}


Clear-Host
Start-Sleep -Seconds 2
Write-Host -ForegroundColor Yellow 'Checking if required Modules are installed.'
Start-Sleep -Seconds 2
  
  
  
  
  try {
##--Security Principal Check for If statement--##
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
Start-Sleep -Seconds 2


$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Connect-AzureAD -Credential $UserCredential

Write-Host -ForegroundColor Green 'Ready...'




}
catch { ## this is where any error handling will appear if required
write-host -foregroundcolor Yellow "Please run PowerShell as an Administrator"
Start-Sleep -Seconds 2
Read-Host -prompt 'Press any key to quit'
}
























