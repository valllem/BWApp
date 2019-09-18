﻿if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}

read-host -prompt 'Continue...'
$Path = Get-Content -Path "C:\Temp\BWApp\Path.txt"
write-host $Path
read-host -prompt 'Continue...'
Set-Location -Path $Path
Clear-Host
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Start-Sleep -Seconds 1
Connect-AzureAD -Credential $UserCredential
Connect-MsolService -Credential $UserCredential



Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host                                                                       
        Write-Host -ForegroundColor Yellow ''
        Write-Host -ForegroundColor Yellow ''
        Write-Host -ForegroundColor Magenta '  ███╗   ███╗ █████╗ ██╗███╗   ██╗    ███╗   ███╗███████╗███╗   ██╗██╗   ██╗'
        Write-Host -ForegroundColor Magenta '  ████╗ ████║██╔══██╗██║████╗  ██║    ████╗ ████║██╔════╝████╗  ██║██║   ██║'
        Write-Host -ForegroundColor Magenta '  ██╔████╔██║███████║██║██╔██╗ ██║    ██╔████╔██║█████╗  ██╔██╗ ██║██║   ██║'
        Write-Host -ForegroundColor Magenta '  ██║╚██╔╝██║██╔══██║██║██║╚██╗██║    ██║╚██╔╝██║██╔══╝  ██║╚██╗██║██║   ██║'
        Write-Host -ForegroundColor Magenta '  ██║ ╚═╝ ██║██║  ██║██║██║ ╚████║    ██║ ╚═╝ ██║███████╗██║ ╚████║╚██████╔╝'
        Write-Host -ForegroundColor Magenta '  ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚═╝  ╚═══╝    ╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝ ╚═════╝ '
        Write-Host -ForegroundColor Yellow ''
        Write-Host -ForegroundColor Yellow ''
        
        Write-Host -Object '  1.  TEAMS MANAGEMENT ** Under Development ** ' -ForegroundColor Gray
        Write-Host -Object ''
        Write-Host -Object '  2.  CALENDAR MANAGEMENT ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  3.  MAILBOX MANAGEMENT ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  4.  WINDOWS MANAGEMENT ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  5.  Configuration... ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  6.  365 Security ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  7.  USER MANAGEMENT ' -ForegroundColor Magenta
        Write-Host -Object ''
        Write-Host -Object '  Q.  Quit Application' -ForegroundColor Red
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\menuTeams.ps1            
                anyKey
            }
            2 
            {
                .\menuCalendar.ps1
                anyKey
            }
            3 
            {
                .\menuMailbox.ps1
                anyKey
            }
            4 
            {
                .\menuWindows.ps1
                anyKey
            }
            5 
            {
                .\menuConfiguration.ps1
                anyKey
            }
            6 
            {
                .\menu365Security.ps1
                anyKey
            }
            7 
            {
                .\menuUsers.ps1
                anyKey
            }
            99 
            {
                .\testscript.ps1
                anyKey
            }
            Q 
            {
                Exit-PSSession
                Write-Host -ForegroundColor Green 'Disconnecting from 365'
                Start-Sleep -Seconds 4
                Exit
            }   
            default
            {
                $errout = 'Invalid option please try again........Try 0-3 or Q only'
            }
 
        }
    }
    until ($Menu -eq 'q')
}   
 
# Launch The Menu
Menu

Exit