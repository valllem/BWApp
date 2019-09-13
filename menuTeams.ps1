Connect-MicrosoftTeams -Credential $UserCredential
Function Menu 
{
          
    Do
    {
        Clear-Host                                                                       
        Write-Host -ForegroundColor Yellow '  ████████╗███████╗ █████╗ ███╗   ███╗███████╗'
        Write-Host -ForegroundColor Yellow '  ╚══██╔══╝██╔════╝██╔══██╗████╗ ████║██╔════╝'
        Write-Host -ForegroundColor Yellow '     ██║   █████╗  ███████║██╔████╔██║███████╗   '
        Write-Host -ForegroundColor Yellow '     ██║   ██╔══╝  ██╔══██║██║╚██╔╝██║╚════██║   '
        Write-Host -ForegroundColor Yellow '     ██║   ███████╗██║  ██║██║ ╚═╝ ██║███████║   '
        Write-Host -ForegroundColor Yellow '     ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝   '
        Write-Host -ForegroundColor Yellow ''

        Write-Host -Object '1.  Create Team '
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\TeamsCreate.ps1            
                anyKey
            }
            2 
            {
                .\TeamsAddMember.ps1
                anyKey
            }
            3 
            {
                .\TeamsRemoveMember.ps1
                anyKey
            }
            Q 
            {
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