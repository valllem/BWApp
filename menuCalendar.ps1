

Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host 
        
        
            Write-Host -ForegroundColor Yellow ' ██████╗ █████╗ ██╗     ███████╗███╗   ██╗██████╗  █████╗ ██████╗      █████╗  ██████╗ ██████╗███████╗███████╗███████╗'
            Write-Host -ForegroundColor Yellow '██╔════╝██╔══██╗██║     ██╔════╝████╗  ██║██╔══██╗██╔══██╗██╔══██╗    ██╔══██╗██╔════╝██╔════╝██╔════╝██╔════╝██╔════╝'
            Write-Host -ForegroundColor Yellow '██║     ███████║██║     █████╗  ██╔██╗ ██║██║  ██║███████║██████╔╝    ███████║██║     ██║     █████╗  ███████╗███████╗'
            Write-Host -ForegroundColor Yellow '██║     ██╔══██║██║     ██╔══╝  ██║╚██╗██║██║  ██║██╔══██║██╔══██╗    ██╔══██║██║     ██║     ██╔══╝  ╚════██║╚════██║'
            Write-Host -ForegroundColor Yellow '╚██████╗██║  ██║███████╗███████╗██║ ╚████║██████╔╝██║  ██║██║  ██║    ██║  ██║╚██████╗╚██████╗███████╗███████║███████║'
            Write-Host -ForegroundColor Yellow ' ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═══╝╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝    ╚═╝  ╚═╝ ╚═════╝ ╚═════╝╚══════╝╚══════╝╚══════╝'
            Write-Host -ForegroundColor Yellow ''
            Write-Host -ForegroundColor Yellow '======================================================================================================================'
            Write-Host -ForegroundColor Yellow ''
            Write-Host -ForegroundColor Yellow ''     

        Write-Host -Object '1.  Give user access to someone elses Calendar '
        Write-Host -Object ''
        Write-Host -Object '2.  Give a user access to ALL Calendars '
        Write-Host -Object ''
        Write-Host -Object '3.  Remove Calendar Access '
        Write-Host -Object ''
        Write-Host -Object '4.  Check Calendar Access '
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\CalendarAccess1.ps1            
                anyKey
            }
            2 
            {
                .\CalendarAccessAll.ps1
                anyKey
            }
            3 
            {
                .\CalendarAccessRemove.ps1
                anyKey
            }
            4 
            {
                .\CalendarCheck.ps1
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