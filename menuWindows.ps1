Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host 
        
        
 Write-Host -ForegroundColor Yellow ' _       ___           __                      __  ___                                                  __ '
 Write-Host -ForegroundColor Yellow '| |     / (_)___  ____/ /___ _      _______   /  |/  /___ _____  ____ _____ ____  ____ ___  ___  ____  / /_'
 Write-Host -ForegroundColor Yellow '| | /| / / / __ \/ __  / __ \ | /| / / ___/  / /|_/ / __ `/ __ \/ __ `/ __ `/ _ \/ __ `__ \/ _ \/ __ \/ __/'
 Write-Host -ForegroundColor Yellow '| |/ |/ / / / / / /_/ / /_/ / |/ |/ (__  )  / /  / / /_/ / / / / /_/ / /_/ /  __/ / / / / /  __/ / / / /_  '
 Write-Host -ForegroundColor Yellow '|__/|__/_/_/ /_/\__,_/\____/|__/|__/____/  /_/  /_/\__,_/_/ /_/\__,_/\__, /\___/_/ /_/ /_/\___/_/ /_/\__/  '
 Write-Host -ForegroundColor Yellow '                                                                   /____/                                  '
 Write-Host -ForegroundColor Yellow '==========================================================================================================='
 Write-Host -ForegroundColor Yellow ''                                                                             
        
        Write-Host -Object '1.  Reinstall Packages (Win10 Apps) '
        Write-Host -Object ''
        Write-Host -Object '2.  DISM Cleanup Image '
        Write-Host -Object ''
        Write-Host -Object '3.  Check for Windows Updates '
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\windowsApps.ps1            
                anyKey
            }
            2 
            {
                .\windowsDISM.ps1
                anyKey
            }
            3 
            {
                .\windowsUpdate.ps1
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