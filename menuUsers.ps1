Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host                                                                       
        Write-Host -Object '  '
        Write-Host     -Object '**********************'
        Write-Host -Object 'USER MANAGEMENT' -ForegroundColor Yellow
        Write-Host     -Object '**********************'
        Write-Host -Object '1.  Rename Users Primary UPN '
        Write-Host -Object ''
        Write-Host -Object '2.  Rename First & Last Names '
        Write-Host -Object ''
        Write-Host -Object '3.  Block a User (Signs out and disables all access)'
        Write-Host -Object ''
        Write-Host -Object '4.  Enable a User Account (Does not enable pop/IMAP)'
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\RenameSyncedUsers.ps1            
                anyKey
            }
            2 
            {
                .\RenameFirstLastName.ps1            
                anyKey
            }
            3 
            {
                .\365BlockUser.ps1            
                anyKey
            }
            4 
            {
                .\365EnableUser.ps1            
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