Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host                                                                       
        Write-Host -Object '  '
        Write-Host -Object '**********************'
        Write-Host -Object '   365 Security   ' -ForegroundColor Yellow
        Write-Host -Object '**********************'
        Write-Host -Object '1.  Enable Audit Logging '
        Write-Host -Object ''
        Write-Host -Object '2.  Disable Audit Logging '
        Write-Host -Object ''
        Write-Host -Object '3.  CHECK Audit Logging status '
        Write-Host -Object ''
        Write-Host -Object '4.  Block EMAIL '
        Write-Host -Object ''
        Write-Host -Object '5.  Block DOMAIN '
        Write-Host -Object ''
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\365EnableAuditLog.ps1            
                anyKey
            }
            2 
            {
                .\365DisableAuditLog.ps1
                anyKey
            }
            3 
            {
                .\365CheckAuditLog.ps1
                anyKey
            }
            4 
            {
                .\365BlockSender.ps1
                anyKey
            }
            5 
            {
                .\365BlockDomain.ps1
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