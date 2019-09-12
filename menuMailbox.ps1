Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host                                                                       
        Write-Host -Object '  '
        Write-Host     -Object '**********************'
        Write-Host -Object 'MAILBOX MANAGEMENT' -ForegroundColor Yellow
        Write-Host     -Object '**********************'
        Write-Host -Object '1.  Grant User Full Access '
        Write-Host -Object ''
        Write-Host -Object '2.  Remove User Full Access '
        Write-Host -Object ''
        Write-Host -Object '3.  Grant User Send As Permissions '
        Write-Host -Object ''
        Write-Host -Object '4.  Forward Emails '
        Write-Host -Object ''
        Write-Host -Object '5.  Grant User Full Access to ALL MAILBOXES '
        Write-Host -Object ''
        Write-Host -Object '6.  Remove User Full Access to ALL MAILBOXES '
        Write-Host -Object ''
        Write-Host -Object 'Q.  Quit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or Q to Quit)'
 
        switch ($Menu) 
        {
           1 
            {
                .\MailboxGrantFull.ps1            
                anyKey
            }
            2 
            {
                .\MailboxRemoveFull.ps1
                anyKey
            }
            3 
            {
                .\MailboxSend.ps1
                anyKey
            }
            4 
            {
                .\MailboxForward.ps1
                anyKey
            }
            5 
            {
                .\MailboxGrantALL.ps1
                anyKey
            }
            6 
            {
                .\MailboxRemoveALL.ps1
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