Clear-host 
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




##-- Prompt for mailbox the user will access--##
Write-Host -ForegroundColor Yellow "Which Calendar do you want to Manage?"
$UserIdentity = Read-Host 'Enter Users Email Address...'

##-- Prompt to enter user email that needs access--##
Write-Host -ForegroundColor Yellow "Which User needs access?"
$UserEmail = Read-Host 'Enter Email Address...'

##--
$accessRights = $null
$accessRights = Get-MailboxFolderPermission ($mailbox+':\calendar') -User $userRequiringAccess -erroraction SilentlyContinue
$mailboxes = $UserIdentity
$mailbox = $UserIdentity
$userRequiringAccess = $UserEmail
$LimitedDetails = "LimitedDetails"
$Author = "Author"
$PublishingEditor = "PublishingEditor"
$Owner = "Owner"





Write-Host -ForegroundColor Yellow "Please select the permissions you want to give" $UserEmail
Write-Host -ForegroundColor Yellow "=============================================="

Function Menu 
{
            
    Do
    {
                                                                               
        Write-Host -Object '  '
        Write-Host     -Object '**********************'
        Write-Host -Object 'SELECT PERMISSION' -ForegroundColor Yellow
        Write-Host     -Object '**********************'
        Write-Host -Object '1.  Limited Details (Subject and Location) '
        Write-Host -Object ''
        Write-Host -Object '2.  Author (Create and Edit own items) '
        Write-Host -Object ''
        Write-Host -Object '3.  Publishing Editor (Create and Edit ALL items) '
        Write-Host -Object ''
        Write-Host -Object '4.  Owner (Same rights as the Owner) '
        Write-Host -Object ''
        Write-Host -Object 'M.  Cancel & Return to previous Menu'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt '(0-3 or M to return to previous Menu)'
 
        switch ($Menu) 
        {
           1 
            {
            if ($accessRights.accessRights -notmatch $LimitedDetails -and $mailbox -notcontains $userRequiringAccess -and $mailbox -notmatch "DiscoverySearchMailbox") {
                Write-Host "Adding or updating permissions for $mailbox Calendar" -ForegroundColor Yellow
            try {
                Add-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights LimitedDetails -ErrorAction SilentlyContinue    
            }
            catch {
                Set-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights LimitedDetails -ErrorAction SilentlyContinue
            }        
            $accessRights = Get-MailboxFolderPermission ($mailbox+':\calendar') -User $userRequiringAccess
            if ($accessRights.accessRights -match $LimitedDetails) {
                Write-Host "Successfully added LimitedDetails permissions on $mailbox's calendar for $userrequiringaccess" -ForegroundColor Green
            }
            else {
                Write-Host "$userrequiringaccess already has other permissions on $mailbox's calendar." -ForegroundColor Red
            }
            }else{
                Write-Host "Permission level already exists for $userrequiringaccess on" $mailbox"'s calendar" -foregroundColor Green
            }


                $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
                $HOST.UI.RawUI.Flushinputbuffer()
                Exit
            }
            2 
            {
            if ($accessRights.accessRights -notmatch $Author -and $mailbox -notcontains $userRequiringAccess -and $mailbox -notmatch "DiscoverySearchMailbox") {
                Write-Host "Adding or updating permissions for $mailbox Calendar" -ForegroundColor Yellow
            try {
                Add-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights Author -ErrorAction SilentlyContinue    
            }
            catch {
                Set-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights Author -ErrorAction SilentlyContinue
            }        
            $accessRights = Get-MailboxFolderPermission ($mailbox+':\calendar') -User $userRequiringAccess
            if ($accessRights.accessRights -match $Author) {
                Write-Host "Successfully added Author permissions on $mailbox's calendar for $userrequiringaccess" -ForegroundColor Green
            }
            else {
                Write-Host "$userrequiringaccess already has other permissions on $mailbox's calendar." -ForegroundColor Red
            }
            }else{
                Write-Host "Permission level already exists for $userrequiringaccess on" $mailbox"'s calendar" -foregroundColor Green
            }


                $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
                $HOST.UI.RawUI.Flushinputbuffer()
                Exit
            }
            3 
            {
            if ($accessRights.accessRights -notmatch $PublishingEditor -and $mailbox -notcontains $userRequiringAccess -and $mailbox -notmatch "DiscoverySearchMailbox") {
                Write-Host "Adding or updating permissions for $mailbox Calendar" -ForegroundColor Yellow
            try {
                Add-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights PublishingEditor -ErrorAction SilentlyContinue    
            }
            catch {
                Set-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights PublishingEditor -ErrorAction SilentlyContinue
            }        
            $accessRights = Get-MailboxFolderPermission ($mailbox+':\calendar') -User $userRequiringAccess
            if ($accessRights.accessRights -match $PublishingEditor) {
                Write-Host "Successfully added PublishingEditor permissions on $mailbox's calendar for $userrequiringaccess" -ForegroundColor Green
            }
            else {
                Write-Host "$userrequiringaccess already has other permissions on $mailbox's calendar." -ForegroundColor Red
            }
            }else{
                Write-Host "Permission level already exists for $userrequiringaccess on" $mailbox"'s calendar" -foregroundColor Green
            }


                $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
                $HOST.UI.RawUI.Flushinputbuffer()
                Exit
            }
            4 
            {
            if ($accessRights.accessRights -notmatch $Owner -and $mailbox -notcontains $userRequiringAccess -and $mailbox -notmatch "DiscoverySearchMailbox") {
                Write-Host "Adding or updating permissions for $mailbox Calendar" -ForegroundColor Yellow
            try {
                Add-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights Owner -ErrorAction SilentlyContinue    
            }
            catch {
                Set-MailboxFolderPermission ($UserIdentity+':\calendar') -User $userRequiringAccess -AccessRights Owner -ErrorAction SilentlyContinue
            }        
            $accessRights = Get-MailboxFolderPermission ($mailbox+':\calendar') -User $userRequiringAccess
            if ($accessRights.accessRights -match $Owner) {
                Write-Host "Successfully added Owner permissions on $mailbox's calendar for $userrequiringaccess" -ForegroundColor Green
            }
            else {
                Write-Host "$userrequiringaccess already has other permissions on $mailbox's calendar." -ForegroundColor Red
            }
            }else{
                Write-Host "Permission level already exists for $userrequiringaccess on" $mailbox"'s calendar" -foregroundColor Green
            }


                $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
                $HOST.UI.RawUI.Flushinputbuffer()
                Exit
            }
            M 
            {
                Exit
            }   
            default
            {
                $errout = 'Invalid option please try again........Try 0-3 or m only'
            }
 
        }
    }
    until ($Menu -eq 'm')
}   
 
# Launch The Menu
Menu


