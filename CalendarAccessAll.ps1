


Write-Host -ForegroundColor Yellow '   █████╗ ██╗     ██╗          ██████╗ █████╗ ██╗     ███████╗███╗   ██╗██████╗  █████╗ ██████╗ ███████╗'
Write-Host -ForegroundColor Yellow '  ██╔══██╗██║     ██║         ██╔════╝██╔══██╗██║     ██╔════╝████╗  ██║██╔══██╗██╔══██╗██╔══██╗██╔════╝'
Write-Host -ForegroundColor Yellow '  ███████║██║     ██║         ██║     ███████║██║     █████╗  ██╔██╗ ██║██║  ██║███████║██████╔╝███████╗'
Write-Host -ForegroundColor Yellow '  ██╔══██║██║     ██║         ██║     ██╔══██║██║     ██╔══╝  ██║╚██╗██║██║  ██║██╔══██║██╔══██╗╚════██║'
Write-Host -ForegroundColor Yellow '  ██║  ██║███████╗███████╗    ╚██████╗██║  ██║███████╗███████╗██║ ╚████║██████╔╝██║  ██║██║  ██║███████║'
Write-Host -ForegroundColor Yellow '  ╚═╝  ╚═╝╚══════╝╚══════╝     ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═══╝╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝'
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''                                                                                                      
Write-Host -ForegroundColor Yellow '==========================================================================================================='


##-- Prompt to enter user email that needs access--##

$userRequiringAccess = Read-Host 'Please enter the email address of the person who will be given access to all calendars...'


##--Run Script--##

$accessRight = "PublishingAuthor"
 
$mailboxes = Get-mailbox
$userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -erroraction SilentlyContinue
         
    if ($accessRights.accessRights -notmatch $accessRight -and $mailbox.primarysmtpaddress -notcontains $userRequiringAccess.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress
        if ($accessRights.accessRights -match $accessRight) {
            Write-Host "Successfully added $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Green
        }
        else {
            Write-Host "Could not add $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Red
        }
    }else{
        Write-Host "Permission level already exists for $($userrequiringaccess.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
    }
}



Write-Host -ForegroundColor Green '======================================='
Write-Host -ForegroundColor Green '         All Tasks Complete!           '


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()