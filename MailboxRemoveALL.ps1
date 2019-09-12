Clear-Host


Write-Host -ForegroundColor Yellow ' █████╗ ██╗     ██╗         ███╗   ███╗ █████╗ ██╗██╗     ██████╗  ██████╗ ██╗  ██╗███████╗███████╗'
Write-Host -ForegroundColor Yellow '██╔══██╗██║     ██║         ████╗ ████║██╔══██╗██║██║     ██╔══██╗██╔═══██╗╚██╗██╔╝██╔════╝██╔════╝'
Write-Host -ForegroundColor Yellow '███████║██║     ██║         ██╔████╔██║███████║██║██║     ██████╔╝██║   ██║ ╚███╔╝ █████╗  ███████╗'
Write-Host -ForegroundColor Yellow '██╔══██║██║     ██║         ██║╚██╔╝██║██╔══██║██║██║     ██╔══██╗██║   ██║ ██╔██╗ ██╔══╝  ╚════██║'
Write-Host -ForegroundColor Yellow '██║  ██║███████╗███████╗    ██║ ╚═╝ ██║██║  ██║██║███████╗██████╔╝╚██████╔╝██╔╝ ██╗███████╗███████║'
Write-Host -ForegroundColor Yellow '╚═╝  ╚═╝╚══════╝╚══════╝    ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝╚═════╝  ╚═════╝ ╚═╝  ╚═╝╚══════╝╚══════╝'
Write-Host -ForegroundColor Yellow ''
Write-Host -ForegroundColor Yellow ''                                                                                                      
Write-Host -ForegroundColor Yellow '==========================================================================================================='


##-- Prompt to enter user email that needs access--##

$userRequiringAccess = Read-Host 'Please enter the email address of the person who will be removed from all mailboxes...'


##--Run Script--##

$accessRight = "FullAccess"
 
$mailboxes = Get-mailbox
$userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    Remove-MailboxPermission "$($mailbox.primarysmtpaddress)" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -Force -ErrorAction SilentlyContinue
    Write-Host -ForegroundColor Yellow "Removing $userRequiringAccess from $mailbox "
}



Write-Host -ForegroundColor Green '======================================='
Write-Host -ForegroundColor Green '         All Tasks Complete!           '


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()