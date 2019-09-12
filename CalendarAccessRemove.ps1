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
            Write-Host -ForegroundColor Yellow 'Remove Calendar Access'
            Write-Host -ForegroundColor Yellow '======================'
            Write-Host -ForegroundColor Yellow '                      '



##-- Prompt for mailbox the user will access--##
$UserIdentity = Read-Host 'Enter email of account to Manage'

##-- Prompt to enter user email that needs access--##
$UserEmail = Read-Host 'Who are you removing from' $UserIdentity '?'

##-- Shows Details --##

Write-Host -ForegroundColor Yellow ' The account' $UserEmail 'will be removed from accessing' $UserIdentity"'s calendar"
##--Run Script--##

Remove-MailboxFolderPermission -Identity ($UserIdentity+':\calendar') -user $UserEmail -Confirm

##--Show result--##
Write-Host -ForegroundColor Green ' ' $UserEmail 'will no longer be able to access' ($UserIdentity+"'s calendar.") 
write-Host -ForegroundColor Green 'Please allow a few minutes for the changes to take effect.'
write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow 'Press any key to return to previous menu...'


$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()