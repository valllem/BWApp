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


##-- Prompt for mailbox the user will check--##
Write-Host -ForegroundColor Yellow "Who's Calendar do you want to Check?"
$UserIdentity = Read-Host 'Enter Users Email Address...'

Get-MailboxFolderPermission -identity ($UserIdentity+':\calendar')


write-Host -ForegroundColor Yellow ''
write-Host -ForegroundColor Yellow 'Press any key to return to previous menu...'
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()
Exit