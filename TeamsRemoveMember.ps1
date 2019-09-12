Clear-Host


        Write-Host -ForegroundColor Yellow '  ████████╗███████╗ █████╗ ███╗   ███╗███████╗'
        Write-Host -ForegroundColor Yellow '  ╚══██╔══╝██╔════╝██╔══██╗████╗ ████║██╔════╝'
        Write-Host -ForegroundColor Yellow '     ██║   █████╗  ███████║██╔████╔██║███████╗   '
        Write-Host -ForegroundColor Yellow '     ██║   ██╔══╝  ██╔══██║██║╚██╔╝██║╚════██║   '
        Write-Host -ForegroundColor Yellow '     ██║   ███████╗██║  ██║██║ ╚═╝ ██║███████║   '
        Write-Host -ForegroundColor Yellow '     ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝   '
        Write-Host -ForegroundColor Yellow ''



Write-Host -ForegroundColor Yellow '================================='
Write-Host -ForegroundColor Yellow '|        CREATE NEW TEAM        |'
Write-Host -ForegroundColor Yellow '================================='
Write-Host -ForegroundColor Yellow '                                 '
$TeamName = Read-Host "Please enter the Name of the new Team"
New-Team -DisplayName $TeamName

Write-Host -ForegroundColor Yellow '                                 '
Write-Host -ForegroundColor Yellow '                                 '
Write-Host -ForegroundColor Green '================================='
Write-Host -ForegroundColor Green '' $TeamName ''
Write-Host -ForegroundColor Green 'CREATED'
Write-Host -ForegroundColor Green '================================='
Write-host ' '
Write-host ' '
Write-host ' '
Write-host 'Press any Key to return to the Menu'
##-- Waits for user input before continuing --##
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()