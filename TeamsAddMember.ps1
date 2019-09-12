Clear-Host

        Write-Host -ForegroundColor Yellow '  ████████╗███████╗ █████╗ ███╗   ███╗███████╗'
        Write-Host -ForegroundColor Yellow '  ╚══██╔══╝██╔════╝██╔══██╗████╗ ████║██╔════╝'
        Write-Host -ForegroundColor Yellow '     ██║   █████╗  ███████║██╔████╔██║███████╗   '
        Write-Host -ForegroundColor Yellow '     ██║   ██╔══╝  ██╔══██║██║╚██╔╝██║╚════██║   '
        Write-Host -ForegroundColor Yellow '     ██║   ███████╗██║  ██║██║ ╚═╝ ██║███████║   '
        Write-Host -ForegroundColor Yellow '     ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝   '
        Write-Host -ForegroundColor Yellow ''

        Write-Host -ForegroundColor Yellow '================================='
        Write-Host -ForegroundColor Yellow '|            Add Member         |'
        Write-Host -ForegroundColor Yellow '================================='
        Write-Host -ForegroundColor Yellow '                                 '



$users = Get-Team
    if ($users.count -gt 1){
        Write-Host "Multiple users were found" -ForegroundColor Yellow
        Write-Host "Please select a user" -ForegroundColor Yellow
        for($i = 0; $i -lt $users.count; $i++){
            Write-Host "$($i): $($users[$i].DisplayName)"
        }
        $selection = Read-Host -Prompt "Enter the number of the user"
        $final_user = $users[$selection]
    }



Write-Host '' $final_user  ''

Write-Host -ForegroundColor Yellow '                                 '
Write-Host -ForegroundColor Yellow '                                 '
Write-Host -ForegroundColor Green '================================='
Write-Host -ForegroundColor Green '        ' $TeamName '            '
Write-Host -ForegroundColor Green '           CREATED               '
Write-Host -ForegroundColor Green '================================='
Write-host ' '
Write-host ' '
Write-host ' '
Write-host 'Press any Key to return to the Menu'
##-- Waits for user input before continuing --##
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()