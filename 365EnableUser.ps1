Clear-Host
$colourInfo = "green"
$colourAlert = "yellow"
$colourError = "red"

write-host -foregroundcolor $colourAlert "Enable a User"

$useremail=read-host -prompt 'Enter user email address'
try {   ## See whether input matches a user in Azure AD
    $user= get-azureaduser -objectid $useremail -erroraction stop
    write-host -ForegroundColor $colourAlert "Reset user's password on local AD if synced or in cloud"
    Read-Host -prompt "Press any key to continue" 
}
catch
{       # if there is no match provide warning and terminate script
 Â Â Â write-host $useremail,"doesn't appear to be a valid AD user account" -ForegroundColor Red
    return
}

write-host -foregroundcolor $colourInfo "Found",$user.displayname

## Disable account to block user logins
Set-AzureADUser -objectid $user.ObjectId -AccountEnabled $true
write-host -foregroundcolor $colourInfo "Enabled login"
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId
write-host -foregroundcolor $colourAlert "Revoked Token"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
Set-CASMailbox $useremail -ActiveSyncEnabled $true
write-host -foregroundcolor $colourAlert "ActiveSync enabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
Set-casmailbox $useremail -owaenabled $true
write-host -foregroundcolor $colourAlert "OWA enabled"


## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
Set-casmailbox $useremail -mapienabled $true
write-host -foregroundcolor $colourAlert "Enabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
Set-casmailbox $useremail -OWAforDevicesEnabled $true
write-host -foregroundcolor $colourAlert "Enabled MAPI for devices"



## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
Set-casmailbox $useremail -universaloutlookenabled $true
write-host -foregroundcolor $colourAlert "Enabled Outlook"

## User will be signed out of browser, desktop and mobile applications accessing Office 365 resources across all devices.
## It can take up to an hour to sign out from all devices.



write-host -ForegroundColor $colourAlert "Completed all Tasks"
Read-Host -prompt "Press any key to continue" 