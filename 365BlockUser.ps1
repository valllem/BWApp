Clear-Host

$colourInfo = "green"
$colourAlert = "yellow"
$colourError = "red"

write-host -foregroundcolor $colourAlert "Block a User"

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
Set-AzureADUser -objectid $user.ObjectId -AccountEnabled $false
write-host -foregroundcolor $colourInfo "Disabled login"
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId
write-host -foregroundcolor $colourAlert "Revoked Token"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
Set-CASMailbox $useremail -ActiveSyncEnabled $false
write-host -foregroundcolor $colourAlert "ActiveSync disabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
Set-casmailbox $useremail -owaenabled $false
write-host -foregroundcolor $colourAlert "OWA disabled"

## TheActiveSyncAllowedDeviceIDs parameter specifies one or more Exchange ActiveSync device IDs that are allowed to synchronize with the mailbox.
## Setting this to $NULL clears the list of device IDs
Set-casmailbox $useremail -activesyncalloweddeviceids $null
write-host -foregroundcolor $colourAlert "Removed allowed ActiveSync devices"

## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
Set-casmailbox $useremail -mapienabled $false
write-host -foregroundcolor $colourAlert "Disabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
Set-casmailbox $useremail -OWAforDevicesEnabled $false
write-host -foregroundcolor $colourAlert "Disabled MAPI fo devices"

## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
Set-casmailbox $useremail -popenabled $false
write-host -foregroundcolor $colourAlert "Disabled POP"

## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
Set-casmailbox $useremail -imapenabled $false
write-host -foregroundcolor $colourAlert "Disabled IMAP"

## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
Set-casmailbox $useremail -universaloutlookenabled $false
write-host -foregroundcolor $colourAlert "Disabled Outlook"

## User will be signed out of browser, desktop and mobile applications accessing Office 365 resources across all devices.
## It can take up to an hour to sign out from all devices.
Revoke-SPOUserSession -user $useremail -Confirm:$false


write-host -ForegroundColor $colourAlert "Completed all Tasks"
Read-Host -prompt "Press any key to continue" 

