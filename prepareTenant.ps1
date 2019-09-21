if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}

$colourInfo = "green"
$colourAlert = "cyan"
$colourError = "red"
$errormessagecolor = "red"
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='


## Information
write-host -ForegroundColor $colourAlert 'First we need some information'
write-host -ForegroundColor $colourAlert '=============================='

write-host -ForegroundColor Yellow 'This script will require you to login several times...'
write-host -ForegroundColor Yellow 'After running this script, please quit the App and relaunch'
read-host -Prompt 'Press Enter to continue...'




## LOGIN
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking -AllowClobber
Start-Sleep -Seconds 1
Connect-AzureAD -Credential $UserCredential
Connect-MsolService -Credential $UserCredential





$domains=read-host -prompt 'Enter Tenant Email DOMAIN for Spam and Malware Policy...'



Write-Host
Write-Host
Write-Host 'Enabling Organization Customization'
Enable-OrganizationCustomization

Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Start-Sleep -Seconds 2
##Enable Modern Auth
Write-Host
Write-Host
Write-Host
Write-Host
write-host -foregroundcolor $colourAlert "Enabling Modern Authentication"
write-host
write-host
write-host

Write-Host -foregroundcolor $colourAlert "Enter your Sharepoint url E.g https://busworks-admin.sharepoint.com"
##Import connection to sharepoint - this will request the admin url E.g https://busworks-admin.sharepoint.com
Connect-SPOService -Credential $UserCredential
Start-Sleep -Seconds 2
$org=get-organizationconfig
write-host -ForegroundColor $colourAlert "Exchange setting is currently",$org.OAuth2ClientProfileEnabled
## Run this command to enable modern authentication for Exchange Online
Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
write-host -foregroundcolor $colourInfo "Exchange command completed"
$org=get-organizationconfig
write-host -ForegroundColor $colourInfo "Exchange setting updated to",$org.OAuth2ClientProfileEnabled
## To set Sharepoint apps that dont use modern authentication to block
$org=Get-SPOTenant
write-host -ForegroundColor $colourAlert "SharePoint setting is currently",$org.legacyauthprotocolsenabled
Set-spotenant -legacyauthprotocolsenabled $false
$org=Get-SPOTenant
write-host -ForegroundColor $colourInfo "SharePoint setting is updated to",$org.legacyauthprotocolsenabled


Start-Sleep -Seconds 4


## Audit Logging Enabling
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host -ForegroundColor $colourAlert 'Enabling Audit Logging'
Start-Sleep -Seconds 2

Write-Host -ForegroundColor Yellow 'This can take several seconds'


Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true

Write-Host -ForegroundColor $colourAlert 'Audit Logging Enabled'
Start-Sleep -Seconds 2

Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host -ForegroundColor $colourAlert 'Enabling Audit Logging'
Start-Sleep -Seconds 2


write-host -ForegroundColor $colourAlert "Getting shared mailboxes"
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited
write-host -ForegroundColor $colourAlert "Start checking shared mailboxes"
write-host
foreach ($mailbox in $mailboxes) {
    $accountdetails=get-azureaduser -objectid $mailbox.userprincipalname        ## Get the Azure AD account connected to shared mailbox
    If ($accountdetails.accountenabled){                                        ## if that login is enabled
        Write-host -foregroundcolor $colourError $mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
        If ($secure) {                                                          ## if the secure variable is true disable login to shared mailbox
            Set-AzureADUser -ObjectID $mailbox.userprincipalname -AccountEnabled $false     ## disable shared mailbox account
            $accountdetails=get-azureaduser -objectid $mailbox.userprincipalname            ## Get the Azure AD account connected to shared mailbox again
            write-host -ForegroundColor $colourAlert "*** SECURED"$mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
        }
    } else {
        Write-host -foregroundcolor $colourAlert $mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
    }
}
write-host -ForegroundColor $colourInfo "`nFinish checking mailboxes"












## POP/IMAP

Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Start-Sleep -Seconds 2
write-host -ForegroundColor $colourInfo 'Disabling POP / IMAP for any existing mailboxes'
Start-Sleep -Seconds 2
## Lock Down each existing mailbox
$mailboxes = Get-mailbox
foreach ($mailbox in $mailboxes) {
             
    ## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -popenabled $false
    write-host -foregroundcolor $colourInfo "Disabled POP for " $mailbox ""

    ## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -imapenabled $false
    write-host -foregroundcolor $colourInfo "Disabled IMAP for " $mailbox ""
    write-host -foregroundcolor $colourAlert ""
    write-host -foregroundcolor $colourAlert ""
}

## Set Default POP IMAP settings

write-host -ForegroundColor $colourInfo 'Setting default POP / IMAP settings to DISABLED'
Start-Sleep -Seconds 2
Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" } | set-CASMailboxPlan -ImapEnabled $false -PopEnabled $false





## Configure the Spam Policy.

$policyname = "BusinessWorks Policy"
$rulename = "BusinessWorks Configured Recpients"
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
write-host -foregroundcolor $colourAlert "Setting new spam policy"
Start-Sleep -Seconds 2

$policyparams = @{
    "name" = $policyname;
    'Bulkspamaction' =  'movetojmf';
    'bulkthreshold' =  '7';
    'highconfidencespamaction' =  'movetojmf';
    'inlinesafetytipsenabled' = $true;
    'markasspambulkmail' = 'on';
    'increasescorewithimagelinks' = 'off'
    'increasescorewithnumericips' = 'on'
    'increasescorewithredirecttootherport' = 'on'
    'increasescorewithbizorinfourls' = 'on';
    'markasspamemptymessages' ='on';
    'markasspamjavascriptinhtml' = 'on';
    'markasspamframesinhtml' = 'on';
    'markasspamobjecttagsinhtml' = 'on';
    'markasspamembedtagsinhtml' ='on';
    'markasspamformtagsinhtml' = 'on';
    'markasspamwebbugsinhtml' = 'on';
    'markasspamsensitivewordlist' = 'on';
    'markasspamspfrecordhardfail' = 'on';
    'markasspamfromaddressauthfail' = 'on';
    'markasspamndrbackscatter' = 'on';
    'phishspamaction' = 'movetojmf';
    'spamaction' = 'movetojmf';
    'zapenabled' = $true
}
Start-Sleep -Seconds 2
new-hostedcontentfilterpolicy @policyparams

write-host -foregroundcolor $colourAlert "Set new filter rule"



$ruleparams = @{
    'name' = $rulename;
    'hostedcontentfilterpolicy' = $policyname;     ## this needs to match the above policy name
    'recipientdomainis' = $domains;
    'Enabled' = $true
    }

New-hostedcontentfilterrule @ruleparams

Start-Sleep -Seconds 2


##  End Spam Policy



Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='
Write-Host '====================================================================================================='

##MALWARE POLICY
$mpolicyname = "BusinessWorks Malware Policy"
$mrulename = "BusinessWorks Malware Rules"


write-host -foregroundcolor $colourAlert "MALWARE POLICY`n"

write-host -foregroundcolor $colourAlert "Set new malware policy"
Start-Sleep -Seconds 2

$mpolicyparams = @{
    "Name" = $mpolicyname;
    'Action' =  'deletemessage';
    'Enablefilefilter' =  $true;
    'Enableinternalsendernotifications' =  $true;
    'Zap' = $true
}
new-malwarefilterpolicy @mpolicyparams

write-host -foregroundcolor $colourAlert "Set new malware filter rule"

$mruleparams = @{
     'name' = $mrulename;
     'comments' = 'This is a custom policy added via a PowerShell script';
     'malwarefilterpolicy' = $mpolicyname; ## this needs to match the above policy name
     'recipientdomainis' = $domains; ## this needs to match the domains you wish to protect in your tenant
     'Priority' = 0;
     'Enabled' = $true
} 

New-malwarefilterrule @ruleparams

write-host -foregroundcolor $colourAlert "Malware Policy Set`n"

Start-Sleep -Seconds 2




## End Script

write-host -foregroundcolor $colourInfo "Tenancy now prepared for the Real World"
read-host -Prompt "Press Enter to Complete"
Exit-PSSession
Exit