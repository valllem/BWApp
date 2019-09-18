Clear-Host

$colourInfo = "green"
$colourAlert = "cyan"
$colourError = "red"




$CASMailbox = Get-CASMailboxPlan

foreach ($CAS in $CASMailbox) {
             
    write-host -foregroundcolor $colourInfo "================================="
    Set-casmailbox $CAS -popenabled $false
    write-host -foregroundcolor $colourInfo "Disabled POP for " $CAS ""  
    Set-casmailbox $CAS -imapenabled $false
    write-host -foregroundcolor $colourInfo "Disabled IMAP for " $mailbox ""
    write-host -foregroundcolor $colourInfo "================================="
}

Read-Host -Prompt 'Press Enter to continue...'










## Information
write-host -ForegroundColor $colourAlert 'First we need some information'
write-host -ForegroundColor $colourAlert '=============================='
write-host
write-host
$domains=read-host -prompt 'Enter Tenant Email DOMAIN for Spam Policy...'
Write-Host -foregroundcolor $colourAlert "Enter your Sharepoint url E.g https://busworks-admin.sharepoint.com"
##Import connection to sharepoint - this will request the admin url E.g https://busworks-admin.sharepoint.com
Connect-SPOService -Credential $UserCredential

## Lock Down each existing mailbox
$mailboxes = Get-mailbox
foreach ($mailbox in $mailboxes) {
             
    ## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -popenabled $false
    write-host -foregroundcolor $colourInfo "Disabled POP for " $mailbox ""

    ## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -imapenabled $false
    write-host -foregroundcolor $colourInfo "================================="
    write-host -foregroundcolor $colourInfo "Disabled IMAP for " $mailbox ""
    write-host -foregroundcolor $colourAlert ""
    write-host -foregroundcolor $colourAlert ""
}


##Modern Auth

Clear-Host
Write-Host
Write-Host
Write-Host
Write-Host


write-host -foregroundcolor $colourAlert "Enabling Modern Authentication"
write-host
$org=get-organizationconfig
write-host -ForegroundColor $colourInfo "Exchange setting is currently",$org.OAuth2ClientProfileEnabled
## Run this command to enable modern authentication for Exchange Online
Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
write-host -foregroundcolor $colourInfo "Exchange command completed"
$org=get-organizationconfig
write-host -ForegroundColor $colourInfo "Exchange setting updated to",$org.OAuth2ClientProfileEnabled
## To set Sharepoint apps that donâ€™t use modern authentication to block
$org=Get-SPOTenant
write-host -ForegroundColor $colourInfo "SharePoint setting is currently",$org.legacyauthprotocolsenabled
Set-spotenant -legacyauthprotocolsenabled $false
$org=Get-SPOTenant
write-host -ForegroundColor $colourInfo "SharePoint setting is updated to",$org.legacyauthprotocolsenabled




## Configure the Spam Policy.

$policyname = "BW Std Policy"
$rulename = "BW Configured Recipients"

write-host -foregroundcolor $colourAlert "Setting new spam policy"


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
write-host -foregroundcolor $colourInfo "Tenancy now prepared for the Real World"
Read-Host -Prompt 'Press any key to return to previous menu...'