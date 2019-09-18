

$systemmessagecolor = "cyan"
$processmessagecolor = "green"

Clear-Host
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host -ForegroundColor $systemmessagecolor "Enter your Sharepoint url E.g https://busworks-admin.sharepoint.com"
Connect-SPOService -Credential $UserCredential

write-host -foregroundcolor $systemmessagecolor "Enabling Modern Authentication"
write-host
$org=get-organizationconfig
write-host -ForegroundColor $processmessagecolor "Exchange setting is currently",$org.OAuth2ClientProfileEnabled
## Run this command to enable modern authentication for Exchange Online
Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
write-host -foregroundcolor $processmessagecolor "Exchange command completed"
$org=get-organizationconfig
write-host -ForegroundColor $processmessagecolor "Exchange setting updated to",$org.OAuth2ClientProfileEnabled
## To set Sharepoint apps that don’t use modern authentication to block
$org=Get-SPOTenant
write-host -ForegroundColor $processmessagecolor "SharePoint setting is currently",$org.legacyauthprotocolsenabled
Set-spotenant -legacyauthprotocolsenabled $false
$org=Get-SPOTenant
write-host -ForegroundColor $processmessagecolor "SharePoint setting is updated to",$org.legacyauthprotocolsenabled
write-host -foregroundcolor $systemmessagecolor "Script completed`n"
Read-Host -prompt 'Press Enter to return'