<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    PrepareTenancy
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '344,330'
$Form.text                       = "Prepare Tenancy"
$Form.BackColor                  = "#f5a623"
$Form.TopMost                    = $false

$TextBoxDomain                   = New-Object system.Windows.Forms.TextBox
$TextBoxDomain.multiline         = $false
$TextBoxDomain.width             = 265
$TextBoxDomain.height            = 20
$TextBoxDomain.location          = New-Object System.Drawing.Point(38,92)
$TextBoxDomain.Font              = 'Microsoft Sans Serif,10'

$LabelDomain                     = New-Object system.Windows.Forms.Label
$LabelDomain.text                = "Domain to Prepare (E.g. contoso.com)"
$LabelDomain.AutoSize            = $true
$LabelDomain.width               = 25
$LabelDomain.height              = 10
$LabelDomain.location            = New-Object System.Drawing.Point(38,71)
$LabelDomain.Font                = 'Microsoft Sans Serif,10'

$LabelSharepoint                 = New-Object system.Windows.Forms.Label
$LabelSharepoint.text            = "Tenant Admin"
$LabelSharepoint.AutoSize        = $true
$LabelSharepoint.width           = 25
$LabelSharepoint.height          = 10
$LabelSharepoint.location        = New-Object System.Drawing.Point(38,131)
$LabelSharepoint.Font            = 'Microsoft Sans Serif,10'

$TextBoxTenant                   = New-Object system.Windows.Forms.TextBox
$TextBoxTenant.multiline         = $false
$TextBoxTenant.width             = 266
$TextBoxTenant.height            = 20
$TextBoxTenant.location          = New-Object System.Drawing.Point(38,150)
$TextBoxTenant.Font              = 'Microsoft Sans Serif,10'

$RunScript                       = New-Object system.Windows.Forms.Button
$RunScript.BackColor             = "#7ed321"
$RunScript.text                  = "Run Script"
$RunScript.width                 = 263
$RunScript.height                = 30
$RunScript.location              = New-Object System.Drawing.Point(40,189)
$RunScript.Font                  = 'Microsoft Sans Serif,10'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 328
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(7,234)

$Status                          = New-Object system.Windows.Forms.Label
$Status.text                     = "Status:Ready"
$Status.AutoSize                 = $true
$Status.width                    = 25
$Status.height                   = 10
$Status.location                 = New-Object System.Drawing.Point(7,308)
$Status.Font                     = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($TextBoxDomain,$LabelDomain,$LabelSharepoint,$TextBoxTenant,$RunScript,$ProgressBar1,$Status))

$RunScript.Add_Click({

$ProgressBar1.Value = 1
$Status.Text = "Starting..."

$Domain = $TextBoxDomain.Text
$Tenant = $TextBoxTenant.Text

Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 30
$Status.Text = "Status: Preparing, please wait"
Start-Sleep -Seconds 2
#Enable-OrganizationCustomization
Start-Sleep -Seconds 2
$ProgressBar1.Value = 60

Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 20
$Status.Text = "Status: Enabling Audit Logging"
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true

Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 30
$Status.Text = "Status: Securing Mailboxes"

Start-Sleep -Seconds 2
$ProgressBar1.Value = 35
$Status.Text = "Status: Disabling POP/IMAP"
$mailboxes = Get-mailbox
foreach ($mailbox in $mailboxes) {
    ## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -popenabled $false
    $Status.Text = "Status: Disabling POP for $mailbox"
    write-host -foregroundcolor yellow "Disabled POP for " $mailbox ""
    Start-Sleep -Milliseconds 400
    ## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
    Set-casmailbox $mailbox.PrimarySmtpAddress -imapenabled $false
    $Status.Text = "Status: Disabling IMAP for $mailbox"
    write-host -foregroundcolor yellow "Disabled IMAP for " $mailbox ""
    Start-Sleep -Milliseconds 400
}
## Set Default POP IMAP settings
write-host -ForegroundColor cyan 'Setting default POP / IMAP settings to DISABLED'
$ProgressBar1.Value = 40
$Status.Text = "Status: Setting POP/IMAP default to disabled"
Start-Sleep -Milliseconds 400
Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" } | set-CASMailboxPlan -ImapEnabled $false -PopEnabled $false


Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 50
$Status.Text = "Status: Configure SPAM Policy"
#SPam Policy
$policyname = "BusinessWorks Policy"
$rulename = "BusinessWorks Configured Recpients"
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
Start-Sleep -Milliseconds 400
new-hostedcontentfilterpolicy @policyparams
write-host -foregroundcolor yellow "Set new filter rule"
$ruleparams = @{
    'name' = $rulename;
    'hostedcontentfilterpolicy' = $policyname;     ## this needs to match the above policy name
    'recipientdomainis' = $domain;
    'Enabled' = $true
    }
New-hostedcontentfilterrule @ruleparams
write-host -foregroundcolor green "SPAM Policy Set`n"
Start-Sleep -Milliseconds 400


Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 60
$Status.Text = "Status: Configure Malware Policy"
try{
$mpolicyname = "BusinessWorks Malware Policy"
$mrulename = "BusinessWorks Malware Rules"
write-host -foregroundcolor cyan "MALWARE POLICY`n"
write-host -foregroundcolor cyan "Set new malware policy"
Start-Sleep -Milliseconds 400
$mpolicyparams = @{
    "Name" = $mpolicyname;
    'Action' =  'deletemessage';
    'Enablefilefilter' =  $true;
    'Enableinternalsendernotifications' =  $true;
    'Zap' = $true
}
new-malwarefilterpolicy @mpolicyparams
write-host -foregroundcolor yellow "Set new malware filter rule"
$mruleparams = @{
     'name' = $mrulename;
     'comments' = 'This is a custom policy added via a PowerShell script';
     'malwarefilterpolicy' = $mpolicyname; ## this needs to match the above policy name
     'recipientdomainis' = $domain; ## this needs to match the domains you wish to protect in your tenant
     'Priority' = 0;
     'Enabled' = $true
} 
New-malwarefilterrule @ruleparams
write-host -foregroundcolor green "Malware Policy Set`n"
Start-Sleep -Milliseconds 400
}
catch{
write-host -foregroundcolor Red "Malware Policy NOT Set, Tenant may not have correct license`n"
}


Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 70
$Status.Text = "Status: Transport Rule: Outside Sender Matches Internal User"
Write-Host -ForegroundColor Yellow "Creating Transport rule: Outside Sender Matches Internal User"
$ruleName = "External Senders with matching Display Names"
$ruleHtml = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>This message was sent from outside the company by someone with a display name matching a user in your organisation. Please do not click links or open attachments unless you recognise the source of this email and know the content is safe. <o:p></o:p></span></p></div></td></tr></table>"
$rule = Get-TransportRule | Where-Object {$_.Identity -contains $ruleName}
$displayNames = (Get-Mailbox -ResultSize Unlimited).DisplayName
if (!$rule) {
    Write-Host "Rule not found, creating rule" -ForegroundColor Green
    New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" `
        -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
}
else {
    Write-Host "Rule found, updating rule" -ForegroundColor Green
    Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" `
        -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
}
Write-Host -ForegroundColor Green "Transport Rule Complete"
###########################################################################################################################################
Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 80
$Status.Text = "Status: Preparing to Import Policies - user input required"
##########################################################################


function Get-AuthToken {

[cmdletbinding()]

param
(
    [Parameter(Mandatory=$true)]
    $User
)

$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User

$tenant = $userUpn.Host

Write-Host "Checking for AzureAD module..."

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    if ($AadModule -eq $null) {

        Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable

    }

    if ($AadModule -eq $null) {
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        exit
    }

# Getting path to ActiveDirectory Assemblies
# If the module count is greater than 1 find the latest version

    if($AadModule.count -gt 1){

        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]

        $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }

            # Checking if there are multiple versions of the same module found

            if($AadModule.count -gt 1){

            $aadModule = $AadModule | select -Unique

            }

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

    else {

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"

$redirectUri = "urn:ietf:wg:oauth:2.0:oob"

$resourceAppIdURI = "https://graph.microsoft.com"

$authority = "https://login.microsoftonline.com/$Tenant"

    try {

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"

    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result

        # If the accesstoken is valid then create the authentication header

        if($authResult.AccessToken){

        # Creating header for Authorization token

        $authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }

        return $authHeader

        }

        else {

        Write-Host
        Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
        Write-Host
        break

        }

    }

    catch {

    write-host $_.Exception.Message -f Red
    write-host $_.Exception.ItemName -f Red
    write-host
    break

    }

}

####################################################

Function Test-JSON(){

<#
.SYNOPSIS
This function is used to test if the JSON passed to a REST Post request is valid
.DESCRIPTION
The function tests if the JSON passed to the REST Post is valid
.EXAMPLE
Test-JSON -JSON $JSON
Test if the JSON is valid before calling the Graph REST interface
.NOTES
NAME: Test-JSON
#>

param (

$JSON

)

    try {

    $TestJSON = ConvertFrom-Json $JSON -ErrorAction Stop
    $validJson = $true

    }

    catch {

    $validJson = $false
    $_.Exception

    }

    if (!$validJson){
    
    Write-Host "Provided JSON isn't in valid JSON format" -f Red
    break

    }

}

####################################################

Function Add-DeviceCompliancePolicy(){

<#
.SYNOPSIS
This function is used to add a device compliance policy using the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and adds a device compliance policy
.EXAMPLE
Add-DeviceCompliancePolicy -JSON $JSON
Adds an iOS device compliance policy in Intune
.NOTES
NAME: Add-DeviceCompliancePolicy
#>

[cmdletbinding()]

param
(
    $JSON
)

$graphApiVersion = "Beta"
$Resource = "deviceManagement/deviceCompliancePolicies"
    
    try {

        if($JSON -eq "" -or $JSON -eq $null){

        write-host "No JSON specified, please specify valid JSON for the iOS Policy..." -f Red

        }

        else {

        Test-JSON -JSON $JSON

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $JSON -ContentType "application/json"

        }

    }
    
    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break

    }

}

####################################################

#region Authentication

write-host

# Checking if authToken exists before running authentication
if($global:authToken){

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

        if($TokenExpires -le 0){

        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host

            # Defining User Principal Name if not present

            if($User -eq $null -or $User -eq ""){

            $User = $TextBoxTenant.text
            Write-Host

            }

        $global:authToken = Get-AuthToken -User $User

        }
}

# Authentication doesn't exist, calling Get-AuthToken function

else {

    if($User -eq $null -or $User -eq ""){

    $User = $TextBoxTenant.text
    Write-Host

    }

# Getting the authorization token
$global:authToken = Get-AuthToken -User $User

}

#endregion

####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Compliance\AndroidNoPassRequired.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5

# Adding Scheduled Actions Rule to JSON
$scheduledActionsForRule = '"scheduledActionsForRule":[{"ruleName":"PasswordRequired","scheduledActionConfigurations":[{"actionType":"block","gracePeriodHours":0,"notificationTemplateId":"","notificationMessageCCList":[]}]}]'        

$JSON_Output = $JSON_Output.trimend("}")

$JSON_Output = $JSON_Output.TrimEnd() + "," + "`r`n"

# Joining the JSON together
$JSON_Output = $JSON_Output + $scheduledActionsForRule + "`r`n" + "}"
            
write-host
write-host "Compliance Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Compliance Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceCompliancePolicy -JSON $JSON_Output

###########################################################################################################################################
Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 90
$Status.Text = "Status: Preparing to Import Device Policies"
##########################################################################


function Get-AuthToken {

<#
.SYNOPSIS
This function is used to authenticate with the Graph API REST interface
.DESCRIPTION
The function authenticate with the Graph API Interface with the tenant name
.EXAMPLE
Get-AuthToken
Authenticates you with the Graph API interface
.NOTES
NAME: Get-AuthToken
#>

[cmdletbinding()]

param
(
    [Parameter(Mandatory=$true)]
    $User
)

$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User

$tenant = $userUpn.Host

Write-Host "Checking for AzureAD module..."

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    if ($AadModule -eq $null) {

        Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable

    }

    if ($AadModule -eq $null) {
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        exit
    }

# Getting path to ActiveDirectory Assemblies
# If the module count is greater than 1 find the latest version

    if($AadModule.count -gt 1){

        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]

        $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }

            # Checking if there are multiple versions of the same module found

            if($AadModule.count -gt 1){

            $aadModule = $AadModule | select -Unique

            }

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

    else {

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"

$redirectUri = "urn:ietf:wg:oauth:2.0:oob"

$resourceAppIdURI = "https://graph.microsoft.com"

$authority = "https://login.microsoftonline.com/$Tenant"

    try {

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"

    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result

        # If the accesstoken is valid then create the authentication header

        if($authResult.AccessToken){

        # Creating header for Authorization token

        $authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }

        return $authHeader

        }

        else {

        Write-Host
        Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
        Write-Host
        break

        }

    }

    catch {

    write-host $_.Exception.Message -f Red
    write-host $_.Exception.ItemName -f Red
    write-host
    break

    }

}

####################################################

Function Add-DeviceConfigurationPolicy(){

<#
.SYNOPSIS
This function is used to add an device configuration policy using the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and adds a device configuration policy
.EXAMPLE
Add-DeviceConfigurationPolicy -JSON $JSON
Adds a device configuration policy in Intune
.NOTES
NAME: Add-DeviceConfigurationPolicy
#>

[cmdletbinding()]

param
(
    $JSON
)

$graphApiVersion = "Beta"
$DCP_resource = "deviceManagement/deviceConfigurations"
Write-Verbose "Resource: $DCP_resource"

    try {

        if($JSON -eq "" -or $JSON -eq $null){

        write-host "No JSON specified, please specify valid JSON for the Device Configuration Policy..." -f Red

        }

        else {

        Test-JSON -JSON $JSON

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($DCP_resource)"
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $JSON -ContentType "application/json"

        }

    }
    
    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break

    }

}

####################################################

Function Test-JSON(){

<#
.SYNOPSIS
This function is used to test if the JSON passed to a REST Post request is valid
.DESCRIPTION
The function tests if the JSON passed to the REST Post is valid
.EXAMPLE
Test-JSON -JSON $JSON
Test if the JSON is valid before calling the Graph REST interface
.NOTES
NAME: Test-AuthHeader
#>

param (

$JSON

)

    try {

    $TestJSON = ConvertFrom-Json $JSON -ErrorAction Stop
    $validJson = $true

    }

    catch {

    $validJson = $false
    $_.Exception

    }

    if (!$validJson){
    
    Write-Host "Provided JSON isn't in valid JSON format" -f Red
    break

    }

}

####################################################

#region Authentication

write-host

# Checking if authToken exists before running authentication
if($global:authToken){

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

        if($TokenExpires -le 0){

        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host

            # Defining User Principal Name if not present

            if($User -eq $null -or $User -eq ""){

            $User = $TextBoxTenant.text
            Write-Host

            }

        $global:authToken = Get-AuthToken -User $User

        }
}

# Authentication doesn't exist, calling Get-AuthToken function

else {

    if($User -eq $null -or $User -eq ""){

    $User = $TextBoxTenant.text
    Write-Host

    }

# Getting the authorization token
$global:authToken = Get-AuthToken -User $User

}

#endregion

####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Configuration\DeviceSecurity.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

####################################################

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version,supportsScopeTags

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5
            
write-host
write-host "Device Configuration Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Device Configuration Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceConfigurationPolicy -JSON $JSON_Output


#####################
####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Configuration\DeviceSecurity.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

####################################################

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version,supportsScopeTags

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5
            
write-host
write-host "Device Configuration Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Device Configuration Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceConfigurationPolicy -JSON $JSON_Output


#####################
####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Configuration\DisableWindowsHello.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

####################################################

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version,supportsScopeTags

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5
            
write-host
write-host "Device Configuration Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Device Configuration Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceConfigurationPolicy -JSON $JSON_Output


#####################
####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Configuration\EndpointProtection.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

####################################################

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version,supportsScopeTags

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5
            
write-host
write-host "Device Configuration Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Device Configuration Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceConfigurationPolicy -JSON $JSON_Output




#####################
####################################################

$ImportPath = "C:\BWApp\BWApp-master\Policies\Configuration\Updates.json"

# Replacing quotes for Test-Path
$ImportPath = $ImportPath.replace('"','')

if(!(Test-Path "$ImportPath")){

Write-Host "Import Path for JSON file doesn't exist..." -ForegroundColor Red
Write-Host "Script can't continue..." -ForegroundColor Red
Write-Host
break

}

####################################################

$JSON_Data = gc "$ImportPath"

# Excluding entries that are not required - id,createdDateTime,lastModifiedDateTime,version
$JSON_Convert = $JSON_Data | ConvertFrom-Json | Select-Object -Property * -ExcludeProperty id,createdDateTime,lastModifiedDateTime,version,supportsScopeTags

$DisplayName = $JSON_Convert.displayName

$JSON_Output = $JSON_Convert | ConvertTo-Json -Depth 5
            
write-host
write-host "Device Configuration Policy '$DisplayName' Found..." -ForegroundColor Yellow
write-host
$JSON_Output
write-host
Write-Host "Adding Device Configuration Policy '$DisplayName'" -ForegroundColor Yellow
Add-DeviceConfigurationPolicy -JSON $JSON_Output
Write-Host
Write-Host -ForegroundColor Yellow "Remember to assign the Intune Policies to users or groups"

############################################################################################################################################

Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 100
$Status.Text = "Status: Tenancy Preparation Completed"

Write-Host
Write-Host -ForegroundColor Green "Tenancy Preparation Complete!"

Start-Sleep -Milliseconds 400
$ProgressBar1.Value = 100
$Status.Text = "Status: Tenancy Preparation Completed"


  })

$Form.ShowDialog()