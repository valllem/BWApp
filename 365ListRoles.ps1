Clear-Host
##--Script from: https://gist.github.com/ciphertxt/2036e614edf4bf920796059017fbbc3d--##

Import-Module MSOline -EA 0

Connect-MsolService -Credential (Get-Credential)

$admins=@()

$roles = Get-MsolRole 

foreach ($role in $roles) {
    $roleUsers = Get-MsolRoleMember -RoleObjectId $role.ObjectId

    foreach ($roleUser in $roleUsers) {
        $roleOutput = New-Object -TypeName PSObject
        $roleOutput | Add-Member -MemberType NoteProperty -Name RoleMemberType -Value $roleUser.RoleMemberType
        $roleOutput | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $roleUser.EmailAddress
        $roleOutput | Add-Member -MemberType NoteProperty -Name DisplayName -Value $roleUser.DisplayName
        $roleOutput | Add-Member -MemberType NoteProperty -Name isLicensed -Value $roleUser.isLicensed
        $roleOutput | Add-Member -MemberType NoteProperty -Name RoleName -Value $role.Name

        $admins += $roleOutput
    }
}
Write-Host -ForegroundColor Yellow 'Saving Export'
$admins | Export-Csv -NoTypeInformation ~\Downloads\365RolesUsers.csv
##--End of downloaded script--##
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()