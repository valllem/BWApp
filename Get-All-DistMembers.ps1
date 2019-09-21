$FilePath = "C:\BWApp\Logs"
Write-Host "Saving csv file to $Path\All-Distribution-Group-Members.csv "

$Result=@()
$groups = Get-DistributionGroup -ResultSize Unlimited
$totalmbx = $groups.Count
$i = 1
$groups | ForEach-Object {
Write-Progress -activity "Processing $_.DisplayName" -status "$i out of $totalmbx completed"
$group = $_
Get-DistributionGroupMember -Identity $group.Name -ResultSize Unlimited | ForEach-Object {
$member = $_
$Result += New-Object PSObject -property @{
GroupName = $group.DisplayName
Member = $member.Name
EmailAddress = $member.PrimarySMTPAddress
RecipientType= $member.RecipientType
}}
$i++
}
$Result | Export-CSV "$FilePath\All-Distribution-Group-Members.csv" -NoTypeInformation -Encoding UTF8

Write-Host -ForegroundColor Green "Saved to: $FilePath\All-Distribution-Group-Members.csv"
Write-Host -ForegroundColor Green "Opening File..."
Start-Sleep -Seconds 2
Invoke-Item "$FilePath\Permission.csv"