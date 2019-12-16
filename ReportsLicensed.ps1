$FilePath = "C:\BWApp\Logs"
write-host -ForegroundColor Yellow "Getting Licensed Users, check the logs for csv file"
write-host -ForegroundColor Yellow "Opening results"
$result = Get-MsolUser -All | where {$_.isLicensed -eq $true}
$result | Out-GridView -Title "Licensed Users - Check the logs folder for a csv export"
$result | Export-Csv "$FilePath\LicensedUsers.csv"

$allusers = Get-MsolUser -All
$allusers | Out-GridView -Title "All Users - Check the logs folder for a csv export"
$allusers | Export-Csv "$FilePath\AllUsers.csv"