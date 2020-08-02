

$base = $env:USERPROFILE
$desktop = "$base\Desktop"
$date = (Get-Date).ToString('MM-dd-yyyy')

$domains = Get-MsolDomain 
$default = $domains | ? {$_.isdefault -eq $true} | select -expand Name

$Sites = Get-SPOSite -IncludePersonalSite $true

$personalSites = $Sites | ? {$_.template -eq "SPSPERS#10"}
$usersCount = $personalSites.count
$report = New-Object -TypeName System.Collections.Generic.List[PSCustomObject]
$i = 0
foreach ($u in $personalSites)
{
    $LastContentModifiedDate = $u.LastContentModifiedDate
    $status = $u.Status
    $StorageUsageCurrent = $u.StorageUsageCurrent
    $url = $u.Url
    $owner = $u.Owner
    $StorageQuota = $u.StorageQuota
    $StorageQuotaWarningLevel = $u.StorageQuotaWarningLevel
    $ResourceQuota = $u.ResourceQuota
    $Title = $u.Title
    $SharingCapability = $u.SharingCapability

    $object = [pscustomobject]@{Owner = $owner; Title = $Title; LastContentModifiedDate = $LastContentModifiedDate; Url = $url; Status = $status; `
        StorageUsageCurrent = $StorageUsageCurrent; StorageQuota = $StorageQuota; SharingCapability = $SharingCapability}

    $report.Add($object)
}

$report | Export-Csv -NoTypeInformation -Path "$desktop\$default ODfB Usage Report $date.csv"
