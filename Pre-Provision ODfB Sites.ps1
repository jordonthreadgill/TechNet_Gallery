# pre-provision ODfB sites

$users = get-msoluser -All | ? {$_.islicensed -eq $true}
$count = $users.count
$list = @()
$i = 0
$ii = 0

#Option 1
# Loop of users, in "bulk" for allowed user max - 199/200
Function Option-1($optionOne){
    foreach ($u in $users){
        $i++; $ii++; write-host "$i of $count"

        if ($ii -lt 1){
            $upn = $u.userprincipalname
            $list += $upn
        }
        if ($ii -gt 1){
            Request-SPOPersonalSite -UserEmails $list 
            Start-Sleep -Milliseconds 655
            $list = @()
            $ii = 0
        }
    }
    Request-SPOPersonalSite -UserEmails $list -ErrorAction Continue
}
Option-1

#Option 2
# Loops through each end user. [This will take longer]
Function Option-2($optionTwo){
    foreach ($u in $users) {Request-SPOPersonalSite -UserEmails $($u.UserPrincipalName); Start-Sleep -Milliseconds 655}
}
Option-2

# Add additional ODfB Owner for purpose of completing a migration
$admin = Read-Host "Enter the Admin account that is completing the data migration..." 
Get-SPOSite -Limit all -IncludePersonalSite $true -erroraction silentlycontinue | % {Set-SPOUser -site $($_.Url) -LoginName $admin -IsSiteCollectionAdmin:$true -ErrorAction SilentlyContinue; start-sleep -milliseconds 110}
