# Bulk Add or Correct Aliases

$base = $env:USERPROFILE
$desktop = "$base/Desktop"

Function Add-Alias {

$i = 0

foreach ($u in $users)
{
    $i++; Write-Host "$i of $count"

    $upn = $u.UserPrincipalName
    $an = $upn.SPlit("@")[0]
    $ud = $upn.SPlit("@")[1]
    
    $domains1 = $domains | ? {$_ -notlike "*$ud*"}

    foreach ($d in $domains1.Name)
    {
        
        $alias = "smtp:" + $an + "@" + $d; Write-Host $alias

        Set-Mailbox -Identity $upn -EmailAddresses @{add=$alias}
    }

    Start-Sleep -Milliseconds 480
    
}
}

Function Correct-Alias {
$domains1 = $domains | ? {$_.IsInitial -eq $false}

$initial = $domains | ? {$_.IsInitial -eq $true} | select -expand Name

$i = 0

foreach ($u in $users)
{
    $i++; Write-Host "$i of $count"

    $upn = $u.UserPrincipalName
    $an = $upn.SPlit("@")[0]
    $ud = $upn.SPlit("@")[1]
    $sys = "smtp:" + $an+"@"+$initial

    $domains2 = $domains1 | ? {$_ -notlike "*$ud*"}

    Set-Mailbox -Identity $upn -EmailAddresses $upn

    Start-Sleep -Milliseconds 480

    foreach ($d in $domains2.Name)
    {
        
        $alias = "smtp:" + $an + "@" + $d; Write-Host $alias

        Set-Mailbox -Identity $upn -EmailAddresses @{add=$alias}
    }

    Set-Mailbox -Identity $upn -EmailAddresses @{add=$sys}
    
}
}

Function The-Ask {
$TheAsk = Read-Host "Do you want to ADD new aliases to existing accounts, or DELETE all aliases and re-add them?  Type ADD or DEL  "

if ($TheAsk -notlike "del" -or $TheAsk -notlike "add")
{
    The-Ask
}

}


$domains = Get-MsolDomain 
$default = $domains | ? {$_.IsDefault -eq $true} | select -expand Name
$users = Get-MsolUser -All | ? {$_.IsLicensed -eq $true}
$count = $users.count

$report = @()
foreach ($u in $users)
{
	$upn = $u.UserPrincipalName
	$aliases = Get-Mailbox -Identity $upn | select -expand EmailAddresses
	$aliases = [string]$aliases
	$aliases = $aliases -replace (" ",";")

	$report += New-Object psobject -Property @{UserPrincipalName = $upn; Aliases = $aliases} 
}
$report | select UserPrincipalName,Aliases | Export-Csv -NoTypeInformation -Path "$desktop/$default User Aliases Report Before.csv"

if ($TheAsk -like "add")
{
    Add-Alias
}

if ($TheAsk -like "del")
{
    Correct-Alias
}

if ($TheAsk -notlike "del" -or $TheAsk -notlike "add")
{
    The-Ask
}

$report = @()
foreach ($u in $users)
{
	$upn = $u.UserPrincipalName
	$aliases = Get-Mailbox -Identity $upn | select -expand EmailAddresses
	$aliases = [string]$aliases
	$aliases = $aliases -replace (" ",";")

	$report += New-Object psobject -Property @{UserPrincipalName = $upn; Aliases = $aliases} 
}
$report | select UserPrincipalName,Aliases | Export-Csv -NoTypeInformation -Path "$desktop/$default User Aliases Report After.csv"
