# Office 365 DNS configuration using Azure AD module v2

$base = $env:USERPROFILE
$desktop = "$base/Desktop"

$cred= Get-Credential
Connect-AzureAD -Credential $cred

$domains = Get-AzureADDomain
$default = $domains | ? {$_.IsDefault -eq $true}
$fileName = "$desktop\$default DNS Records.csv"

$DNS = @()
		
foreach ($d in $domains)
{
	$domain = $d.Name
	Set-AzureADDomain -Name $domain -SupportedServices Email, OfficeCommunicationsOnline, OrgIdAuthentication, Intune
	Start-Sleep -Seconds 2
			
	$records = Get-AzureADDomainServiceConfigurationRecord -Name $domain | Select RecordType, Ttl, IsOptional, Port, Priority, Protocol, MailExchange, Service, Weight, Preference, Text, SupportedService, @{ Name = "Host"; Expression = { $_.Label } }, @{ Name = "Target"; Expression = { $_.CanonicalName } }, @{ Name = "Domain"; Expression = { $domain } } | select Domain, IsOptional, RecordType, Host, Ttl, Target, MailExchange,Preference, Text, Priority, Protocol, Service, Weight, Port, SupportedService

	foreach ($n in $records)
	{
		$domain = $n.domain
		if ($n.RecordType -eq "Mx")
		{
			$DNS += New-Object psobject -Property @{Domain = $Domain; IsOptional = $($n.IsOptional); `
				RecordType = $($n.RecordType); Ttl = $($n.Ttl); Host = "@"; Target = $($n.MailExchange); `
				Priority = $($n.Preference); SupportedService = $($n.SupportedService)}
		}
				
		if ($n.RecordType -eq "Txt")
		{
			$DNS += New-Object psobject -Property @{Domain = $Domain; IsOptional = $($n.IsOptional); `
				RecordType = $($n.RecordType); Ttl = $($n.Ttl); Host = "@"; Target = $($n.Text); `
				SupportedService = $($n.SupportedService)}
					
		}
				
		if ($n.RecordType -eq "CName")
		{
			$HostName = $n.Host -replace (".$domain", "")
			$Host1 = $HostName
					
			$DNS += New-Object psobject -Property @{Domain = $Domain; IsOptional = $($n.IsOptional); `
				RecordType = $($n.RecordType); Ttl = $($n.Ttl); Host = $Host1; Target = $($n.Target); `
				SupportedService = $($n.SupportedService)}
		}
				
		if ($n.RecordType -eq "Srv")
		{
			if ($n.Service -like "*_sip*")
			{
				$sip = "sipdir.online.lync.com"
			}
			if ($n.Service -like "*_sipfederationtls*")
			{
				$sip = "sipfed.online.lync.com"
			}
				
			$DNS += New-Object psobject -Property @{Domain = $Domain; IsOptional = $($n.IsOptional); Host = $($n.Service);`
				RecordType = $($n.RecordType); Ttl = $($n.Ttl); Priority = $($n.Priority); Protocol = $($n.Protocol); `
				Weight = $($n.Weight); Port = $($n.Port); Target = $sip; SupportedService = $($n.SupportedService)}
		}
	}
}
		
$DNS | Select Domain, RecordType, IsOptional, Ttl, Host, Target, Priority, Protocol, Weight, Port, SupportedService | Export-Csv -NoTypeInformation -Path $fileName

Disconnect-AzureAD		

