# Mobile Device Report

$base = $env:USERPROFILE
$desktop = "$base/Desktop"
$date = (Get-Date).ToString('MM-dd-yyyy')
$outpath = "$desktop\MobileReport $date.csv"  ##  change this to whatever you want

$mobileReport = @()

$mailboxes = Get-Mailbox -ResultSize unlimited
$mailboxesCount = $mailboxes.count

$licensedMailboxes = $mailboxes | ? { $_.RecipientTypeDetails -eq "UserMailbox" }

foreach ($m in $licensedMailboxes)
{
    $alias = $m.alias
    $upn = $m.PrimarySmtpAddress

    $Devices = Get-MobileDevice -Mailbox $upn | select Identity
    
    $mobileDevices = @()
    foreach ($mobile in $devices)
    {
        $mobileDevices += Get-MobileDevice -Identity $mobile.identity | Select FriendlyName,DeviceID,DeviceOS,DeviceModel,isDisabled,Identity,GUID,WhenChanged
    }

        foreach ($device in $mobileDevices)
        {
            $id = $device.Identity

            $mobileStats = Get-MobileDeviceStatistics -Identity $id | Select DeviceType,LastSyncAttemptTime,LastSuccessSync,isRemoteWipeSupported,LastDeviceWipeRequestor,LastAccountOnlyDeviceWipeRequestor,DeviceAccessState

            $mobileReport += New-Object psobject -Property @{ UserPrincipalName = $upn; DeviceType = $mobileStats.DeviceType; FriendlyName = $device.FriendlyName; `
            DeviceID = $device.DeviceID; DeviceModel = $device.DeviceModel; DeviceOS = $device.DeviceOS; IsDisabled = $device.isDisabled; `
            IsRemoteWipeSupported = $mobileStats.isremotewipesupported; WhenChanged = $device.whenchanged; LastSyncAttemptTime = $mobileStats.lastsyncattempttime; `
            LastSuccessSync = $mobileStats.lastsuccesssync; LastDeviceWipeRequest = $mobileStats.lastdevicewiperequestor; DeviceAccessState = $mobileStats.DeviceAccessState; `
            GUID = $device.GUID }
        }
}

$mobileReport | select UserPrincipalName,DeviceType,FriendlyName,DeviceID,DeviceModel,DeviceOS,IsDisabled,IsRemoteWipeSupported,WhenChanged,LastSyncAttemptTime,LastSuccessSync,LastDeviceWipeRequest,DeviceAccessState,GUID | `
Export-Csv -NoTypeInformation -Path $outpath


