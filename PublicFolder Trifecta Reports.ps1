
$base = $env:USERPROFILE
$desktop = "$base/Desktop"
$date = (Get-Date).ToString('MM-dd-yyyy')

$defaultDomain = Get-MsolDomain | ? {$_.isDefault -eq $true} | select -expand Name

$pfs = Get-PublicFolder -Recurse
$count = $pfs.count; Write-Host "$count Public Folders found "

$report = @()
$report2 = @()
$report3 = @()

$i = 0
foreach ($pf in $PFS)
{
    $i++; Write-Host "$i of $count"

    $id = $pf.Identity

    $folder = Get-PublicFolder -Identity $id
    $stats = Get-PublicFolderStatistics -Identity $id

    $name = $stats.name
    $path = $folder.parentpath
    $itemcount = $stats.itemcount
    $ownerid = $folder.mailboxownerid
    $contactcount = $stats.contactcount
    $conMailboxname = $folder.contentmailboxname
    $foldersize = $folder.foldersize
    $deleteditems = $stats.deleteditemsize
    $creation = $stats.creationtime

    $oSize = $stats.TotalItemSize
    $splitSize = $oSize.Split("(")[1]
    $bytes = $splitSize -replace (" bytes*","")
    $bytes1 = $bytes.Split(")")[0]
    $bytes2 = $bytes1 -replace (",","") 
    $bytes3 = $bytes2 / 1Kb
    $bytes4 = [math]::Round($bytes3,3)
    $bytes5 = $bytes2 / 1Mb
    $bytes6 = [math]::Round($bytes5,3)
    $bytesKB = $bytes4
    $bytesMB = $bytes6

    $report3 += New-Object psobject -Property @{ FolderName = $name; ParentPath = $path; `
        "TotalItemSize KB" = $bytesKB; "TotalItemSize MB" = $bytesMB; TotalItemCount = $itemcount;`
        MailboxOwnerID = $OwnerID; ContentMailboxName = $conMailboxname; FolderSize = $foldersize; `
        ContactCount = $contactcount; CreationTime = $creation; `
        TotalDeletedItemSize = $deleteditem }

    Start-Sleep -Milliseconds 350

    $itemStats = Get-PublicFolderItemStatistics -Identity $id | Select Subject,LastModificationTime,HasAttachments,ItemType,MessageSize
    if ($stats -ne $null)
    {
        foreach ($s in $itemStats)
        {
            $subject = $s.Subject
            $lastModTime = $s.LastModificationTime
            $attachments = $s.HasAttachments
            $itemType = $s.ItemType
            $messageSize = $s.MessageSize

            $report += New-Object psobject -Property @{Identity = $id; Subject = $subject; `
                LastModificationTime = $lastModTime; HasAttachments = $attachments; `
                ItemType = $itemType; MessageSize = $messageSize}
        }
        $perms = Get-PublicFolderClientPermission -Identity $id | select FolderName,User,AccessRights
        if($perms -ne $null)
        {
            foreach ($p in $perms)
            {
                $folderName = $p.FolderName
                $user = $p.User
                $accessRights = $p.AccessRights 

                $report2 += new-object psobject -property @{Identity = $id; FolderName = $folderName; `
                    User = $user; AccessRights = $accessRights }
            }
        }    
    }

    Get-Variable -Name stats,perms,folder,itemstats -ErrorAction SilentlyContinue | Remove-Variable -Force
}

$report | select Identity,Subject,LastModificationTime,HasAttachments,ItemType,MessageSize | Sort -Property MessageSize | Export-Csv -NoTypeInformation -Path "$desktop\$defaultDomain Public Folder Item Statistics $date.csv"
$report2 | select Identity,FolderName,User,AccessRights | Export-Csv -NoTypeInformation -Path "$desktop\$defaultDomain Public Folder Folder Permissions $date.csv"
$report3 | select FolderName,Parentpath,"TotalItemSize KB","TotalItemSize MB",TotalItemCount,ContentMailboxName,MailboxOwnerID,FolderSize,ContactCount,CreationTime,TotalDeletedItemSize | Export-Csv -NoTypeInformation -Path "$desktop\$defaultDomain Public Folder Statistics $date.csv"

