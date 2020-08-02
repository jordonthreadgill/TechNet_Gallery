# AutoComplete
$base = $env:USERPROFILE
$desktop = "$base/Desktop"
$date = (Get-Date).ToString('MM-dd-yyyy')
$nk2 = "$base\AppData\Roaming\Microsoft\Outlook"
$ac = "$base\AppData\Local\Microsoft\Outlook\RoamCache"
$FolderName = "$desktop\AutoComplete_Data_Copy $date"

$NewFolder = New-Item -Type Directory -Path $FolderName

Write-Host "Please make sure you have completed the steps of generating a new AutoComplete file.

1.  Create a new Outlook profile
2.  Open Outlook with that new profile and send at least 1 message out
3.  Close and reopen Outlook with that profile
4.  Count to 5 and exit Outlook

"

Read-Host "Press Enter to continue "

$acs = Get-ChildItem -Path $ac | ? {$_.Name -like "*Stream_AutoComplete*"} 
$nk2s = Get-ChildItem -Path $nk2 | ? {$_.FullName -like "*.nk2*"}

if ($acs.Count -lt 2)
{
    Read-Host "There is only 1 AutoStream file found.  The new AutoComplete file has not been generated yet.  Please try sending a test message from the new Outlook profile, close and reopen Outlook, and then retry. "
    Break
} 

$acs | select Name,CreationTime,LastWriteTime,LastAccessTime,Length

foreach ($file in $acs)
{
    $fn = $file.FullName

    Copy-Item -Path $fn -Destination $FolderName
}

# Find the newest file and get the name
$newest = $acs | sort -Property LastWriteTime -Descending | select -First 1

# Find largest file
$largest = $acs | sort -Property Length -Descending | select -First 1

# is the script seeing the same file?
if ($newest.Name -eq $largest.Name)
{
    $con = Read-Host = "The largest AutoStream file matches the newest AutoStream file.
    It is likely that the new AutStream file is not generated yet.
    
    Do you wish to continue?  Y or N  "

    if ($con -like "n")
    {
        Exit
    }
}

$newName = "$ac\"+$newest.Name
Copy-Item -Path $($largest.FullName) -Destination $newName -Force 


