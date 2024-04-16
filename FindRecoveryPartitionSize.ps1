#Finding the Recovery Partition Size

$disksObject = @()
Get-WmiObject Win32_Volume -Filter "DriveType='3'" | ForEach-Object {
    $VolObj = $_
    $ParObj = Get-Partition | Where-Object { $_.AccessPaths -contains $VolObj.DeviceID }
    if ( $ParObj ) {
        $disksobject += [pscustomobject][ordered]@{
            DiskID = $([string]$($ParObj.DiskNumber) + "-" + [string]$($ParObj.PartitionNumber)) -as [string]
            Mountpoint = $VolObj.Name
            Letter = $VolObj.DriveLetter
            Label = $VolObj.Label
            FileSystem = $VolObj.FileSystem
            'Capacity(mB)' = ([Math]::Round(($VolObj.Capacity / 1MB),2))
            'FreeSpace(mB)' = ([Math]::Round(($VolObj.FreeSpace / 1MB),2))
            'Free(%)' = ([Math]::Round(((($VolObj.FreeSpace / 1MB)/($VolObj.Capacity / 1MB)) * 100),0))
        }
    }
}
$disksObject | Sort-Object DiskID | Format-Table -AutoSize