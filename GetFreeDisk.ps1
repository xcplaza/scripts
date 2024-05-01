#GET FREE SPACE DISK
clear-host
Write-host "Check Disk..." -foregroundcolor yellow
Write-host ""

    $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter DriveType=3 -credential $creds | 
        Select-Object DeviceID, 
            @{'Name'='Size'; 'Expression'={[math]::truncate($_.size / 1GB)}}, 
            @{'Name'='Freespace'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}
            
Write-host $server -foregroundcolor green
    #$server
    foreach ($disk in $disks)
    {
        $disks = $disk.DeviceID + $disk.FreeSpace.ToString("N0") + " GB / " + $disk.Size.ToString("N0") + " GB"
        if ((($disk.size - $disk.Freespace)/($disk.Freespace + $disk.size)/2) * 100 -gt 90) {
        Write-Host "$disks" -ForegroundColor Red
    } else {
        Write-Host "$disks"
    }

     }