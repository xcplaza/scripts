clear-host
$smbMappings = Get-SmbMapping | ForEach-Object { $_.LocalPath -replace '/\.*' }

foreach ($mapping in $smbMappings) {
    $disks = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$mapping'"
    foreach ($disk in $disks) {
        $total = $disk.Size / 1GB
        $total = [math]::Round($total, 2)
        $free = $disk.FreeSpace / 1GB
        $free = [math]::Round($free, 2)
        $used = $total - $free
        $usedP = ($used / $total) * 100
        $usedP = [math]::Round($usedP)
        $freeP = ($free / $total) * 100
        $freeP = [math]::Round($freeP)

        Write-Host "Share: $($disk.DeviceID)" -ForegroundColor Green
        Write-Host "Total Disk Space: $total GB"
        Write-Host "--------------------------------------"
        
        if (($used / $total) * 100 -gt 90) {
            #Write-Host "Warning: Disk usage is over 90%." -ForegroundColor Red
            Write-Host "Free Disk Space: $free GB - $freeP %"
            Write-Host "Used Disk Space: $used GB - $usedP %" -ForegroundColor Red
        } else {
        }
    }
    Write-Host ""
}
