#GET FREE SPACE DISK
clear-host
Write-host "Check Disk..." -foregroundcolor yellow
Write-host ""

$creds = "eilot.org.il\administrator"
$servers = @("192.168.100.5", "192.168.100.7")
Foreach ($server in $servers)
{
    $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter DriveType=3 -credential $creds | 
        Select-Object DeviceID, 
            @{'Name'='Size'; 'Expression'={[math]::truncate($_.size / 1GB)}}, 
            @{'Name'='Freespace'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}
            
Write-host $server -foregroundcolor green
    #$server
    foreach ($disk in $disks)
    {
        $disk.DeviceID + $disk.FreeSpace.ToString("N0") + "GB / " + $disk.Size.ToString("N0") + "GB" + " "

     }
 }