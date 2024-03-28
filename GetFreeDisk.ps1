#GET FREE SPACE DISK

$servers = @("TIC-AEB-PROD", "TIC-AEB-TEST", "TIC-CA-ROOT", "TIC-CA-SUB", "TIC-CAMERAS", "TIC-DATABANK", "TIC-DC01", "TIC-DC02-NEW", "TIC-DFM", "TIC-HARMONY", "TIC-MASAV", "TIC-PRI", "TIC-PRI-TEST", "TIC-PRINTERS", "TIC-VPN", "TIC-SCCM2016", "TIC-SQL1", "TIC-SQL2", "TIC-VEEAM", "TIC-RDS", "Tic-Duo", "TIC-Cognos", "TIC-CONTROL")
Foreach ($server in $servers)
{
    $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter DriveType=3 | 
        Select-Object DeviceID, 
            @{'Name'='Size'; 'Expression'={[math]::truncate($_.size / 1GB)}}, 
            @{'Name'='Freespace'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}

    $server

    foreach ($disk in $disks)
    {
        $disk.DeviceID + $disk.FreeSpace.ToString("N0") + "GB / " + $disk.Size.ToString("N0") + "GB"

     }
 }