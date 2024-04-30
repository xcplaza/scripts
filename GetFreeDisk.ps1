#GET FREE SPACE DISK
clear-host
Write-host "Check Disk..." -foregroundcolor yellow
Write-host ""

$username = "eilot.org.il\administrator"
$password = ConvertTo-SecureString "imp@ctIT2006" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

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
        $disks = $disk.DeviceID + $disk.FreeSpace.ToString("N0") + " GB / " + $disk.Size.ToString("N0") + " GB"
        if (($disk.Freespace / $disk.size)*100 -gt 90) {
        Write-Host "$disks"
    } else {
        Write-Host "$disks" -ForegroundColor Red
    }

     }
 }