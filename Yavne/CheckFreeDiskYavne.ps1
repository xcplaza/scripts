#GET FREE SPACE DISK
clear-host
Write-host "Check Disk Yavne..." -foregroundcolor yellow
Write-host ""

$username = "DOMAIN\administrator"
$password = ConvertTo-SecureString "*******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("10.50.50.157", "10.50.50.13", "10.50.50.7", "10.50.50.8", "10.50.50.4", "10.50.50.25", "10.50.50.105", "10.50.50.16", "10.50.50.11", "10.50.50.17", "10.50.50.15", "10.50.50.19", "10.50.50.12", "10.50.50.18", "10.50.50.32", "10.50.50.60", "10.50.50.24")
Foreach ($server in $servers) {
    $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $server -credential $creds -Filter DriveType=3 | 
    Select-Object DeviceID, 
    @{'Name' = 'Size'; 'Expression' = { [math]::truncate($_.size / 1GB) } }, 
    @{'Name' = 'Freespace'; 'Expression' = { [math]::truncate($_.freespace / 1GB) } }
            
    Write-host $server -foregroundcolor green
    #$server
    foreach ($disk in $disks) {
        $disks = $disk.DeviceID + $disk.FreeSpace.ToString("N0") + " GB / " + $disk.Size.ToString("N0") + " GB"
        if ((($disk.size - $disk.Freespace) / ($disk.Freespace + $disk.size) / 2) * 100 -gt 90) {
            Write-Host "$disks" -ForegroundColor Red
        }
        else {
            Write-Host "$disks"
        }
    }
}