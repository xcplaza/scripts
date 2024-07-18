#GET FREE SPACE DISK
clear-host
$currentTime = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
Write-host "Check Disk Yavne - $currentTime" -foregroundcolor yellow
Write-host ""

$domain = "******"
$username = "$domain\******"
$password = ConvertTo-SecureString "******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("******", "******")
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

Foreach ($server in $serversWithDomain) {
    $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $server -credential $creds -Filter DriveType=3 | 
    Select-Object DeviceID, 
    @{'Name' = 'Size'; 'Expression' = { [math]::truncate($_.size / 1GB) } }, 
    @{'Name' = 'Freespace'; 'Expression' = { [math]::truncate($_.freespace / 1GB) } }
            
    Write-host $server -foregroundcolor green
    #$server
    foreach ($disk in $disks) {
        $disks = $disk.DeviceID + $disk.FreeSpace.ToString("N0") + " GB / " + $disk.Size.ToString("N0") + " GB"
        $AVGdisk = 100 - (($disk.Freespace * 100) / $disk.size)
        if ($AVGdisk -gt 90) {
            Write-Host "$disks" -ForegroundColor Red
        }
        else {
            Write-Host "$disks"
        }
    }
}