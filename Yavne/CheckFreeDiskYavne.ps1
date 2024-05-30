#GET FREE SPACE DISK
clear-host
Write-host "Check Disk Yavne..." -foregroundcolor yellow
Write-host ""

$domain = "DOMAIN"
$username = "$domain\administrator"
$password = ConvertTo-SecureString "***" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("YavneSQL-Complot", "YavneSQL", "YavneDC1", "YavneDC2", "yavnedc4", "YavneVeeam", "YavneVeeamSV", "DC-365", "YavneFS1", "yavneEX16", "yavctxdc1", "yavctxdc2", "YAVCTXSMS", "YavneApp", "YavnePS1", "yavnedc3", "Biyavne", "yavnesysaid")
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
        if ((($disk.size - $disk.Freespace) / ($disk.Freespace + $disk.size) / 2) * 100 -gt 90) {
            Write-Host "$disks" -ForegroundColor Red
        }
        else {
            Write-Host "$disks"
        }
    }
}