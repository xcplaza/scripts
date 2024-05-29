#REPORT ABOUT FREE DISK

Set-ExecutionPolicy Unrestricted -Force
Import-Module ActiveDirectory

# Delete reports older than 60 days
$OldReports = (Get-Date).AddDays(-60)

# Location for disk reports
Get-ChildItem "C:\Temp\DiskSpaceReport\*.*" |
Where-Object { $_.LastWriteTime -le $OldReports } |
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# Create variable for log date
$LogDate = Get-Date -Format yyyyMMddhhmm

# Get all systems
$Systems = Get-ADComputer -Properties * -Filter { OperatingSystem -like "*Windows Server*" } |
Where-Object { $_.Enabled -eq $true } | Select-Object Name, DNSHostName, OperatingSystem, OperatingSystemVersion | Sort-Object Name

# Loop through each system
$DiskReport = ForEach ($System in $Systems) {
    $OperatingSystem = $System.OperatingSystem
    $OperatingSystemVersion = $System.OperatingSystemVersion
    Get-WmiObject Win32_LogicalDisk `
        -ComputerName $System.DNSHostName -Filter "DriveType=3" `
        -ErrorAction SilentlyContinue |
    Select-Object `
    @{Label = "HostName"; Expression = { $_.SystemName } },
    @{Label = "DriveLetter"; Expression = { $_.DeviceID } },
    @{Label = "DriveName"; Expression = { $_.VolumeName } },
    @{Label = "Total Capacity (GB)"; Expression = { "{0:N1}" -f ($_.Size / 1gb) } },
    @{Label = "Free Space (GB)"; Expression = { "{0:N1}" -f ($_.Freespace / 1gb ) } },
    @{Label = 'Free Space (%)'; Expression = { "{0:P0}" -f ($_.Freespace / $_.Size) } },
    @{Label = "Operating System"; Expression = { $OperatingSystem } },
    @{Label = "Operating System Version"; Expression = { $OperatingSystemVersion } }
}

# Create disk report
$DiskReport |
Export-Csv -Path "C:\Temp\DiskSpaceReport\DiskReport_$LogDate.csv" -NoTypeInformation #-Delimiter ";"