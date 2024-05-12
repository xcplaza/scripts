#GET MEMORY
clear-host
Write-host "Check Memory Yavne..." -foregroundcolor yellow
Write-host ""

$username = "DOMAIN\administrator"
$password = ConvertTo-SecureString "*******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("10.50.50.157", "10.50.50.13", "10.50.50.7", "10.50.50.8", "10.50.50.4", "10.50.50.25", "10.50.50.105", "10.50.50.16", "10.50.50.11", "10.50.50.17", "10.50.50.15", "10.50.50.19", "10.50.50.12", "10.50.50.18", "10.50.50.32", "10.50.50.60", "10.50.50.24")

foreach ($server in $servers) {
    try {
        $memoryUsage = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -Credential $creds |
        ForEach-Object { [math]::Round(($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) / $_.TotalVisibleMemorySize * 100, 2) }
        
        if ($memoryUsage -gt 85) {
            Write-Host "$($server) - $($memoryUsage)%" -ForegroundColor Red
        }
        else {
            Write-Host "$($server) - $($memoryUsage)%" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error retrieving memory usage for $($server): $_" -ForegroundColor Yellow
    }
}