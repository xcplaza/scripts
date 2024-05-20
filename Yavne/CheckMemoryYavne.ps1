#GET MEMORY
clear-host
Write-host "Check Memory Yavne..." -foregroundcolor yellow
Write-host ""

$username = "domain\administrator"
$password = ConvertTo-SecureString "***" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$domain = ".yavned.muni"
$servers = @("YavneSQL-Complot", "YavneSQL", "YavneDC1", "YavneDC2", "yavnedc4", "YavneVeeam", "DC-365", "YavneFS1", "yavneEX16", "yavctxdc1", "yavctxdc2", "YAVCTXSMS", "YavneApp", "YavnePS1", "yavnedc3", "Biyavne", "yavnesysaid")
$serversWithDomain = $servers | ForEach-Object { "$_$domain" }

foreach ($server in $serversWithDomain) {
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