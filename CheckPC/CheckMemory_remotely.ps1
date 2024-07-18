#GET MEMORY
clear-host
$currentTime = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
Write-host "Check Memory Yavne - $currentTime" -foregroundcolor yellow
Write-host ""

$domain = "******"
$username = "$domain\******"
$password = ConvertTo-SecureString "******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("******", "******")
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

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