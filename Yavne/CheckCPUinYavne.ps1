#GET CPU
clear-host
Write-host "Check CPU Yavne..." -foregroundcolor yellow
Write-host ""

$username = "DOMAIN\administrator"
$password = ConvertTo-SecureString "*******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("10.50.50.157", "10.50.50.13", "10.50.50.7", "10.50.50.8", "10.50.50.4", "10.50.50.25", "10.50.50.105", "10.50.50.16", "10.50.50.11", "10.50.50.17", "10.50.50.15", "10.50.50.19", "10.50.50.12", "10.50.50.18", "10.50.50.32", "10.50.50.60", "10.50.50.24")

foreach ($server in $servers) {
    $serverIP = $server

    # Define the threshold
    $threshold = 90

    # Get the CPU usage data for the last 5 minutes
    $cpuData = Get-WmiObject -ComputerName $serverIP -Credential $creds -Class Win32_PerfFormattedData_PerfOS_Processor |
    Where-Object { $_.Name -eq '_Total' } |
    Select-Object -ExpandProperty PercentProcessorTime |
    Measure-Object -Average

    # Calculate the average CPU usage over the last 5 minutes
    $averageCPU = $cpuData.Average

    # Check if average CPU usage exceeds the threshold
    if ($averageCPU -gt $threshold) {
        Write-Host "$server - $averageCPU%" -ForegroundColor Red
        # You can add additional actions here, like sending an alert or logging the event
    }
    else {
        Write-Host "$server - $averageCPU%" -ForegroundColor Green
    }   
}
