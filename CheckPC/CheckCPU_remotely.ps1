#GET CPU
clear-host
$currentTime = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
Write-host "Check CPU Yavne - $currentTime" -foregroundcolor yellow
Write-host ""

$domain = "******"
$username = "$domain\******"
$password = ConvertTo-SecureString "******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("******", "******")
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

foreach ($server in $serversWithDomain) {
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
