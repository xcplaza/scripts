#GET CPU
clear-host
Write-host "Check CPU Yavne..." -foregroundcolor yellow
Write-host ""

$domain = "DOMAIN"
$username = "$domain\administrator"
$password = ConvertTo-SecureString "***" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("YavneSQL-Complot", "YavneSQL", "YavneDC1", "YavneDC2", "yavnedc4", "YavneVeeam", "YavneVeeamSV", "DC-365", "YavneFS1", "yavneEX16", "yavctxdc1", "yavctxdc2", "YAVCTXSMS", "YavneApp", "YavnePS1", "yavnedc3", "Biyavne", "yavnesysaid")
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
