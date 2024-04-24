#CHECK AVARAGE CPU >90 5min

# Define the threshold
$threshold = 90

# Get the current time
$currentTime = Get-Date

# Calculate the time 5 minutes ago
$startTime = $currentTime.AddMinutes(-5)

# Define the counter path for CPU usage
$counterPath = '\Processor(_Total)\% Processor Time'

# Get the CPU usage data for the last 5 minutes
$cpuData = Get-WmiObject Win32_PerfFormattedData_PerfOS_Processor |
           Where-Object { $_.Name -eq '_Total' } |
           Select-Object -ExpandProperty PercentProcessorTime |
           Measure-Object -Average

clear-host
Write-host "Check CPU..." -foregroundcolor yellow
Write-host ""

# Calculate the average CPU usage over the last 5 minutes
$averageCPU = $cpuData.Average

# Check if average CPU usage exceeds the threshold
if ($averageCPU -gt $threshold) {
    Write-Host "Average CPU: $averageCPU%" -ForegroundColor Red
    # You can add additional actions here, like sending an alert or logging the event
} else {
    Write-Host "Average CPU: $averageCPU%" -ForegroundColor Gree
}
