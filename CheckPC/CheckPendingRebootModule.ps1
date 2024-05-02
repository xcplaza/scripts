#CHECKING PENDING REBOOT

$server = Get-ADComputer -Filter * | Where-Object {$_.Name -like "TIC-*"} | Select -Property Name
Test-PendingReboot -ComputerName $server.Name -Detailed -SkipConfigurationManagerClientCheck | select ComputerName, IsRebootPending