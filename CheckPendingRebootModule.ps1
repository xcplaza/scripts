#CHECKING PENDING REBOOT

$S = Get-ADComputer -Filter * -properties * | select Name
ForEach ($Server in $S) {
Test-PendingReboot -ComputerName $Server -Detailed -SkipConfigurationManagerClientCheck | select ComputerName, IsRebootPending
}