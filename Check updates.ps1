#CHECK UPDATES

#$servers = Get-ADComputer -Filter {(Enabled -eq "True") -and (OperatingSystem -like "*Windows*Server*")} -Properties * | Sort Name | select -Unique Name, Enabled,ipv4address, OperatingSystem
$servers = Get-ADComputer -Filter * -properties * | select Name, Enabled,ipv4address, OperatingSystem

foreach ($server in $servers){
write-host $server.Name

  Invoke-Command -ComputerName $server.Name -ScriptBlock{
#(New-Object -com "Microsoft.Update.AutoUpdate").Results.LastInstallationSuccessDate}
(Get-Hotfix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1).InstalledOn}
}