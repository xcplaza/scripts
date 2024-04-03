#GET ALL SERVER IN DOMAIN

Get-ADComputer -Filter * -properties *  | Where-Object {$_.OperatingSystem -like '*Windows*Server*'} | select Name, Enabled,ipv4address, OperatingSystem