#GET ALL SERVER IN DOMAIN

Get-ADComputer -Filter * -properties * | select Name, Enabled,ipv4address, OperatingSystem