#Find the PowerCLI module in the PowerShell Gallery repositories:
Find-Module -Name VMware.PowerCLI

#install PowerCLI module for all users, run the command:
Install-Module -Name VMware.PowerCLI

#The command to install PowerCLI only for the current user and without administrative privileges:
Install-Module -Name VMware.PowerCLI -Scope CurrentUser

#Check the PowerCLI version after finishing installation:
Get-PowerCLIVersion

#If you see Power CLI Error: Invalid server certificate.
#Use Set-PowerCLIConfiguration to set the value for the InvalidCertificateAction option
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

#Connection to VIServer VMware
Connect-VIServer 10.50.50.50
#or
Connect-VIServer -Server 10.23.112.235 -Protocol https -User admin -Password pass

#list VMs with a connected device
Get-VM | Get-CDDrive | Where {$_.extensiondata.connectable.connected -eq $true} | Select Parent

#Run this command to remove and disconnect an attached CD-ROM/DVD devices:
Get-VM | Get-CDDrive | Where {$_.extensiondata.connectable.connected -eq $true} | Set-CDDrive -NoMedia -confirm:$false

#Disable ISO on VMs and 
$vm = Get-VM -Name "Qradar - EC"
New-AdvancedSetting -Entity $vm -Name cdrom.showIsoLockWarning -Value False -Confirm:$false -Force:$true
New-AdvancedSetting -Entity $vm -Name msg.autoanswer -Value TRUE -Confirm:$false -Force:$true