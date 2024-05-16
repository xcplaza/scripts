#Command for windows update from Powershell

#Install
Install-Module -Name PSWindowsUpdate

#Import
Import-Module PSWindowsUpdate

#Get all updates
Get-WindowsUpdate

#Install update (KB)
Install-WindowsUpdate -KBArticleID KB5037782

#Hide update (KB)
Hide-WindowsUpdate -KBArticleID KB5007885


