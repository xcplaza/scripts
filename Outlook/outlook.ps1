=========== Exchange online ============
Get-ExecutionPolicy -List | Format-Table -AutoSize
Install-Module MSOnline -Force
Install-Module AzureAD -Force
Install-Module AzureADPreview -Force

#install module for to connect to Microsoft Graph PowerShell
Install-Module Microsoft.Graph -Force
Install-Module Microsoft.Graph.Beta -AllowClobber -Force

#connect to MsolUser
Connect-MsolService

#recive deleted mailbox
Get-MsolUser -ReturnDeletedUsers

#recive ImmutableID
Get-MsolUser -UserPrincipalName MoetzaArchive@eilot.org.il | FL immutableId
Get-MsolUser -ReturnDeletedUsers | FL UserPrincipalName,immutableID

#connect to O365
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline
Connect-ExchangeOnline -UserPrincipalName impact@gfnadlan.com

#close connection
[Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()

#get message
Get-MessageTrace -RecipientAddress "eyal@gfnadlan.com" -StartDate "2024-06-16" -EndDate "2024-06-18" | Format-Table -Property Received, SenderAddress, Subject


Test-NetConnection -Port 444 -ComputerName 185.145.254.249