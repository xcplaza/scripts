=========== Exchange online ============
Get-ExecutionPolicy -List | Format-Table -AutoSize
Install-Module MSOnline -Force
Install-Module AzureAD -Force
Install-Module AzureADPreview -Force

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
Connect-ExchangeOnline -UserPrincipalName hsaidian@keystonedental.com

#close connection
[Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()

#get message
Get-MessageTrace -RecipientAddress "pds@keystonedental.com" -StartDate "2024-03-19" -EndDate "2024-03-24" | Format-Table -Property Received, SenderAddress, Subject


Test-NetConnection -Port 444 -ComputerName 185.145.254.249