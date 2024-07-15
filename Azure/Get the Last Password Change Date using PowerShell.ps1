# 1 - To get the last password change date for a particular user, use this Microsoft Graph PowerShell script:
#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
#Get the User
$User = Get-MgUser -UserId "salaudeen@Crescent.com" -Property UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime
#Get the user's last password change date and time
$User | Select UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime

# 2 - Similarly, to get the last password change date timestamp of all users, use the following PowerShell script:
#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
#Retrieve the password change date timestamp of all users
Get-MgUser -All -Property UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime | Select -Property UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime

# 2a - with file CSV
#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
#Set the properties to retrieve
$Properties = @(
    "id",
    "DisplayName",
    "userprincipalname",
    "PasswordPolicies",
    "lastPasswordChangeDateTime",
    "mail",
    "jobtitle",
    "department"
    )
#Retrieve the password change date timestamp of all users
$AllUsers = Get-MgUser -All -Property $Properties | Select -Property $Properties
#Export to CSV
$AllUsers | Export-Csv -Path "C:\Temp\PasswordChangeTimeStamp.csv" -NoTypeInformation



# 3 - Step-by-step guide for using PowerShell to get the last password change date in Office 365
#Connect to Microsoft Online Service
Connect-MsolService
#Get the User
$user=Get-MsolUser -UserPrincipalName "salaudeen@Crescent.com"
#Get the last password change date
$user.LastPasswordChangeTimestamp

# 4 - How to check when the Office 365 password expires?
#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
#Get the User
$User = Get-MgUser -UserId "Salaudeen@crescent.com" -Property UserPrincipalName, lastPasswordChangeDateTime
#Get the user's password expiring date
$User.lastPasswordChangeDateTime.AddDays(90)

# 5 - Similarly, with MSOL module, you can obtain the password expiring date as:
#Connect to Microsoft Online Service
Connect-MsolService
#Get the User
$user = Get-MsolUser -UserPrincipalName "salaudeen@crescent.com"
#Get the password Expiring date
$user.LastPasswordChangeTimestamp.AddDays(90)

# 6 - This script calculates the password expiration date based on the user’s last password change and the default 90-day password expiration policy. How about finding the password expiry date for all users in your Office 365 tenant?
#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
#Set the properties to retrieve
$Properties = @(
    "id",
    "DisplayName",
    "userprincipalname",
    "PasswordPolicies",
    "lastPasswordChangeDateTime",
    "AccountEnabled",
    "userType"
    )
#Retrieve All users from Microsoft 365 and Filter
$AllUsers = Get-MgUser -All -Property $Properties | Select -Property $Properties | `
    Where {$_.AccountEnabled -eq $true -and $_.PasswordPolicies -notcontains "DisablePasswordExpiration" -and $_.userType -ne "Guest"} 
#Filter and Export data to CSV
$FilteredUsers = $AllUsers | Select Id, DisplayName, UserPrincipalName, @{Name="ExpiryDate";Expression={$_.lastPasswordChangeDateTime.AddDays(90)}} 
$FilteredUsers
$FilteredUsers | Export-Csv -Path "C:\Temp\PasswordExpiryDate.csv" -NoTypeInformation
Write-host "Password Expiry Date for all users is exported!" -f Green

# 7 - To obtain the password expiry date for a particular user account, use this PowerShell script:
#Parameter
$UserAccount = "Salaudeen@Crescent.com"
#Connect to Office 365 from PowerShell
Connect-MsolService
#Get the Default Domain
$Domain = Get-MsolDomain | where {$_.IsDefault -eq $true}
#Get the Password Policy 
$PasswordPolicy = Get-MsolPasswordPolicy -DomainName $Domain.Name
#Get the User account
$UserAccount = Get-MsolUser -UserPrincipalName $UserAccount
#Get Password Expiry date 
$PasswordExpirationDate = $UserAccount.LastPasswordChangeTimestamp.AddDays($PasswordPolicy.ValidityPeriod)
$PasswordExpirationDate
#Get the Password Expiring date
$UserAccount | Select LastPasswordChangeTimestamp, @{Name="Password Age";Expression={((Get-Date).ToUniversalTime())-$_.LastPasswordChangeTimeStamp}}

# 8 - To get the password expiration date for all users using the MSOL module, use this script:
#Connect to Microsoft Online Service
Connect-MsolService
#Get all Users
$AllUsers = Get-MsolUser -All | Select ObjectId, DisplayName,UserPrincipalName, BlockCredential, UserType, @{Name="ExpiryDate";Expression={$_.LastPasswordChangeTimeStamp.AddDays(90)}} 
#Filter users
$FilteredUsers = $AllUsers | Where {$_.PasswordNeverExpires -ne $true -and $_.UserType -ne "Guest" -and $_.BlockCredential -ne $true } | Select DisplayName,UserPrincipalName, ExpiryDate
$FilteredUsers
#Get the password Expiring date
$FilteredUsers | Export-Csv -Path "C:\Temp\PasswordExpiryDates.csv" -NoTypeInformation