<#
    .SYNOPSIS
    Export-AADUsers.ps1

    .DESCRIPTION
    Export Azure Active Directory users to CSV file.

    .LINK
    www.alitajran.com/export-azure-ad-users-to-csv-powershell

    .NOTES
    Written by: ALI TAJRAN
    Website:    www.alitajran.com
    LinkedIn:   linkedin.com/in/alitajran

    .CHANGELOG
    V1.10, 06/20/2023 - Initial version
    V1.10, 06/21/2023 - Added license status and MFA status including methods
    V1.20, 06/22/2023 - Added progress bar and last sign in date
    V1.30, 07/24/2023 - Update for Microsoft Graph PowerShell changes
    V1.40, 04/07/2024 - Added domain name
#>

# Connect to Microsoft Graph API
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "AuditLog.Read.All"

# Create variable for the date stamp
$LogDate = Get-Date -f yyyyMMddhhmm

# Define CSV file export location variable
$Csvfile = "C:\temp\AllAADUsers_$LogDate.csv"

# Retrieve users using the Microsoft Graph API with property
$propertyParams = @{
    All            = $true
    # Uncomment below if you have Azure AD P1/P2 to get last sign in date
    # Property = 'SignInActivity'
    ExpandProperty = 'manager'
}

$users = Get-MgBetaUser @propertyParams
$totalUsers = $users.Count

# Initialize progress counter
$progress = 0

# Initialize an array to store user objects
$userObjects = @()

# Loop through all users and collect user objects
foreach ($index in 0..($totalUsers - 1)) {
    $user = $users[$index]

    # Update progress counter
    $progress++

    # Calculate percentage complete
    $percentComplete = ($progress / $totalUsers) * 100

    # Define progress bar parameters
    $progressParams = @{
        Activity        = "Processing Users"
        Status          = "User $($index + 1) of $totalUsers - $($user.userPrincipalName) - $($percentComplete -as [int])% Complete"
        PercentComplete = $percentComplete
    }

    # Display progress bar
    Write-Progress @progressParams

    # Get manager information
    $managerDN = $user.Manager.AdditionalProperties.displayName
    $managerUPN = $user.Manager.AdditionalProperties.userPrincipalName

    # Create an object to store user properties
    $userObject = [PSCustomObject]@{
        "ID"                          = $user.id
        "First name"                  = $user.givenName
        "Last name"                   = $user.surname
        "Display name"                = $user.displayName
        "User principal name"         = $user.userPrincipalName
        "Domain name"                 = $user.userPrincipalName.Split('@')[1]
        "Email address"               = $user.mail
        "Job title"                   = $user.jobTitle
        "Manager display name"        = $managerDN
        "Manager user principal name" = $managerUPN
        "Department"                  = $user.department
        "Company"                     = $user.companyName
        "Office"                      = $user.officeLocation
        "Employee ID"                 = $user.employeeID
        "Mobile"                      = $user.mobilePhone
        "Phone"                       = $user.businessPhones -join ','
        "Street"                      = $user.streetAddress
        "City"                        = $user.city
        "Postal code"                 = $user.postalCode
        "State"                       = $user.state
        "Country"                     = $user.country
        "User type"                   = $user.userType
        "On-Premises sync"            = if ($user.onPremisesSyncEnabled) { "enabled" } else { "disabled" }
        "Account status"              = if ($user.accountEnabled) { "enabled" } else { "disabled" }
        "Account Created on"          = $user.createdDateTime
        # Uncomment below if you have Azure AD P1/P2 to get last succesful sign in date
        # "Last sign in"                 = if ($user.SignInActivity.LastSuccessfulSignInDateTime) { $user.SignInActivity.LastSuccessfulSignInDateTime } else { "No sign in" }
        "Licensed"                    = if ($user.assignedLicenses.Count -gt 0) { "Yes" } else { "No" }
        "MFA status"                  = "-"
        "Email authentication"        = "-"
        "FIDO2 authentication"        = "-"
        "Microsoft Authenticator App" = "-"
        "Password authentication"     = "-"
        "Phone authentication"        = "-"
        "Software Oath"               = "-"
        "Temporary Access Pass"       = "-"
        "Windows Hello for Business"  = "-"
    }

    $MFAData = Get-MgBetaUserAuthenticationMethod -UserId $user.userPrincipalName

    # Check authentication methods for each user
    foreach ($method in $MFAData) {
        Switch ($method.AdditionalProperties["@odata.type"]) {
            "#microsoft.graph.emailAuthenticationMethod" {
                $userObject."Email authentication" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.fido2AuthenticationMethod" {
                $userObject."FIDO2 authentication" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $userObject."Microsoft Authenticator App" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.passwordAuthenticationMethod" {
                $userObject."Password authentication" = $true
                # When only the password is set, then MFA is disabled.
                if ($userObject."MFA status" -ne "Enabled") {
                    $userObject."MFA status" = "Disabled"
                }
            }
            "#microsoft.graph.phoneAuthenticationMethod" {
                $userObject."Phone authentication" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.softwareOathAuthenticationMethod" {
                $userObject."Software Oath" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                $userObject."Temporary Access Pass" = $true
                $userObject."MFA status" = "Enabled"
            }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                $userObject."Windows Hello for Business" = $true
                $userObject."MFA status" = "Enabled"
            }
        }
    }

    # Add user object to the array
    $userObjects += $userObject
}

# Complete the progress bar
Write-Progress -Activity "Processing Users" -Completed

# Export all user objects to CSV
$userObjects | Sort-Object "Display name" | Export-Csv -Path $Csvfile -NoTypeInformation -Encoding UTF8 #-Delimiter ";"