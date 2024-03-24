# Assign the local AD and Azure AD connectors value and remember it's case sensitive
$adConnector  = "eilot.org.il"
$aadConnector = "eilotorgil.onmicrosoft.com - AAD"

# Import AzureAD Sync module
Import-Module ADSync

# Create a new ForceFullPasswordSync configuration parameter object
$c = Get-ADSyncConnector -Name $adConnector

# Update the existing connector with the following new configuration
$p = New-Object Microsoft.IdentityManagement.PowerShell.ObjectModel.ConfigurationParameter "Microsoft.Synchronize.ForceFullPasswordSync", String, ConnectorGlobal, $null, $null, $null
$p.Value = 1
$c.GlobalParameters.Remove($p.Name)
$c.GlobalParameters.Add($p)
$c = Add-ADSyncConnector -Connector $c

# Disable Azure AD Connect
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $false

# Re-enable Azure AD Connect to force a full password synchronization
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $true