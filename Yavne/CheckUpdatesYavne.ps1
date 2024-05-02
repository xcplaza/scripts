#CHECK UPDATES
clear-host
Write-host "Check update..." -foregroundcolor yellow
Write-host ""

$username = "YAVNED.MUNI\administrator2"
$password = ConvertTo-SecureString "edr!23@4" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)
#$servers = @("10.50.50.157", "10.50.50.13", "10.50.50.7", "10.50.50.8", "10.50.50.4", "10.50.50.25", "10.50.50.105", "10.50.50.16", "10.50.50.11", "10.50.50.17", "10.50.50.15", "10.50.50.19", "10.50.50.12", "10.50.50.18", "10.50.50.32", "10.50.50.60", "10.50.50.24")

# Add the IP address of the remote server to the list of trusted hosts
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "10.50.50.157" -Force

# Retry the script
$servers = @("10.50.50.157")

foreach ($server in $servers) {
    Write-Host "Checking updates for $server"
    try {
        # Invoke commands directly on the remote computer
        Invoke-Command -ComputerName $server -Credential $creds -ScriptBlock {
            # Create a new COM object for Windows Update Agent
            $updateSession = New-Object -ComObject Microsoft.Update.Session

            # Create a new search object
            $updateSearcher = $updateSession.CreateUpdateSearcher()

            # Search for pending updates
            $pendingUpdates = $updateSearcher.Search("IsInstalled=0")

            if ($pendingUpdates.Updates.Count -gt 0) {
                Write-Host "Pending updates found on $($env:COMPUTERNAME):"
                foreach ($update in $pendingUpdates.Updates) {
                    Write-Host "Title: $($update.Title)"
                    Write-Host "Description: $($update.Description)"
                    Write-Host "----------------------"
                }
            }
            else {
                Write-Host "No pending updates found on $($env:COMPUTERNAME)."
            }
        }
    }
    catch {
        Write-Host "Failed to establish a session with $($server): $_"
    }
}
