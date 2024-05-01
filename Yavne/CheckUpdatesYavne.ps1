#CHECK UPDATES
clear-host
Write-host "Check update..." -foregroundcolor yellow
Write-host ""

$username = "YAVNED.MUNI\administrator2"
$password = ConvertTo-SecureString "edr!23@4" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList($username, $password)

$servers = @("10.50.50.157")
#$servers = @("10.50.50.157", "10.50.50.13", "10.50.50.7", "10.50.50.8", "10.50.50.4", "10.50.50.25", "10.50.50.105", "10.50.50.16", "10.50.50.11", "10.50.50.17", "10.50.50.15", "10.50.50.19", "10.50.50.12", "10.50.50.18", "10.50.50.32", "10.50.50.60", "10.50.50.24")

foreach ($server in $servers){
    Write-Host "Checking updates for $server"
    try {
        $session = New-PSSession -ComputerName $server -Credential $creds -ErrorAction Stop
        $pendingUpdates = Invoke-Command -Session $session -ScriptBlock {
            (Get-WmiObject -Query "SELECT * FROM Win32_QuickFixEngineering WHERE HotFixID != 'File 1'")
        }
        if ($pendingUpdates) {
            Write-Host "Pending updates found on $($server):"
            $pendingUpdates | Select-Object -Property Description
        } else {
            Write-Host "No pending updates found on $server."
        }
    } catch {
        Write-Host "Failed to establish a session with $($server): $_"
    } finally {
        if ($session) {
            Remove-PSSession $session
        }
    }
}
