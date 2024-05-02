$Servers = Get-ADComputer -Filter * -Properties * | Where-Object {$_.OperatingSystem -like '*Windows*Server*' -or $_.Name -match 'TIC-MASAV'} | Select-Object -ExpandProperty Name

$StartDate = (Get-Date).AddDays(-14)

ForEach ($ServerName in $Servers) {
    try {
        Get-WinEvent -FilterHashtable @{
            LogName = "System"
            ID = @(4740, 4719, 4099, 4688, 4670, 4672, 1125, 1006)
            StartTime = $StartDate
        } -ComputerName $ServerName -ErrorAction Stop | 
        Select-Object @{Name='ServerName'; Expression={$ServerName}}, ID, TimeCreated | 
        Format-Table -AutoSize
    }
    catch [Exception] {
        if ($_.Exception -match "No events were found that match the specified selection criteria") {
            Write-Host "$ServerName - No events found"
        }
    }
}
