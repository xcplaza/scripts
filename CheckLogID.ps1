#CHECK LOGs ID WINDOWS 

$S = Get-ADComputer -Filter * -properties *  | Where-Object {$_.OperatingSystem -like '*Windows*Server*' -or $_.name -match 'TIC-MASAV'} | select Name

$sDate = (get-date).AddDays(-1)
ForEach ($Server in $S) {
try {Get-WinEvent -FilterHashtable @{logname="System"; id=@(4740, 4719, 4099, 4688, 4670, 4672, 1125, 1006); StartTime=$sDate} -ComputerName $Server.Name -ErrorAction Stop | 
      Select-Object LogMode, MaximumSizeInBytes, RecordCount, LogName,
          @{name='ComputerName'; expression={$Server}} | Select ComputerName |  Format-Table -AutoSize}
    catch [Exception] {
if ($_.Exception -match "No events were found that match the specified selection criteria") {Write-Host $Server.Name "- No events found";}
    }
 }