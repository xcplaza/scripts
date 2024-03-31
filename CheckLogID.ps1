#CHECK LOGs ID WINDOWS 

$S = Get-ADComputer -Filter * -properties * | select Name
$sDate = (get-date).AddDays(-14)
ForEach ($Server in $S) {
try {Get-WinEvent -FilterHashtable @{logname="System"; id=@(4740, 4719, 4099, 4688, 4670, 4672, 1125, 1006); StartTime=$sDate} -ComputerName $Server -ErrorAction Stop| 
      Select-Object LogMode, MaximumSizeInBytes, RecordCount, LogName,
          @{name='ComputerName'; expression={$Server}} |
  Format-Table -AutoSize}
    catch [Exception] {
if ($_.Exception -match "No events were found that match the specified selection criteria") {Write-Host "No events found";}
    }
 }