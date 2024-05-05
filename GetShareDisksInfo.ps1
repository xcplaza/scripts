#GET SHARE DISKS

param ( [Parameter(Mandatory=$true)]
    [string]$share)

$drive = (New-Object -com scripting.filesystemobject).getdrive("$share")
$free = ($drive.FreeSpace / 1GB)
    $free = [math]::Round($free,2)
$total = ($drive.TotalSize / 1GB)
    $total = [math]::Round($total,2)
$used = ($total - $free)
    $used = [math]::Round($used,2)
$usedP = ($used / $total)*100
    $usedP = [math]::Round($usedP)
$freeP = ($free / $total)*100
    $freeP = [math]::Round($freeP)

clear-host
Write-Host "Total Disk Space: $total GB"
#Write-Host "Statistic.totalGB: $total"
Write-host ""
Write-Host "Free Disk Space: $free GB - $freeP %"
#Write-Host "Statistic.freeGB: $free"
Write-Host "Used Disk Space: $used GB - $usedP %"
#Write-Host "Statistic.usedGB: $used"
#Write-host ""
#Write-Host "Message.freeP: Free Disk Space: $freeP %"
#Write-Host "Statistic.freeP: $freeP"
#Write-host ""
#Write-Host "Message.usedP: Used Disk Space: $usedP %"
#Write-Host "Statistic.usedP: $usedP"
Write-host ""