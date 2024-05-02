#ALL BACKUPS VEEAM DISC AND WITHOUT DUBLE
Get-VBRBackupSession | Sort-Object JobName, CreationTime -Descending | 
    Group-Object JobName | 
    ForEach-Object { $_.Group[0] } |
    Select-Object -Unique JobName, JobType, CreationTime, EndTime, Result, State, @{Name="BackupSize";Expression={$_.BackupStats.BackupSize}} | 
    Format-Table