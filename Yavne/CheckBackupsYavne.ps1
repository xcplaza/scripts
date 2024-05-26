#ALL BACKUPS VEEAM DISC AND WITHOUT DUBLE

# Очищаем экран
Clear-Host
Write-Host "Check Backups..." -ForegroundColor Yellow
Write-Host ""

# Указываем домен и учетные данные
$domain = "DOMAIN"
$username = "$domain\administrator"
$password = ConvertTo-SecureString "***" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

# Список серверов
$servers = "YavneVeeamSV"
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

# Скрипт для выполнения на удаленном сервере
$script = {
    #Get-VBRBackupSession | ?{$_.JobType -eq "Backup"} | Where-Object { $_.CreationTime -ge (Get-Date).AddDays(-1) } | 
    #Select-Object JobName, JobType, CreationTime, EndTime, Result, State, 
    #@{Name="BackupSize"; Expression = { $_.BackupStats.BackupSize}} |
    #Sort-Object CreationTime | 
    #Format-Table -AutoSize
    $ListJobs = Get-VBRJob | ?{$_.JobType -eq "Backup"} | Sort-Object typetostring, name
    $ListSession = [Veeam.Backup.Core.CBackupSession]::GetAll()
    $summary = @()
    foreach ($job in $ListJobs) {
        $lastSession = $ListSession | where {$_.JobId -eq $job.Id} | sort -Property EndTime | select -Last 1
        $summary += $job | select Name, @{n='LastResult';e={$lastSession.Result}}
        }
    $summary
}

# Выполнение команды на удаленном сервере
try {
    Invoke-Command -ComputerName $serversWithDomain -Credential $creds -ScriptBlock $script -ErrorAction Stop
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "Detailed Error: $($_.Exception.Message)" -ForegroundColor Red
}
