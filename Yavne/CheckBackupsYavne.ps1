#ALL BACKUPS VEEAM DISC AND WITHOUT DUBLE

# Очищаем экран
Clear-Host
Write-Host "Check Backups..." -ForegroundColor Yellow
Write-Host ""

# Указываем домен и учетные данные
$domain = "***"
$username = "$domain\administrator"
$password = ConvertTo-SecureString "***" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

# Список серверов
$servers = "YavneVeeam"
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

# Скрипт для выполнения на удаленном сервере
$script = {
    Get-VBRBackupSession | Where-Object { $_.CreationTime -ge (Get-Date).AddDays(-1) } | 
    Select-Object JobName, JobType, CreationTime, EndTime, Result, State, 
    @{Name="BackupSize"; Expression = { $_.BackupStats.BackupSize }} | 
    Sort-Object CreationTime | 
    Format-Table -AutoSize
}

# Выполнение команды на удаленном сервере
try {
    Invoke-Command -ComputerName $serversWithDomain -Credential $creds -ScriptBlock $script -ErrorAction Stop
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "Detailed Error: $($_.Exception.Message)" -ForegroundColor Red
}
