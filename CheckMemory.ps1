# Функция для получения средней нагрузки на память за последние 5 минут
function Get-MemoryUsageAverage {
    $counter = "\Memory\% Committed Bytes In Use"
    $memoryUsage = $null

    # Попытка получить исторические данные о нагрузке на память
    $memoryUsage = (Get-Counter -Counter $counter -SampleInterval 1 -MaxSamples 300 | 
                    Select-Object -ExpandProperty CounterSamples | 
                    Measure-Object -Property CookedValue -Average).Average

    # Если исторические данные недоступны, ожидаем 5 минут и повторяем попытку
    if ($memoryUsage -eq $null) {
        Start-Sleep -Seconds 10
        $memoryUsage = (Get-Counter -Counter $counter -SampleInterval 1 -MaxSamples 300 | 
                        Select-Object -ExpandProperty CounterSamples | 
                        Measure-Object -Property CookedValue -Average).Average
    }

    return [math]::Round($memoryUsage, 2)
}

clear-host
Write-host "Check Memory..." -foregroundcolor yellow
Write-host ""

# Функция для определения цвета вывода в зависимости от нагрузки на память
function Set-ConsoleColor {
    param (
        [double]$memoryUsage
    )

    if ($memoryUsage -gt 85) {
        Write-Host "Memory Usage: $memoryUsage%" -ForegroundColor Red
    } else {
        Write-Host "Memory Usage: $memoryUsage%" -ForegroundColor Green
    }
}

# Получаем среднюю нагрузку на память за последние 5 минут
$averageMemoryUsage = Get-MemoryUsageAverage

# Выводим результат с учетом цвета
Set-ConsoleColor -memoryUsage $averageMemoryUsage
