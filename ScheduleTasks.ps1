#В Windows 10/Windows Server 2016 для управления задачами в планировщике используется PowerShell модуль ScheduledTasks. 
#Список командлетов в модуле можно вывести так:
Get-Command -Module ScheduledTasks

#Вы можете вывести список всех активных заданий планировщика в Windows с помощью команды:
Get-ScheduledTask -TaskPath | ? state -ne Disabled

#Чтобы получить информацию о конкретном задании:
Get-ScheduledTask CheckServiceState_PS| Get-ScheduledTaskInfo

#Вы можете отключить это задание:
Get-ScheduledTask CheckServiceState_PS | Disable-ScheduledTask

#Чтобы включить задание:
Get-ScheduledTask CheckServiceState_PS | Enable-ScheduledTask

#Чтобы запустить задание немедленно (не дожидаясь расписания), выполните:
Start-ScheduledTask CheckServiceState_PS

#Чтобы полностью удалить задание из Task Scheduler:
Unregister-ScheduledTask -TaskName CheckServiceState_PS

#Если нужно изменить имя пользователя, из-под которого запускается задание и, например, режим совместимости, используйте командлет Set-ScheduledTask:
$task_user = New-ScheduledTaskPrincipal -UserId 'winitpro\kbuldogov' -RunLevel Highest
$task_settings = New-ScheduledTaskSettingsSet -Compatibility 'Win7'
Set-ScheduledTask -TaskName CheckServiceState_PS -Principal $task_user -Settings $task_settings

#В этом примере мы создадим задание планировщика, которое во время запускает определённый файл с PowerShell скриптом во время загруки. 
#Задание выполняется с правами системы (System).
$TaskName = "NewPsTask"
$TaskDescription = "Запуск скрипта PowerShell из планировщика"
$TaskCommand = "c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe"
$TaskScript = "C:\PS\StartupScript.ps1"
$TaskArg = "-WindowStyle Hidden -NonInteractive -Executionpolicy unrestricted -file $TaskScript"
$TaskStartTime = [datetime]::Now.AddMinutes(1)
$service = new-object -ComObject("Schedule.Service")
$service.Connect()

$rootFolder = $service.GetFolder("\")
$TaskDefinition = $service.NewTask(0)
$TaskDefinition.RegistrationInfo.Description = "$TaskDescription"
$TaskDefinition.Settings.Enabled = $true
$TaskDefinition.Settings.AllowDemandStart = $true
$triggers = $TaskDefinition.Triggers
#http://msdn.microsoft.com/en-us/library/windows/desktop/aa383915(v=vs.85).aspx
$trigger = $triggers.Create(8)