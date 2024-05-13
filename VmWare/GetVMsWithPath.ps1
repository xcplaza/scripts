#GET VMs Vmware with path
clear-host
Write-host "Connect to Vmware..." -foregroundcolor yellow
Write-host ""

# Подключаем модуль PowerCLI для работы с VMware
Import-Module VMware.PowerCLI

# Подключаемся к vCenter Server или ESXi-хосту
Connect-VIServer 10.50.50.50
Write-host ""

# Получаем список виртуальных машин
$vms = Get-VM

# Выводим имя виртуальной машины и путь к её диску
foreach ($vm in $vms) {
    $vmName = $vm.Name
    $vmDiskPath = $vm.ExtensionData.Config.Hardware.Device | Where-Object {$_.GetType().Name -eq "VirtualDisk"} | Select-Object -First 1 | Select-Object -ExpandProperty Backing | Select-Object -ExpandProperty FileName
    Write-Host "VM: $vmName"
    Write-Host "Path: $vmDiskPath"
    Write-Host "-----------------------------------------------------------"
}

# Отключаемся от vCenter Server или ESXi-хоста
#Disconnect-VIServer -Confirm:$false
