﻿#Get license key
#VK7JG-NPHTM-C97JM-9MPGT-3V66T
Clear-Host
Write-Host "Check license key..." -ForegroundColor Yellow
Write-Host ""

$service = get-wmiObject -query 'select * from SoftwareLicensingService'
if($key = $service.OA3xOriginalProductKey){
	Write-Host 'Product Key:' $service.OA3xOriginalProductKey -ForegroundColor Green
	$service.InstallProductKey($key)
}else{
	Write-Host 'Key not found.' -ForegroundColor red
}