#Get size folder
(Get-ChildItem -Path "C:\temp" -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB