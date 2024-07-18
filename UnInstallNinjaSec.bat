wmic service 'NinjaRMMAgent' get pathname | find "C:\" > C:\temp\NinjaRMMuninstaller.txt
SET /p NinjaFolder= < C:\temp\NinjaRMMuninstaller.txt
SET Uninstaller=%NinjaFolder:NinjaRMMAgentPatcher.exe"=uninstall.exe" --mode unattended%
%Uninstaller%