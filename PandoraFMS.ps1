#INSTALL PANDORA FMS

#Pandora FMS server on CentOS 7.x:
curl -Ls https://pfms.me/deploy-pandora | bash

#Pandora FMS server on RHEL 8.x / RockyLinux 8:
curl -sL https://pfms.me/deploy-pandora-el8 | bash

#Pandora FMS on Ubuntu server 22.04:
curl -sL https://pfms.me/deploy-pandora-ubuntu | bash

#Pandora FMS Software agent on *nix:
export PANDORA_SERVER_IP=<PandoraServer IP or FQDN> && curl -Ls https://pfms.me/agent-deploy | bash

#Pandora FMS Software agent on Windows 10 or higher:
Invoke-WebRequest -Uri https://pfms.me/windows-agent -OutFile ${env:tmp}\pandora-agent-windows.exe; & ${env:tmp}\pandora-agent-windows.exe /S –ip<PANDORASERVER IP or NAME>
NET START PandoraFMSAgent