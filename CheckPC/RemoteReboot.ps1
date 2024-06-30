Invoke-Command -HostName IP_ADDRESS -UserName administrator -ScriptBlock {Restart-Computer}
Sleep 30  ## wait before checking
$SSH=$null
While(-not $SSH){
  Try  {$SSH = New-PSSession -HostName 'MyServer' -ea Stop}
  Catch{Write-Host 'Waiting for reboot...'}
  sleep 10
}