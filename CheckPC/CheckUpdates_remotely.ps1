# GET UPDATES
Clear-Host
$currentTime = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
Write-Host "Check updates Yavne - $currentTime" -ForegroundColor Yellow
Write-Host ""

$domain = "******"
$username = "$domain\******"
$password = ConvertTo-SecureString "******" -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

$servers = @("******", "******")
$serversWithDomain = $servers | ForEach-Object { "$_.$domain" }

# Проверка обновлений на каждом сервере
foreach ($server in $serversWithDomain) {
    $updateScript = {
        param ($serverName)

        # Получение информации о доступных обновлениях
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateupdateSearcher()
        $searchResult = $updateSearcher.Search("Type='Software' and IsHidden=0 and IsInstalled=0")
        $updates = $searchResult.Updates
            
        #Write-Host ""
        if ($updates.Count -gt 0) {
            #foreach ($update in $updates) {
                #$title = $update.Title
                #$kbArticleIDs = $update.KBArticleIDs
                #$securityBulletinIDs = $update.SecurityBulletinIDs
                #$msrcSeverity = $update.MsrcSeverity
                #$lastDeploymentChangeTime = $update.LastDeploymentChangeTime
                #$moreInfoUrls = $update.MoreInfoUrls

                #Write-Host "--------------------------------------------"
                #Write-Host "Title: $title"
                #if ($kbArticleIDs) { Write-Host "KB #: $kbArticleIDs" }
                #if ($msrcSeverity) { Write-Host "Rating: $msrcSeverity" }
                #if ($lastDeploymentChangeTime) { Write-Host "Released: $lastDeploymentChangeTime" }
                #if ($moreInfoUrls) { Write-Host "More Info: $moreInfoUrls" }
                #Write-Host ""
            #}
            Write-Host "$serverName - $($updates.Count)" -ForegroundColor Red
        } else {
            Write-Host "$serverName - No Updates Found" -ForegroundColor Green
        }
    }

    # Выполнение скрипта на удаленном сервере через Invoke-Command
    Invoke-Command -ComputerName $server -Credential $creds -ScriptBlock $updateScript -ArgumentList $server
}
                Write-Host "--------------------------------------------"