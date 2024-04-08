#Поиск дубликатов файлов с помощью PowerShell
Get-ChildItem –path C:\Share\ -Recurse | Get-FileHash | Group-Object -property hash | Where-Object { $_.count -gt 1 } | ForEach-Object { $_.group | Select-Object Path, Hash }


#Если файлов в каталоге много и для каждого считать хэш, это займет довольно много времени. 
#Проще сначала сравнить файлы по размеру (это готовый атрибут файла, который не надо вычислять). 
#Хэш будем получать только для файлов с одинаковым размером:
$file_dublicates = Get-ChildItem –path C:\Share\ -Recurse| Group-Object -property Length| Where-Object { $_.count -gt 1 }| Select-Object –Expand Group| Get-FileHash | Group-Object -property hash | Where-Object { $_.count -gt 1 }| ForEach-Object { $_.group | Select-Object Path, Hash }



#Можно предложить пользователю выбрать файлы, которые можно удалить. 
#Для этого список дубликатов файлов нужно передать конвейером в командлет Out-GridView:
$file_dublicates | Out-GridView -Title "Выберите файлы для удаления" -OutputMode Multiple –PassThru|Remove-Item –Verbose –WhatIf



#Скрипт который предлагает заменять дубликаты файлов на жесткие ссылки
#Приведу полный код скрипта с его сайта здесь (https://www.outsidethebox.ms/20953/):
param(
[Parameter(Mandatory=$True)]
[ValidateScript({Test-Path -Path $_ -PathType Container})]
[string]$dir1,
[Parameter(Mandatory=$True)]
[ValidateScript({(Test-Path -Path $_ -PathType Container) -and $_ -ne $dir1})]
[string]$dir2
)
Get-ChildItem -Recurse $dir1, $dir2 |
Group-Object Length | Where-Object {$_.Count -ge 2} |
Select-Object -Expand Group | Get-FileHash |
Group-Object hash | Where-Object {$_.Count -ge 2} |
Foreach-Object {
$f1 = $_.Group[0].Path
Remove-Item $f1
New-Item -ItemType HardLink -Path $f1 -Target $_.Group[1].Path | Out-Null
#fsutil hardlink create $f1 $_.Group[1].Path

}
#Для запуска файла используйте такой формат команды:
.\hardlinks.ps1 -dir1 d:\fldr1 -dir2 d:\fldr2

#Этот скрипт можно использовать для поиска и замены дубликатов статических файлов (которые не изменяются!) на символические жёсткие ссылки.