# Указываем путь к вашему Excel файлу
$excelFilePath = "C:\Users\Elchin\Documents\NewYavneTests.xlsx"
# Создаем объект Excel
$excel = New-Object -ComObject Excel.Application
# Делаем Excel видимым (можно отключить, если не нужно видеть процесс)
$excel.Visible = $true
# Открываем Excel файл
$workbook = $excel.Workbooks.Open($excelFilePath)
# Получаем количество листов в книге
$sheetCount = $workbook.Sheets.Count
# Получаем последний лист
$lastSheet = $workbook.Sheets.Item($sheetCount)
# Делаем копию последнего листа и вставляем после него
$lastSheet.Copy([ref]$workbook.Sheets.Item($sheetCount))
# Получаем последний лист
$lastSheet = $workbook.Sheets.Item($sheetCount)
#rename header 1 line
$lastSheet.Cells.Item(1,2) = (Get-Date).ToString("dd.MM.yy")
#rename list name
$lastSheet.Name = (Get-Date).ToString("dd.MM")














# Сохраняем изменения
$workbook.Save()





# Закрываем книгу и Excel
$workbook.Close($true)
$excel.Quit()

# Освобождаем COM-объекты
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($lastSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Убираем мусор
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
