 <# Скрпт по ренейму ПК
 Создаем ексель файл, в который все записываем 
 Вводим имя ПК (старое) и потом вводим новое #>
$excel = New-Object -ComObject Excel.Application$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Add()
$worksheet.Cells.Item(1, 1) = "старое имя"$worksheet.Cells.Item(1, 1).Font.Bold = $true
$worksheet.Cells.Item(1, 2) = "новое имя"$worksheet.Cells.Item(1, 2).Font.Bold = $true
$worksheet.Cells.Item(1, 3) = "Ошибки"$worksheet.Cells.Item(1, 3).Font.Bold = $true
$row = 2
while ($true) {
    $computerName = Read-Host "Введи имя пк, который надо переименовать"
    if ($computerName -eq "q") {        
        break
        }

    $computers = Get-ADComputer -Filter {name -eq $computerName}

    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {        
        $newName = Read-Host "Введи новое имя ПК"
        <# Проверяем на возмодность переимоеновать, отсеиваем ошибки и записываем их в столбец в ексель файле
        а так же выводим сообщение в консоль #>
        try {        
            Rename-Computer -ComputerName $computerName -NewName $newName -DomainCredential $credential -Force -ErrorAction Stop
        Write-Host "Компьютер $computerName переименован в $newName"         
        }
        catch {        
            $err = $_.Exception.Message
            if ($err -match "Отказано в доступе") {
            $err_write = "Ошибка у $computerName : Отказано в доступе"
            Write-Host "--- ОШИБКА У $computerName!!! ---"
            Write-Error $err_write
            $worksheet.Cells.Item($row, 3) = $err_write
            }
        else {            
            Write-Error "Ошибка : $_"
            $err_write = "Ошибка : $_"            
            $worksheet.Cells.Item($row, 3) = $err_write
            }        
            }

        $worksheet.Cells.Item($row, 1) = $computerName
        $worksheet.Cells.Item($row, 2) = $newName

        $row++

        }        
        else {
        Write-Host "Компьютер $computerName недоступен\не в сети."        
        }
   }

<# Форматируем содержимое и сохраняем в файл rename_pc.xlsx#>
$worksheet.Columns.AutoFit()
$worksheet.Rows.AutoFit()$workbook.SaveAs("D:\rename_pc.xlsx")
$workbook.close()$excel.Quit()
