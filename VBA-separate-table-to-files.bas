Attribute VB_Name = "Module1"
Sub VBA_separate_table_to_files()

    Application.DisplayAlerts = False 'отключаем оповещение об удалении листа
    
    'Сохраняем ссылку на исходную книгу
    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = ActiveWorkbook
    
    'Первая часть исполняемого макроса - делит таблицу на отдельные листы

    For Each cell In Range("ID") 'в скобках указываем наимнование индексирующей таблицы
        Range("Общая").AutoFilter Field:=8, Criteria1:=cell.Value 'в скобках указываем наименование индексируемой таблицы
        Range("Общая[#All]").SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = cell.Value
        ActiveSheet.UsedRange.Columns.AutoFit
    Next cell

'Вторая часть исполняемого макроса - делит листы на отдельные файлы

    Dim s As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    For Each s In wb.Worksheets
        'Пропускаем лист "Данные" при создании отдельных файлов
        If s.Name <> "Данные" Then
            s.Copy                                                  'сохраняем лист как новый файл
            ActiveWorkbook.SaveAs wb.Path & "\" & s.Name & ".xlsx"  'сохраняем файл
            ActiveWorkbook.Close                                    'закрываем все созданные файлы
        End If
    Next s

'Третья часть исполняемого макроса - удаляет листы из первого файла, кроме исходного

    Dim mySheet As Worksheet
    For Each mySheet In sourceWorkbook.Worksheets
        If mySheet.Name <> "Данные" Then
            mySheet.Delete
        End If
    Next mySheet

    Application.DisplayAlerts = True 'включаем оповещение об удалении листа обратно
    
    'Закрываем исходный файл без сохранения
    sourceWorkbook.Close SaveChanges:=False

End Sub
