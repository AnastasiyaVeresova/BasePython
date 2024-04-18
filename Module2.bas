Attribute VB_Name = "Module2"
Sub FindAndFillModels()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim modelRange As Range
    Dim cell As Range
    Dim searchWords As Variant
    Dim word As Variant
    Dim found As Boolean
    
    ' Указываем рабочий лист
    Set ws = ThisWorkbook.Sheets("Маркетинговые данные")
    
    ' Определяем последнюю заполненную строку в столбце P
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    
    ' Добавляем новый столбец "Модель" после столбца P
    Set modelRange = ws.Range("P1:P" & lastRow).Offset(0, 1)
    modelRange.Value = "Модель"
    
    ' Слова для поиска
    searchWords = Array("gls", "GT_AMG", "x5", "i3", "с180", "e220", "x1", "x3", "c200", "318", "530d", "e400", "GLE", "cla", "cls", "glc")
    
    ' Проходимся по каждой ячейке в столбце P и ищем слова для заполнения столбца "Модель"
    For Each cell In ws.Range("P1:P" & lastRow)
        found = False
        For Each word In searchWords
            If InStr(1, cell.Value, word, vbTextCompare) > 0 Then
                cell.Offset(0, 1).Value = word
                found = True
                Exit For
            End If
        Next word
        If Not found Then
            cell.Offset(0, 1).Value = "Нет данных"
        End If
    Next cell
End Sub

