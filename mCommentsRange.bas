Attribute VB_Name = "mCommentsRange"

' Setting style comments
' Status: work
' Used in: Set_Comments_ZTRK, Set_CommentsBBDate

Sub Set_CommentStyle(ByVal cell As Range)
    
    With cell.Comment.Shape
        .TextFrame.AutoSize = True
        '.AutoShapeType = msoShapeRoundedRectangle
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.1
        .line.Visible = msoTrue ' Показываем границу комментария
        .line.ForeColor.RGB = RGB(0, 0, 0) ' Задаем цвет границы (черный)
        .line.Weight = 0.1 ' Задаем толщину границы
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    End With
    
End Sub

' Устанавить многоуровневые комментарии для каждой ячейки в столбце OL
' Status: work
' Depends: GetVLookupResult

Sub Set_Comments_ZTRK()
    
    Application.ScreenUpdating = False
    
    Dim ws_main As Worksheet
    Dim ws_buffer As Worksheet
    Dim lastRow As Long
    Dim targetRange As Range
    
    Dim item_showbox As Variant
    Dim piece_box As Variant
    Dim box_layer As Variant
    Dim box_pallet As Variant
    
    Set ws_main = ActiveWorkbook.Worksheets("Sheet1")
    Set ws_buffer = ActiveWorkbook.Worksheets("Буфер")
    
    lastRow = ws_main.Cells(ws_main.Rows.Count, "A").End(xlUp).Row
    
    ' Определяем диапазон для обработки
    Set targetRange = ws_main.Range("OL5:OL" & lastRow)
    
    targetRange.ClearComments
    
    ' Цикл по каждой ячейке в заданном диапазоне
    For Each cell In targetRange
        ' Получаем значения из листа "Буфер" с помощью функции GetVLookupResult

        item_showbox = GetVLookupResult(ws_main.Range("$B" & cell.Row).Value, ws_buffer.Range("$B:$AG"), 32)
        
        If item_showbox = "Не настроен ЛЕ = 21" Then
            item_showbox = "-"
        End If
                
        piece_box = GetVLookupResult(ws_main.Range("$B" & cell.Row).Value, ws_buffer.Range("$B:$H"), 5)
        box_layer = GetVLookupResult(ws_main.Range("$B" & cell.Row).Value, ws_buffer.Range("$B:$H"), 6)
        box_pallet = GetVLookupResult(ws_main.Range("$B" & cell.Row).Value, ws_buffer.Range("$B:$H"), 7)
        
        ' Очищаем комментарий ячейки
        cell.ClearComments
        
        ' Устанавливаем многоуровневый комментарий
        cell.AddComment "Затарка" & ChrW(10) & " " & ChrW(10) & _
                        "в шоубоксе: " & item_showbox & " шт." & ChrW(10) & _
                        "в коробке: " & piece_box & " шт." & ChrW(10) & _
                        "в слое: " & box_layer & " кор." & ChrW(10) & _
                        "в паллете: " & box_pallet & " кор."
        
        ' Настраиваем стиль комментария
        Call Set_CommentStyle(cell)
    
    Next cell
    
    ' Включаем обновление экрана после завершения операций
    Application.ScreenUpdating = True
    
End Sub

' best before date. Установка многоуровневых комментариев для затарки.
' Status: work
' Depends: GetVLookupResult

Sub Set_Comments_BBDate()
    Application.ScreenUpdating = False
    
    Dim ws_main As Worksheet
    Dim ws_buffer As Worksheet
    Dim lastRow As Long
    Dim targetRange As Range
    
    Dim osdDays As Variant
    Dim percentPostkaOsg As Variant
    Dim warehouse As Variant
    Dim shop As Variant
    Dim maxTzDays As Variant
    
    ' Укажите ваш лист "Sheet1"
    Set ws_main = ActiveWorkbook.Worksheets("Sheet1")
    
    ' Укажите ваш лист "Буфер"
    Set ws_buffer = ActiveWorkbook.Worksheets("Буфер")
    
    lastRow = ws_main.Cells(ws_main.Rows.Count, "A").End(xlUp).Row
    
    ' Определяем диапазон для обработки (в данном случае, столбцы NU - NZ)
    Set targetRange = ws_main.Range("NU5:NU" & lastRow)
    
    ' Удаляем все имеющиеся комментарии в диапазоне
    targetRange.ClearComments
    
    ' Цикл по каждой ячейке в заданном диапазоне
    For Each cell In targetRange
        ' Заполняем переменные с использованием формул
        SG_max = GetVLookupResult(ws_main.Range("$A" & cell.Row).Value, ws_buffer.Range("$A:$AT"), 40)
        SG_percent = GetVLookupResult(ws_main.Range("$A" & cell.Row).Value, ws_buffer.Range("$A:$AP"), 39)
        OSG_warehouse = GetVLookupResult(ws_main.Range("$A" & cell.Row).Value, ws_buffer.Range("$A:$AT"), 41)
        OSG_shop = GetVLookupResult(ws_main.Range("$A" & cell.Row).Value, ws_buffer.Range("$A:$AT"), 42)
        
        'OSG_max = ws_main.Range("M" & cell.row).Value - ws_main.Range("N" & cell.row).Value - 1
        
        ' Устанавливаем многоуровневый комментарий к активной ячейке
        cell.ClearComments
        cell.AddComment "Срок годности" & ChrW(10) & " " & ChrW(10) & _
                        "Control SG: " & SG_max & " дн." & ChrW(10) & _
                        "% SG KA: " & SG_percent & ChrW(10) & _
                        "Warehouse: " & OSG_warehouse & " дн." & ChrW(10) & _
                        "Magazine: " & OSG_shop & " дн." & ChrW(10) & _
                        "Max TZ for SG: " & maxTzDays & " дн."
        
        ' Настраиваем стиль комментария
        Call Set_CommentStyle(cell)
    
    Next cell
    
    ' Включаем обновление экрана после завершения операций
    Application.ScreenUpdating = True
    
End Sub


' Функция для выполнения VLOOKUP с обработкой ошибок
' Status: work
' used in: Set_Comments_TZ, Set_Comments_Zatarka

Function GetVLookupResult(lookupValue As Variant, lookupRange As Range, columnNumber As Long) As Variant
    On Error Resume Next
        GetVLookupResult = Application.WorksheetFunction.VLookup(lookupValue, lookupRange, columnNumber, 0)
    On Error GoTo 0
    
    ' If error - return empty string
    If IsError(GetVLookupResult) Then
        GetVLookupResult = ""
    End If
End Function
