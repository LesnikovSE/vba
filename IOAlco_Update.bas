Attribute VB_Name = "IOAlco_Update"

' Функция поиска значения последней заполненой строки на листе

Function find_last_row(ByVal ws As Worksheet) As Long

    Dim last_row As Long
    last_row = ws.Cells.Find(What:="*" _
                        , LookAt:=xlPart _
                        , LookIn:=xlFormulas _
                        , SearchOrder:=xlByRows _
                        , searchdirection:=xlPrevious).Row
    find_last_row = last_row
    
End Function

' Функция поиска значения последнего заполненого столбца на листе

Function find_last_column(ByVal ws As Worksheet) As Long

    Dim last_column As Long
    last_column = ws.Cells.Find(What:="*" _
                        , LookAt:=xlPart _
                        , LookIn:=xlFormulas _
                        , SearchOrder:=xlByColumns _
                        , searchdirection:=xlPrevious).Column
    find_last_column = last_column

End Function

' OFF

Sub turn_off_functionalities()

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

End Sub

' ON
Sub turn_on_functionalities()

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub

' конвертация строки в одномерный массив
Function convert_row_to_array_one_demension(ByVal rng As Range) As Variant

Dim arr As Variant
Dim arr_one_dim() As Variant
Dim i As Integer

arr = rng.Value

ReDim arr_one_dim(1 To UBound(arr, 2))

For i = 1 To UBound(arr, 2)
    arr_one_dim(i) = arr(1, i)
Next i

convert_row_to_array_one_demension = arr_one_dim

End Function

' конвертация типа данных элементов массива в тестовый тип
Function convert_array_to_text(ByRef arr As Variant) As Variant

    Dim i As Long
    Dim convert_array() As String
    
    ReDim convert_array(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        convert_array(i) = CStr(arr)
    Next i

End Function


' Поиска индекса элемента в массиве
Function find_index(search_value As Variant, arr As Variant) As Long

    Dim i As Long
    
    find_index = -1

    For i = LBound(arr) To UBound(arr)
        
        If arr(i) = search_value Then
            find_index = i
            Exit Function
        End If
    Next i

End Function

'
Sub close_apc_file()

    Dim wb As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name = "_АКЦИЯ_проверка_цен.xlsb" Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
    
End Sub

' применить фильтр к листу "Вход и выход алкоголя!Данные"
Sub sort_sheet()
    
    Dim ws As Worksheet
    Dim sort_range As Range
    
    Set ws = ThisWorkbook.Worksheets("Данные")
    Set sort_range = ws.Range(ws.Cells(5, 1), ws.Cells(find_last_row(ws), find_last_column(ws)))

    With ws.Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(9) ' Импорт
        .SortFields.Add Key:=ws.Columns(7) ' ТК3
        .SortFields.Add Key:=ws.Columns(8) ' КА
        .SortFields.Add Key:=ws.Columns(5) ' Наименование
        .SortFields.Add Key:=ws.Columns(3), Order:=xlAscending ' РЦ (в порядке возрастания)
        .SetRange sort_range ' диапазон действия фильтров
        .Apply ' Активаровать фильтры
    End With

End Sub


' Закупка на РЦ шт. - столбец X (№ 24)
' Длительность промо в дн. - столбец AN (№ 40)
'
Sub Update()

    Call turn_off_functionalities
    Call close_apc_file
    
    Dim wb_apc As Workbook
    Dim ws_main As Worksheet, ws_apc As Worksheet
    Dim last_row As Long, last_column As Long
    Dim path_to_apc As String
    Dim i As Long, j As Long, x As Long
    
    Dim arr_week_column As Variant ' номера столбцов интересующих номеров недель для заполнения
    Dim arr_duration_column As Variant ' номера столбцов для объемов промо по интересующим неделям
    
    
        arr_week_column = Array(93, 94, 95, 96, 97, 98, 99, 100, 101, 102, _
                                103, 104, 105, 106, 107, 108, 109, 110, 111, 112)
    arr_duration_column = Array(133, 134, 135, 136, 137, 138, 139, 140, 141, 142, _
                                143, 144, 145, 146, 147, 148, 149, 150, 151, 152)
        
    ' Задаем сетевой путь к файлу "АПЦ"
    path_to_apc = "\\dixy.local\Departments-HQ\ORPT\Рассылки\УТЗРЦ\_АКЦИЯ_проверка_цен.xlsb"
    
    Set wb_apc = Workbooks.Open(path_to_apc, ReadOnly:=True)
    Set ws_main = ThisWorkbook.Worksheets("Данные")
    
    ' Очистить данные в диапазонах:
    '   - блок "ПРОМО (отгрузки) | шт."
    '   - блок "Длительность | дн."
    
    ws_main.Range(ws_main.Cells(5, CInt(arr_week_column(0))), _
                  ws_main.Cells(find_last_row(ws_main), CInt(arr_week_column(UBound(arr_week_column))))).ClearContents
    
    ws_main.Range(ws_main.Cells(5, CInt(arr_duration_column(0))), _
                  ws_main.Cells(find_last_row(ws_main), CInt(arr_duration_column(UBound(arr_duration_column))))).ClearContents
    
    Dim range_rcid As Range ' диапазон данных для столбца Лист!Данные > Сцепка 1 (А5:А..)
    Dim ar_rcid As Variant ' вспомогательный массив из Лист!Данные > Сцепка 1 (А5:А..)
    Dim arr_rcid As Variant ' итоговый массив из Лист!Данные > Сцепка 1 (А5:А..)
    
    Set range_rcid = ws_main.Range(ws_main.Cells(5, 1), _
                                   ws_main.Cells(find_last_row(ws_main), 1))
    
    ar_rcid = range_rcid.Value
    ReDim arr_rcid(1 To UBound(ar_rcid, 1))
    For i = LBound(ar_rcid) To UBound(ar_rcid)
        arr_rcid(i) = ar_rcid(i, 1)
    Next i
    Erase ar_rcid
    
    Dim week_number_range As Range ' диапазон данных для столбца Лист!Данные > Блок "ПРОМО отгрузки | шт."
    Dim arr_week_number As Variant ' Массив из Лист!Данные > Блок "Длительность | шт."
    Set week_number_range = ws_main.Range(ws_main.Cells(4, CInt(arr_week_column(0))), _
                                          ws_main.Cells(4, CInt(arr_week_column(UBound(arr_week_column)))))
    
    arr_week_number = convert_row_to_array_one_demension(week_number_range)
    
    For i = LBound(arr_week_number) To UBound(arr_week_number)
        arr_week_number(i) = CStr(arr_week_number(i))
    Next i
    
    For Each ws_apc In wb_apc.Sheets
'        Debug.Print "Обрабатываем лист: " + ws_apc.Name
        
        If InStr(1, ws_apc.Name, "_нед", vbTextCompare) > 0 Then
            what_find = CStr(Replace(ws_apc.Name, "_нед", ""))
        Else
            what_find = ws_apc.Name
        End If

        Dim indx As Long
        indx = find_index(what_find, arr_week_number)
        
        If indx <> -1 Then
'            Debug.Print "[ + ] Индекс: " + CStr(indx)
'            Debug.Print " "
            
            Dim arr_ws_apc() As Variant
            
            last_row = find_last_row(ws_apc)
            last_column = find_last_column(ws_apc)
            
            arr_ws_apc = ws_apc.Range(ws_apc.Cells(2, 1), ws_apc.Cells(last_row, last_column)).Value
            
            For Each rcid In arr_rcid ' идем по столбцу Данные!А5А.. (Сцепка 1)
'                Debug.Print "[ -> ] RCID: " + rcid
                
                For i = LBound(arr_ws_apc, 1) To UBound(arr_ws_apc, 1)
                    
                    If rcid = arr_ws_apc(i, 5) Then
'                        Debug.Print "[ + ] RCID: " + CStr(arr_ws_apc(i, 5))
                        
                        ' + 4 (три пустые строки и 4ая с заголовом)
                        ws_main.Cells(find_index(rcid, arr_rcid) + 4, arr_week_column(indx - 1)).Value = arr_ws_apc(i, 24)
                        ws_main.Cells(find_index(rcid, arr_rcid) + 4, arr_duration_column(indx - 1)).Value = arr_ws_apc(i, 40)
                        
                        Exit For ' ws_apc.rcid
'                    Else
'                        Debug.Print "[ - ] Not found: " + rcid
                    End If
                Next i
                
                index_rcid = index_rcid + 1
            Next rcid
'        Else
'            Debug.Print "[ x ] В словаре отсутствует"
'            Debug.Print " "
        End If
        
    Next ws_apc
    
    Call close_apc_file
    Call turn_on_functionalities
    Call sort_sheet
    
    MsgBox "Обновление выполнено", Title:="[ ! ]"
    
End Sub

