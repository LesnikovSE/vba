Attribute VB_Name = "update_data"

'
' Status: works
' $AU5:EH..

Sub new_formulas_to_weeks()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long '
    Dim target_range As Range '
    Dim i As Long ' переменная счетчик
    Dim arr_info_week() As Variant ' массив с рабочей информации о неделях
    ' arr_info_week(i)(*):
    ' (0) - номер недели
    ' (1) - первый столбец диапазона обрабатываемой недели
    ' (2) - последний столбец диапазона обрабатываемой недели
    ' (3) - связанный столбец из "Расчет заказа | Осталось заказать в шт."
    Dim arr_formuls() As Variant ' массив с шаблонами формул (ВПР) для недели
    ' arr_formuls(i)(*):
    ' (0) - Тип акции
    ' (1) - Дата первого заказа согласно графика
    ' (2) - Начало закупочной цены
    ' (3) - Конец закупочной цены
    ' (4) - Начало акции
    ' (5) - Конец акции
    ' (6) - Объем на акцию в шт.
    ' (7) - Объем минимальной выклади на акцию в шт.
    ' (8) - Объем по плану отгрузки в кор.
    ' (9) - Объем по плану отгрузки в шт.
    ' (10) - Пусто
    ' (11) - Пусто
    ' (12) - Осталось заказать на акцию в кор.
    ' (13) - Осталось заказать на акцию в шт.

    '
    ' В "Расчет заказа - осталось заказать в шт." столбец для 8 недели отсутствует
    '

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    arr_info_week = Array( _
                Array(ws.Cells(3, 41), "AO", "BB", "OM"), _
                Array(ws.Cells(3, 55), "BC", "BP", "ON"), _
                Array(ws.Cells(3, 69), "BQ", "CD", "OO"), _
                Array(ws.Cells(3, 83), "CE", "CR", "OP"), _
                Array(ws.Cells(3, 97), "CS", "DF", "OQ"), _
                Array(ws.Cells(3, 111), "DG", "DT", "OR"), _
                Array(ws.Cells(3, 125), "DU", "EH", ""))

    For i = LBound(arr_info_week, 1) To UBound(arr_info_week, 1)
    
        ' заполняем массив формулами (в соответствии с номером обрабатываемой недели)
        If i < 6 Then
            arr_formuls = Array( _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$AM, 35, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$D:$BG, 1, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$K, 7, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$L, 8, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$F, 2, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$G, 3, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$X, 20, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 63, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 17, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 18, 0), """")", _
                "", _
                "", _
                "=IFERROR(" & arr_info_week(i)(3) & "5 / $GH5, """")", _
                "=IFERROR(" & arr_info_week(i)(3) & "5, """")")
        Else
            arr_formuls = Array( _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$AM, 35, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$D:$BG, 1, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$K, 7, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$L, 8, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$F, 2, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$G, 3, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$X, 20, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 63, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 17, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$BO, 18, 0), """")", _
                "", _
                "", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$X, 19, 0), """")", _
                "=IFERROR(VLOOKUP($A:$A, '" & CStr(arr_info_week(i)(0)) & "_нед'!$E:$X, 20, 0), """")")
        End If

        ' задаем диапазон в который установим формулы ВПР из массива с формулами
        Set target_range = ws.Range(arr_info_week(i)(1) & "5:" & arr_info_week(i)(2) & last_row)

        With target_range
            .ClearContents
            .ClearComments
            .formula = arr_formuls
        End With

    Next i ' следующая неделя

    Call add_func.turn_on_functionalities

    ' Call add_func.show_custom_msgbox("УСПЕШНО")
    MsgBox "обновление формул для блоков недель выполнено", Title:=" [ ! ] "

End Sub

'
' status: debug
' $EI5:EL..

Sub new_formulas_to_SCTD()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim target_range As Range
    Dim arr_formuls() As Variant

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    arr_formuls = Array("", "", "", "")

    Set target_range = ws.Range("EI5:EL" & last_row)

    With target_range
        .ClearComments
        .ClearContents
        .Formuls = arr_formuls
    End With

    Call add_func.turn_on_functionalities

    MsgBox "обновление формул в РАСШИРЕННЫЙ!СЧ_ТД выполнено", Title:=" [ ! ] "

End Sub

'
' Status: work
' $OM5:$OR..

Sub new_formulas_to_order()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim target_range As Range
    Dim arr_formuls() As Variant

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    Set target_range = ws.Range("OM5:OR" & last_row)

    arr_formuls = Array("=AU5", "=BI5", "=BW5", "=CK5", "=CY5", "=DM5")

    With target_range
        .ClearContents
        .ClearComments
        .formula = arr_formuls
    End With

    Call add_func.turn_on_functionalities

    MsgBox "обновление формул в РАСШИРЕННЫЙ!$OM5:OR.. выполнено ", Title:=" [ ! ] "

End Sub

' Сдвинуть и изменить формулы с сохранением данных в "Расчет заказа - осталось заказать"
' status: debug
' $OM5:$OR..

Sub shift_formulas_to_order_with_save_data()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim arr_range() As Variant, arr_name() As Variant, arr_new() As Variant
    Dim target_range As Range
    Dim i As Long, j As Long

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    Set target_range = ws.Range("OM5:OR" & last_row)

    ' Получаем формулы из диапазона и сохраняем их в массив arr_range
    arr_range = ws.Range("OM5:OR" & last_row).formula

    ' Определяем массив имен столбцов
    arr_name = Array("AU", "BI", "BW", "CK", "CY", "DM")

    ' Переопределяем размер массива arr_new, чтобы он соответствовал размеру arr_range
    ReDim arr_new(1 To UBound(arr_range, 1), 1 To UBound(arr_range, 2))

    ' Проходим по всему диапазону и выполняем необходимые действия
    For i = LBound(arr_range, 1) To UBound(arr_range, 1)
'        Debug.Print "I - " & i
        For j = LBound(arr_range, 2) To UBound(arr_range, 2)
'            Debug.Print "J - " & j
            If j < 6 Then
'                Debug.Print "Формула: " & arr_range(i, j)
'                Debug.Print "Значение: " & target_range.Cells(i, j).Value
'                Debug.Print "Адресс ячейки: " & target_range.Cells(1, j).Address

                ' Заменяем имя столбца в формуле на новое
'                Debug.Print arr_name(j) & "|" & arr_name(j - 1)
                arr_new(i, j) = Replace(arr_range(i, j + 1), arr_name(j), arr_name(j - 1))
                
'                Debug.Print "Новая формула | arr_new(i, j): " & arr_new(i, j)
            End If

            ' Если мы в последнем столбце, значение = ссылка на ячейку через формулу
            If j = 6 Then
                arr_new(i, j) = arr_name(5) & "5"
'                Debug.Print "Формула ссылка: " & arr_new(i, j)
            End If
        Next j
    Next i

    With target_range
        .ClearComments
        .ClearContents
        .formula = arr_new
    End With

    Call add_func.turn_on_functionalities

    MsgBox "Смещение данных с сохранением значений выполнено", Title:=" [ ! ] "

End Sub


' проверить и записать новую дату закупочной цены
' status: work

Sub new_date_to_purchase_price()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim target_range As Range ' UL5:UL.. столбец в который будут установлены действующие даты ЗЦ
    Dim cell As Range
    Dim row As DataRow
    Const UL As Long = 558
    
        
    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)
    Set target_range = ws.Range("UL5:UL" & last_row)

    For Each cell In target_range
        
        Set row = New DataRow
        row.Initialize ws.Range("A" & cell.row & ":UL" & cell.row)
        
        If row.week1(2) <= Date And Date <= row.week1(3) Then
            ws.Cells(cell.row, UL).Value = _
                Format(row.week1(2), "dd.mm") & " - " & Format(row.week1(3), "dd.mm")
        ElseIf row.week2(2) <= Date And Date <= row.week2(3) Then
            ws.Cells(cell.row, UL).Value = _
                Format(row.week2(2), "dd.mm") & " - " & Format(row.week2(3), "dd.mm")
        ElseIf row.week3(2) <= Date And Date <= row.week3(3) Then
            ws.Cells(cell.row, UL).Value = _
                Format(row.week3(2), "dd.mm") & " - " & Format(row.week3(3), "dd.mm")
        ElseIf row.week4(2) <= Date And Date <= row.week4(3) Then
            ws.Cells(cell.row, UL).Value = _
                Format(row.week4(2), "dd.mm") & " - " & Format(row.week4(3), "dd.mm")
        ElseIf row.week5(2) <= Date And Date <= row.week5(3) Then
            ws.Cells(cell.row, UL).Value = _
                Format(row.week5(2), "dd.mm") & " - " & Format(row.week5(3), "dd.mm")
        End If

    Next cell
    
    Call add_func.turn_on_functionalities

    MsgBox "Даты действующей закупочной цены установлены", Title:=" [ ! ] "

End Sub
