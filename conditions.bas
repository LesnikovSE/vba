Attribute VB_Name = "conditions"

'
' status:work

Sub highlight_multiple_of_layer()

    Dim ws As Worksheet
    Dim last_row As Long
    Dim target_range As Range

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    Set target_range = ws.Range("OU5:OU" & last_row)
    target_range.FormatConditions.Delete

    target_range.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ОСТАТ($OU5; $OL5)<>0"
    target_range.FormatConditions(target_range.FormatConditions.Count).SetFirstPriority

    With target_range.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 200, 200)
        .TintAndShade = 0
    End With

End Sub

'
' Установить цвет ячеек в диапазонах (СКЛАД, ИТОГО ПЛТ., РЦ_ПОЛУЧАТЕЛЬ) по номеру РЦ
' status: work
' $OW5:$OW.. (411)

Sub highlight_color_to_rc()
    
    Dim ws As Worksheet
    Dim last_row As Long
    Dim target_range As Range
    Dim condition1 As String
    Dim condition2 As String
    Dim condition3 As String
    Dim arr_rng() As Variant
    
    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ws)

    arr_rng = Array("ET5:ET" & last_row, "OW5:OW" & last_row, "UB5:UB" & last_row)

    Dim rng_name As Variant
    For Each rng_name In arr_rng

        ' Указываем диапазон для правила условного форматирования
        Set target_range = ws.Range(rng_name)

        ' Удаляем все существующие правила
        target_range.FormatConditions.Delete

        ' Условия для форматирования
        condition1 = "=$ET5=70007"
        condition2 = "=$ET5=70011"
        condition3 = "=$ET5=70035"

        ' Добавляем правила для подсветки в зависимости от условий
        With target_range.FormatConditions.Add(Type:=xlExpression, Formula1:=condition1)
            .Interior.Color = RGB(255, 220, 220) ' светло светло красный
        End With

        With target_range.FormatConditions.Add(Type:=xlExpression, Formula1:=condition2)
            .Interior.Color = RGB(255, 255, 220) ' светло светло желтый
        End With

        With target_range.FormatConditions.Add(Type:=xlExpression, Formula1:=condition3)
            .Interior.Color = RGB(220, 255, 220) ' светло светло зеленый
        End With

    Next rng_name

End Sub

'' status: debug
'' подсветить в столбцах "Дифицит по реализации" и "Дифицит по прогнозу"
'' $PL5:$PM.. (428, 429)
'
'Sub highlight_negative_values_in_the_deficit_column()
'
'    Dim ws As Worksheet
'    Dim last_row As Long
'    Dim target_range As Range
'    Dim condition1 As String
'
'    ' Указываем рабочий лист
'    Set ws = ActiveWorkbook.Worksheets("Sheet1")
'    last_row = add_func.find_last_row(ws)
'
'    ' Указываем диапазон для правила условного форматирования
'    Set target_range = ws.Range("$PL5:$PM" & last_row)
'
'    ' Условия для форматирования
'    condition1 = "=0"
'
'    ' Удаляем все существующие правила
'    target_range.FormatConditions.Delete
'
'    ' Добавляем правило для подсветки ячеек с отрицательными значениями
'    With target_range.FormatConditions.Add( _
'                        Type:=xlCellValue, Operator:=xlLess, Formula1:=condition1)
'        .Font.Color = RGB(220, 255, 220)
'    End With
'End Sub


