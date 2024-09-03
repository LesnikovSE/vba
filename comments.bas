Attribute VB_Name = "comments"

' установить многострочное примечание(комментарий) с информацией о акции в "Расчет - осталось заказать в шт."
' status: work

Sub new_comments_to_ordering()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim arr_data() As Variant ' данные всего листа Расширенный (без заголовков)
    Dim arr_order As Variant ' данный из диапазона Расширенный!Расчет-осталось заказать в шт. (6 недель, без заголовков)
    Dim comment_text As String ' временная переменная для многоуровневого комментария
    ' тип акции
    ' дата 1 заказ согласно графику заказов
    ' дата акции: с .. по ..
    ' цена: с .. по ..
    ' объем на акцию: кор. | шт.
    ' минимальная выкладка
    ' план отгрузки: кор. | шт.
    Dim arr_week_info() As Variant ' инфо, откуда берутся данные для комментариев
    ' arr_week_info(i)( * )
    ' (0)  - название КА
    ' (1)  - тип акции
    ' (2)  - дата первого заказа согласно графику заказов
    ' (3)  - дата начала ЗЦ
    ' (4)  - дата окончания ЗЦ
    ' (5)  - дата начала акции
    ' (6)  - дата окончания акции
    ' (7)  - объем в шт.
    ' (8)  - минимальная выкладка
    ' (9)  - план отгрузки в кор.
    ' (10) - план отгрузки в шт.
    Dim i As Long ' переменная счетчик строк
    Dim j As Long ' переменная счетчик столбцов

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    last_row = add_func.find_last_row(ByVal ws)

    arr_week_info = Array( _
        Array(147, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50), _
        Array(147, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64), _
        Array(147, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78), _
        Array(147, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92), _
        Array(147, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106), _
        Array(147, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120), _
        Array(147, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134))

    arr_data = ws.Range("A5:UL" & last_row).Value
    arr_order = ws.Range("OM5:OR" & last_row).Value

    For i = LBound(arr_order, 1) To UBound(arr_order, 1)
        For j = LBound(arr_order, 2) To UBound(arr_order, 2)

            ' пустая ячейка не комментируется
            If arr_order(i, j) <> "" And Not IsEmpty(arr_order(i, j)) Then

                ' собираем строку-комментарий из данных в строке
                comment_text = _
                    " " & arr_data(i, arr_week_info(j - 1)(0)) & vbCrLf & _
                    " акция: " & arr_data(i, arr_week_info(j - 1)(1)) & vbCrLf & vbCrLf & _
                    " 1 заказ: " & Format(arr_data(i, arr_week_info(j - 1)(2)), "dd.mm") & vbCrLf & _
                    " даты: с " & Format(arr_data(i, arr_week_info(j - 1)(4)), "dd.mm") & " по " & Format(arr_data(i, arr_week_info(j - 1)(5)), "dd.mm") & vbCrLf & _
                    " цены: с " & Format(arr_data(i, arr_week_info(j - 1)(3)), "dd.mm") & " по " & Format(arr_data(i, arr_week_info(j - 1)(4)), "dd.mm") & vbCrLf & _
                    " объем: " & Round(arr_data(i, arr_week_info(j - 1)(7)) / arr_data(i, 190), 1) & " кор. | " & arr_data(i, arr_week_info(j - 1)(7)) & " шт." & vbCrLf & _
                    " мин. выкладка: " & arr_data(i, arr_week_info(j - 1)(8)) & vbCrLf & _
                    " план отгр.: " & arr_data(i, arr_week_info(j - 1)(9)) & " кор. | " & arr_data(i, arr_week_info(j - 1)(10)) & " шт."

                ' устанавливаем комментарий для ячейки на листе Расширенный!OM5:OR..
                ' (+ 4)   - количество строк заголовков на листе Расширенный
                ' (+ 402) - количество столбцов слева от нужного столбца

                With ws.Cells(i + 4, j + 402)
                    .ClearComments
                    .AddComment Text:=comment_text
                    .comment.Visible = False
                    ' визуальные настройки примечания(комментария)
                    Call add_func.install_comment_style(ws.Cells(i + 4, j + 402))
                End With

            End If

        Next j ' следующая ячейка в строке i
    Next i ' следующая строка диапазона OM5:OR..

    Call add_func.turn_on_functionalities

    MsgBox "Комментарии в ""Расчет - осталось заказать в шт."" установлены", Title:="[ ! ]"

End Sub
