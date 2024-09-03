Attribute VB_Name = "add_func"

' status: work
' визуальное оформление примечания(комментария) ячейки

Sub install_comment_style(cell As Range)
    
    With cell.comment.Shape
        .TextFrame.AutoSize = True
        '.AutoShapeType = msoShapeRoundedRectangle 'закругление углов всплывающей подскахки
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.1
        .Line.Visible = msoTrue ' Показываем границу комментария
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' Задаем цвет границы (черный)
        .Line.Weight = 0.1 ' Задаем толщину границы
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .TextFrame.MarginLeft = 1000   ' Отступ слева
        .TextFrame.MarginRight = 1000  ' Отступ справа
        .TextFrame.MarginTop = 5000  ' Отступ сверху
        .TextFrame.MarginBottom = 5000 ' Отступ снизу
        
    End With
    
End Sub

' status: work
' узнать буквы столбца, зная номер

Function column_name(ByVal column_number As Long) As String
    If column_number <= 0 Then
        column_name = "error"
    Else
        column_name = Left(Cells(1, column_number).Address(False, False), Len(Cells(1, column_number).Address(False, False)) - 1)
    End If
    
End Function

' status: work
' После копирования данных из АПЦ!СЛИП-ЧЕК|ТД
'   установить фильтр по столбцу "Специалист" в ЗАКАЗНИК!СЛИП-ЧЕК|ТД

Sub enable_filter_on_manager(ByRef ws As Worksheet)
    
    If ws.FilterMode Then
        ws.ShowAllData ' сбросить все активные фильтры
    End If
    
    If ws.AutoFilterMode = False Then
        ws.Range("1:1").AutoFilter Field:=33, Criteria1:=Application.UserName
    End If

End Sub

' status: work
' (СЛИП-ЧЕК|ТД)
' сортируем данные по столбцам: Название акции (2), Специалист (33), Название КА (35)

Sub sort_sctd(ByRef ws As Worksheet)

    Dim sort_range As Range
    Set sort_range = ws.Range("A1").CurrentRegion
    
    With ws.Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(2), Order:=xlAscending ' Проверка заложенных объемов %
        .SortFields.Add Key:=ws.Columns(6), Order:=xlAscending ' Акция с
        .SortFields.Add Key:=ws.Columns(33) ' Менеджер (Специалист)
        .SortFields.Add Key:=ws.Columns(35) ' Наименование КА
        .SetRange sort_range ' диапазон действия фильтров
        .Apply ' Активаровать фильтры
    End With

End Sub

' status: work
' Функция поиска значения последней заполненой строки на листе

Function find_last_row(ByVal ws As Worksheet) As Long

    Dim last_row As Long
    last_row = ws.Cells.Find(What:="*" _
                        , LookAt:=xlPart _
                        , LookIn:=xlFormulas _
                        , SearchOrder:=xlByRows _
                        , searchdirection:=xlPrevious).row
    find_last_row = last_row
    
End Function

' status: work
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

' status: work
' OFF

Sub turn_off_functionalities()

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

End Sub

' status: work
' ON
Sub turn_on_functionalities()

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub

' status: debug
' кастомное сообщение с входящими аргументами
' зависимости: close_custom_message()

'Sub show_custom_msgbox(ByVal text_caption As String)
'    Dim msgbox_form As Object
'    Set msgbox_form = CreateObject("Forms.Form")
'
'    With msgbox_form
'        .Width = 300
'        .Height = 150
'        .Caption = "Update"
'        .BorderStyle = 1 ' Фиксированный размер
'
'        ' Текст сообщения
'        Dim label As Object
'        Set label = .Controls.Add("Forms.Label.1", "MsgLabel")
'        With label
'            .Caption = text_caption
'            .Left = 10
'            .Top = 10
'            .Width = 280
'            .Height = 80
'            .TextAlign = 2 ' Выравнивание текста по центру
'        End With
'
'        ' Кнопка "OK"
'        Dim okButton As Object
'        Set okButton = .Controls.Add("Forms.CommandButton.1", "OkButton")
'        With okButton
'            .Caption = "OK"
'            .Width = 100
'            .Height = 30
'            ' Рассчитываем положение кнопки по горизонтали
'            .Left = (.Parent.Width - .Width) / 2
'            ' Рассчитываем положение кнопки по вертикали
'            .Top = (.Parent.Height - .Height) / 2
'            .Default = True ' Сделать кнопку по умолчанию (для нажатия клавиши Enter)
'        End With
'
'        ' Обработчик события для кнопки "OK"
'        msgbox_form.Controls("OkButton").OnAction = "close_custom_message_box"
'
'        .Show
'    End With
'
'End Sub

' status: debug
' используется в: show_custom_msgbox()

'Sub close_custom_message_box()
''    Unload Me
'End Sub

