Attribute VB_Name = "non_delivery"
'

Sub create_a_message_for_a_non_delivery()

    ' Объявление переменных
    Dim selected_value As Variant ' Значение, выбранное в ячейке
    Dim source_book As Workbook ' Книга, в которой производится выбор значения
    Dim target_book As Workbook ' Книга, в которой ищутся данные
    Dim target_sheet As Worksheet ' Лист, в котором ищутся данные
    Dim found_range As Range ' Диапазон найденных значений
    Dim current_cell As Range ' Текущая ячейка в процессе перебора
    Dim email_table As String ' HTML-код таблицы для тела письма
    Dim out_app As Object ' Объект Outlook Application
    Dim out_mail As Object ' Объект Outlook MailItem
    Dim to_recipients As String ' Адресаты в поле "Кому"
    Dim cc_recipients As String ' Адресаты в поле "Копия"
    Dim subject_line As String ' Тема письма
    Dim i As Integer ' Счетчик для строк таблицы
    Dim j As Integer ' Счетчик для столбцов таблицы
    
    ' Указываем имя книги и листа, где будут производиться поиски
    Set source_book = Workbooks("20.06.xlsb")
    Set target_book = Workbooks("Недопоставка.xlsx")
    Set target_sheet = target_book.Sheets("Комментарии")
    
    ' Получение значения из выделенной ячейки
    selected_value = Selection.Value

    Debug.Print "Selected value: " & selected_value ' Отладочный вывод выбранного значения
    
    ' Получение адресата из столбца 10
    to_recipients = source_book.Sheets("Расширенный").Cells(Selection.Row, 161).Value

    Debug.Print "To recipients: " & to_recipients ' Отладочный вывод адресата

    ' Поиск всех строк, содержащих значение
    Set found_range = target_sheet.Cells.Find(What:=selected_value, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    
    If Not found_range Is Nothing Then

        ' Начало таблицы для тела письма
        email_table = "<table border='1' style='border-collapse:collapse'>"
        email_table = email_table & "<tr>"
        
        ' Добавляем заголовки таблицы (первая строка на листе)
        For Each current_cell In target_sheet.Rows(1).Cells

            email_table = email_table & "<th>" & current_cell.Value & "</th>"

        Next current_cell

        email_table = email_table & "</tr>"
        
        ' Перебор найденных значений и добавление их в таблицу
        Do
            email_table = email_table & "<tr>"

            For Each current_cell In target_sheet.Rows(found_range.Row).Cells
                email_table = email_table & "<td>" & current_cell.Value & "</td>"
            Next current_cell

            email_table = email_table & "</tr>"
            
            ' Переход к следующему найденному значению
            Set found_range = target_sheet.Cells.FindNext(found_range)

            Debug.Print "Found value at row: " & found_range.Row ' Отладочный вывод строки найденного значения
            
        Loop While Not found_range Is Nothing And found_range.Address <> target_sheet.Cells.Find(What:=selected_value, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Address
        
        ' Конец таблицы для тела письма
        email_table = email_table & "</table>"
        
    Else
        ' Создание пустой таблицы если значение не найдено
        Debug.Print "Value not found, creating an empty table" ' Отладочный вывод при отсутствии найденных значений

        email_table = "<table border='1' style='border-collapse:collapse'>"

        For i = 1 To 4 ' 4 строки

            email_table = email_table & "<tr>"

            For j = 1 To 10 ' 10 столбцов

                email_table = email_table & "<td>&nbsp;</td>" ' Пустая ячейка

            Next j

            email_table = email_table & "</tr>"

        Next i

        email_table = email_table & "</table>"

    End If

    ' Настройка адресатов и темы письма
    cc_recipients = "" ' Оставляем пустым или добавляем необходимый список
    subject_line = "Отчет по недопоставкам"

    Debug.Print "Subject line: " & subject_line ' Отладочный вывод темы письма
    
    ' Создание и отправка письма
    Set out_app = CreateObject("Outlook.Application")
    Set out_mail = out_app.CreateItem(0)
    
    With out_mail
        .To = to_recipients ' Указываем получателей в поле "Кому"
        .CC = cc_recipients ' Указываем получателей в поле "Копия"
        .Subject = subject_line ' Указываем тему письма
        .HTMLBody = "Здравствуйте,<br><br>Ниже приведен отчет по недопоставкам:<br><br>" & email_table & "<br><br>С уважением,<br>Ваше имя"
        .Display
    '.Send ' для автоматической отправки
    End With
    
    Debug.Print "Email created and displayed/sent" ' Отладочный вывод о создании и отображении/отправке письма
    
    ' Очистка объектов
    Set out_mail = Nothing
    Set out_app = Nothing

Debug.Print "Outlook objects cleared" ' Отладочный вывод об очистке объектов Outlook

End Sub

