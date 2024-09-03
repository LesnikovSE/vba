Attribute VB_Name = "mSettings"
Sub ПрименитьНастройки()
    Dim ws As Worksheet
    Dim filePath As String
    Dim settings As String
    Dim settingsArray() As String
    Dim cellInfo As String
    Dim cellRange As Range
    Dim cell As Range
    Dim line As String
    
    ' Замените "Путь_К_Файлу\Файл.txt" на фактический путь к вашему файлу с настройками
    filePath = "C:\Users\GM\Desktop\SavedCellFormatting.txt"
    
    ' Открываем файл с настройками
    Open filePath For Input As #1
    
    ' Определяем диапазон ячеек, к которым будем применять настройки
    Set cellRange = Range("$A5:$UM5")
    
    ' Читаем настройки из файла
    settings = Input$(LOF(1), #1)
    
    ' Разделяем настройки по ячейкам
    settingsArray = Split(settings, "Cell: ")
    
    ' Закрываем файл
    Close #1
    
    ' Проходим по каждой ячейке и применяем настройки
    For Each cell In cellRange
        ' Определяем информацию о ячейке в настройках
        cellInfo = FindCellInfo(settingsArray, cell.Address)
        
        ' Если информация найдена, применяем настройки
        If cellInfo <> "" Then
            ApplyCellSettings cell, cellInfo
        End If
    Next cell
End Sub

Function FindCellInfo(settingsArray() As String, cellAddress As String) As String
    Dim i As Integer
    
    ' Ищем информацию о ячейке в массиве настроек
    For i = LBound(settingsArray) To UBound(settingsArray)
        If InStr(settingsArray(i), cellAddress) > 0 Then
            FindCellInfo = settingsArray(i)
            Exit Function
        End If
    Next i
End Function

Sub ApplyCellSettings(cell As Range, cellInfo As String)
    ' Применяем настройки к ячейке
    ' Ваш код здесь для применения настроек, например:
    ' cell.Font.Name = "Calibri"
    ' cell.Font.Size = 8
    ' cell.Font.Color = RGB(0, 0, 0)
    ' cell.Interior.Color = RGB(255, 192, 0)
    ' и т.д.
End Sub



Sub СохранитьНастройкиJSON()
    Dim settings As Object
    Dim jsonSettings As String
    Dim filePath As String
    
    ' Определяем объект для хранения настроек
    Set settings = CreateObject("Scripting.Dictionary")
    
    ' Заполняем объект настройками из строки 5
    'InsertFormating(settings, Range("5:5"))
    
    ' Преобразуем объект в строку JSON
    jsonSettings = JsonConverter.ConvertToJson(settings)
    
    ' Задаем путь и имя файла для сохранения
    filePath = "C:\Users\GM\Desktop\CellFormatting.json" ' Замените на фактический путь
    
    ' Сохраняем строку JSON в файл
    Open filePath For Output As #1
    Print #1, jsonSettings
    Close #1
    
End Sub

Function InsertFormating(settings As Object, targetRow As Range)
    ' Добавляем настройки в объект
    settings("Font") = targetRow.Font.Name
    settings("FontSize") = targetRow.Font.Size
    settings("FontColor") = targetRow.Font.Color
    settings("InteriorColor") = targetRow.Interior.Color
    ' Добавьте другие настройки по необходимости
End Function