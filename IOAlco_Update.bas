Attribute VB_Name = "IOAlco_Update"

' ������� ������ �������� ��������� ���������� ������ �� �����

Function find_last_row(ByVal ws As Worksheet) As Long

    Dim last_row As Long
    last_row = ws.Cells.Find(What:="*" _
                        , LookAt:=xlPart _
                        , LookIn:=xlFormulas _
                        , SearchOrder:=xlByRows _
                        , searchdirection:=xlPrevious).Row
    find_last_row = last_row
    
End Function

' ������� ������ �������� ���������� ����������� ������� �� �����

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

' ����������� ������ � ���������� ������
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

' ����������� ���� ������ ��������� ������� � �������� ���
Function convert_array_to_text(ByRef arr As Variant) As Variant

    Dim i As Long
    Dim convert_array() As String
    
    ReDim convert_array(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        convert_array(i) = CStr(arr)
    Next i

End Function


' ������ ������� �������� � �������
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
        If wb.Name = "_�����_��������_���.xlsb" Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
    
End Sub

' ��������� ������ � ����� "���� � ����� ��������!������"
Sub sort_sheet()
    
    Dim ws As Worksheet
    Dim sort_range As Range
    
    Set ws = ThisWorkbook.Worksheets("������")
    Set sort_range = ws.Range(ws.Cells(5, 1), ws.Cells(find_last_row(ws), find_last_column(ws)))

    With ws.Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(9) ' ������
        .SortFields.Add Key:=ws.Columns(7) ' ��3
        .SortFields.Add Key:=ws.Columns(8) ' ��
        .SortFields.Add Key:=ws.Columns(5) ' ������������
        .SortFields.Add Key:=ws.Columns(3), Order:=xlAscending ' �� (� ������� �����������)
        .SetRange sort_range ' �������� �������� ��������
        .Apply ' ������������ �������
    End With

End Sub


' ������� �� �� ��. - ������� X (� 24)
' ������������ ����� � ��. - ������� AN (� 40)
'
Sub Update()

    Call turn_off_functionalities
    Call close_apc_file
    
    Dim wb_apc As Workbook
    Dim ws_main As Worksheet, ws_apc As Worksheet
    Dim last_row As Long, last_column As Long
    Dim path_to_apc As String
    Dim i As Long, j As Long, x As Long
    
    Dim arr_week_column As Variant ' ������ �������� ������������ ������� ������ ��� ����������
    Dim arr_duration_column As Variant ' ������ �������� ��� ������� ����� �� ������������ �������
    
    
        arr_week_column = Array(93, 94, 95, 96, 97, 98, 99, 100, 101, 102, _
                                103, 104, 105, 106, 107, 108, 109, 110, 111, 112)
    arr_duration_column = Array(133, 134, 135, 136, 137, 138, 139, 140, 141, 142, _
                                143, 144, 145, 146, 147, 148, 149, 150, 151, 152)
        
    ' ������ ������� ���� � ����� "���"
    path_to_apc = "\\dixy.local\Departments-HQ\ORPT\��������\�����\_�����_��������_���.xlsb"
    
    Set wb_apc = Workbooks.Open(path_to_apc, ReadOnly:=True)
    Set ws_main = ThisWorkbook.Worksheets("������")
    
    ' �������� ������ � ����������:
    '   - ���� "����� (��������) | ��."
    '   - ���� "������������ | ��."
    
    ws_main.Range(ws_main.Cells(5, CInt(arr_week_column(0))), _
                  ws_main.Cells(find_last_row(ws_main), CInt(arr_week_column(UBound(arr_week_column))))).ClearContents
    
    ws_main.Range(ws_main.Cells(5, CInt(arr_duration_column(0))), _
                  ws_main.Cells(find_last_row(ws_main), CInt(arr_duration_column(UBound(arr_duration_column))))).ClearContents
    
    Dim range_rcid As Range ' �������� ������ ��� ������� ����!������ > ������ 1 (�5:�..)
    Dim ar_rcid As Variant ' ��������������� ������ �� ����!������ > ������ 1 (�5:�..)
    Dim arr_rcid As Variant ' �������� ������ �� ����!������ > ������ 1 (�5:�..)
    
    Set range_rcid = ws_main.Range(ws_main.Cells(5, 1), _
                                   ws_main.Cells(find_last_row(ws_main), 1))
    
    ar_rcid = range_rcid.Value
    ReDim arr_rcid(1 To UBound(ar_rcid, 1))
    For i = LBound(ar_rcid) To UBound(ar_rcid)
        arr_rcid(i) = ar_rcid(i, 1)
    Next i
    Erase ar_rcid
    
    Dim week_number_range As Range ' �������� ������ ��� ������� ����!������ > ���� "����� �������� | ��."
    Dim arr_week_number As Variant ' ������ �� ����!������ > ���� "������������ | ��."
    Set week_number_range = ws_main.Range(ws_main.Cells(4, CInt(arr_week_column(0))), _
                                          ws_main.Cells(4, CInt(arr_week_column(UBound(arr_week_column)))))
    
    arr_week_number = convert_row_to_array_one_demension(week_number_range)
    
    For i = LBound(arr_week_number) To UBound(arr_week_number)
        arr_week_number(i) = CStr(arr_week_number(i))
    Next i
    
    For Each ws_apc In wb_apc.Sheets
'        Debug.Print "������������ ����: " + ws_apc.Name
        
        If InStr(1, ws_apc.Name, "_���", vbTextCompare) > 0 Then
            what_find = CStr(Replace(ws_apc.Name, "_���", ""))
        Else
            what_find = ws_apc.Name
        End If

        Dim indx As Long
        indx = find_index(what_find, arr_week_number)
        
        If indx <> -1 Then
'            Debug.Print "[ + ] ������: " + CStr(indx)
'            Debug.Print " "
            
            Dim arr_ws_apc() As Variant
            
            last_row = find_last_row(ws_apc)
            last_column = find_last_column(ws_apc)
            
            arr_ws_apc = ws_apc.Range(ws_apc.Cells(2, 1), ws_apc.Cells(last_row, last_column)).Value
            
            For Each rcid In arr_rcid ' ���� �� ������� ������!�5�.. (������ 1)
'                Debug.Print "[ -> ] RCID: " + rcid
                
                For i = LBound(arr_ws_apc, 1) To UBound(arr_ws_apc, 1)
                    
                    If rcid = arr_ws_apc(i, 5) Then
'                        Debug.Print "[ + ] RCID: " + CStr(arr_ws_apc(i, 5))
                        
                        ' + 4 (��� ������ ������ � 4�� � ���������)
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
'            Debug.Print "[ x ] � ������� �����������"
'            Debug.Print " "
        End If
        
    Next ws_apc
    
    Call close_apc_file
    Call turn_on_functionalities
    Call sort_sheet
    
    MsgBox "���������� ���������", Title:="[ ! ]"
    
End Sub

