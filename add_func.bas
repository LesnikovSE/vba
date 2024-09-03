Attribute VB_Name = "add_func"

' status: work
' ���������� ���������� ����������(�����������) ������

Sub install_comment_style(cell As Range)
    
    With cell.comment.Shape
        .TextFrame.AutoSize = True
        '.AutoShapeType = msoShapeRoundedRectangle '����������� ����� ����������� ���������
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.1
        .Line.Visible = msoTrue ' ���������� ������� �����������
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' ������ ���� ������� (������)
        .Line.Weight = 0.1 ' ������ ������� �������
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .TextFrame.MarginLeft = 1000   ' ������ �����
        .TextFrame.MarginRight = 1000  ' ������ ������
        .TextFrame.MarginTop = 5000  ' ������ ������
        .TextFrame.MarginBottom = 5000 ' ������ �����
        
    End With
    
End Sub

' status: work
' ������ ����� �������, ���� �����

Function column_name(ByVal column_number As Long) As String
    If column_number <= 0 Then
        column_name = "error"
    Else
        column_name = Left(Cells(1, column_number).Address(False, False), Len(Cells(1, column_number).Address(False, False)) - 1)
    End If
    
End Function

' status: work
' ����� ����������� ������ �� ���!����-���|��
'   ���������� ������ �� ������� "����������" � ��������!����-���|��

Sub enable_filter_on_manager(ByRef ws As Worksheet)
    
    If ws.FilterMode Then
        ws.ShowAllData ' �������� ��� �������� �������
    End If
    
    If ws.AutoFilterMode = False Then
        ws.Range("1:1").AutoFilter Field:=33, Criteria1:=Application.UserName
    End If

End Sub

' status: work
' (����-���|��)
' ��������� ������ �� ��������: �������� ����� (2), ���������� (33), �������� �� (35)

Sub sort_sctd(ByRef ws As Worksheet)

    Dim sort_range As Range
    Set sort_range = ws.Range("A1").CurrentRegion
    
    With ws.Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(2), Order:=xlAscending ' �������� ���������� ������� %
        .SortFields.Add Key:=ws.Columns(6), Order:=xlAscending ' ����� �
        .SortFields.Add Key:=ws.Columns(33) ' �������� (����������)
        .SortFields.Add Key:=ws.Columns(35) ' ������������ ��
        .SetRange sort_range ' �������� �������� ��������
        .Apply ' ������������ �������
    End With

End Sub

' status: work
' ������� ������ �������� ��������� ���������� ������ �� �����

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
' ��������� ��������� � ��������� �����������
' �����������: close_custom_message()

'Sub show_custom_msgbox(ByVal text_caption As String)
'    Dim msgbox_form As Object
'    Set msgbox_form = CreateObject("Forms.Form")
'
'    With msgbox_form
'        .Width = 300
'        .Height = 150
'        .Caption = "Update"
'        .BorderStyle = 1 ' ������������� ������
'
'        ' ����� ���������
'        Dim label As Object
'        Set label = .Controls.Add("Forms.Label.1", "MsgLabel")
'        With label
'            .Caption = text_caption
'            .Left = 10
'            .Top = 10
'            .Width = 280
'            .Height = 80
'            .TextAlign = 2 ' ������������ ������ �� ������
'        End With
'
'        ' ������ "OK"
'        Dim okButton As Object
'        Set okButton = .Controls.Add("Forms.CommandButton.1", "OkButton")
'        With okButton
'            .Caption = "OK"
'            .Width = 100
'            .Height = 30
'            ' ������������ ��������� ������ �� �����������
'            .Left = (.Parent.Width - .Width) / 2
'            ' ������������ ��������� ������ �� ���������
'            .Top = (.Parent.Height - .Height) / 2
'            .Default = True ' ������� ������ �� ��������� (��� ������� ������� Enter)
'        End With
'
'        ' ���������� ������� ��� ������ "OK"
'        msgbox_form.Controls("OkButton").OnAction = "close_custom_message_box"
'
'        .Show
'    End With
'
'End Sub

' status: debug
' ������������ �: show_custom_msgbox()

'Sub close_custom_message_box()
''    Unload Me
'End Sub

