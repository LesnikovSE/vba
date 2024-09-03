Attribute VB_Name = "comments"

' ���������� ������������� ����������(�����������) � ����������� � ����� � "������ - �������� �������� � ��."
' status: work

Sub new_comments_to_ordering()

    Call add_func.turn_off_functionalities

    Dim ws As Worksheet
    Dim last_row As Long
    Dim arr_data() As Variant ' ������ ����� ����� ����������� (��� ����������)
    Dim arr_order As Variant ' ������ �� ��������� �����������!������-�������� �������� � ��. (6 ������, ��� ����������)
    Dim comment_text As String ' ��������� ���������� ��� ��������������� �����������
    ' ��� �����
    ' ���� 1 ����� �������� ������� �������
    ' ���� �����: � .. �� ..
    ' ����: � .. �� ..
    ' ����� �� �����: ���. | ��.
    ' ����������� ��������
    ' ���� ��������: ���. | ��.
    Dim arr_week_info() As Variant ' ����, ������ ������� ������ ��� ������������
    ' arr_week_info(i)( * )
    ' (0)  - �������� ��
    ' (1)  - ��� �����
    ' (2)  - ���� ������� ������ �������� ������� �������
    ' (3)  - ���� ������ ��
    ' (4)  - ���� ��������� ��
    ' (5)  - ���� ������ �����
    ' (6)  - ���� ��������� �����
    ' (7)  - ����� � ��.
    ' (8)  - ����������� ��������
    ' (9)  - ���� �������� � ���.
    ' (10) - ���� �������� � ��.
    Dim i As Long ' ���������� ������� �����
    Dim j As Long ' ���������� ������� ��������

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

            ' ������ ������ �� ��������������
            If arr_order(i, j) <> "" And Not IsEmpty(arr_order(i, j)) Then

                ' �������� ������-����������� �� ������ � ������
                comment_text = _
                    " " & arr_data(i, arr_week_info(j - 1)(0)) & vbCrLf & _
                    " �����: " & arr_data(i, arr_week_info(j - 1)(1)) & vbCrLf & vbCrLf & _
                    " 1 �����: " & Format(arr_data(i, arr_week_info(j - 1)(2)), "dd.mm") & vbCrLf & _
                    " ����: � " & Format(arr_data(i, arr_week_info(j - 1)(4)), "dd.mm") & " �� " & Format(arr_data(i, arr_week_info(j - 1)(5)), "dd.mm") & vbCrLf & _
                    " ����: � " & Format(arr_data(i, arr_week_info(j - 1)(3)), "dd.mm") & " �� " & Format(arr_data(i, arr_week_info(j - 1)(4)), "dd.mm") & vbCrLf & _
                    " �����: " & Round(arr_data(i, arr_week_info(j - 1)(7)) / arr_data(i, 190), 1) & " ���. | " & arr_data(i, arr_week_info(j - 1)(7)) & " ��." & vbCrLf & _
                    " ���. ��������: " & arr_data(i, arr_week_info(j - 1)(8)) & vbCrLf & _
                    " ���� ����.: " & arr_data(i, arr_week_info(j - 1)(9)) & " ���. | " & arr_data(i, arr_week_info(j - 1)(10)) & " ��."

                ' ������������� ����������� ��� ������ �� ����� �����������!OM5:OR..
                ' (+ 4)   - ���������� ����� ���������� �� ����� �����������
                ' (+ 402) - ���������� �������� ����� �� ������� �������

                With ws.Cells(i + 4, j + 402)
                    .ClearComments
                    .AddComment Text:=comment_text
                    .comment.Visible = False
                    ' ���������� ��������� ����������(�����������)
                    Call add_func.install_comment_style(ws.Cells(i + 4, j + 402))
                End With

            End If

        Next j ' ��������� ������ � ������ i
    Next i ' ��������� ������ ��������� OM5:OR..

    Call add_func.turn_on_functionalities

    MsgBox "����������� � ""������ - �������� �������� � ��."" �����������", Title:="[ ! ]"

End Sub
