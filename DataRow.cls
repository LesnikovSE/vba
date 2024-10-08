VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DataRow class module

' vbCrLf

Private Type DataRow
    ' Define attributes of the data row
    za As String ' ��
    idKA As String ' ��� ����������
    nameKA As String '������������ ����������
    contract As String ' ����� ��������� ���������
    ac As String ' �������� �������
    rc As String ' ��� ��
    idProduct As String ' ��� ��������
    nameProduct As String ' ������������ ��������
    lv As Integer ' ����������� �� �� ��
    typeProd As String ' ��� ������ (���������, IN - OUT, ��������, ����������, ���)
    calendar As Variant ' ������ � �������� �������
    deliveryTime As String ' ����� ��������
    week1 As Variant ' ������ � ��� �� 1 ������
    week2 As Variant ' ������ � ��� �� 2 ������
    week3 As Variant ' ������ � ��� �� 3 ������
    week4 As Variant ' ������ � ��� �� 4 ������
    week5 As Variant ' ������ � ��� �� 5 ������
    week6 As Variant ' ������ � ��� �� 6 ������
    week7 As Variant ' ������ � ��� �� 7 ������
'    td As Variant ' ������ � ������� ��� ����� ��� (����, ���������� � ������ � ���.)
'    sc As Variant ' ������ � ������� ��� ���� ��� (����, ���������� � ������ � ���.)
    
    ' [ STATUS: DEBUG]
 
    oos As Variant ' 3 ��������: ���� ��� (0 - %���, 1 - ���_�� ���_�� � ���, 2 - �������� �� ��)
    ztrk As Variant ' 4 ��������: ���� ������� (0 - ��/�������, 1 - ��/���, 2 - ���/����, 3 - ���/���)
    normativ As Variant ' 2 ��������: (0 - �������� � ����, 1 - �������� � ���.)
    avrg_otgr_4_week As Variant ' 4 ��������6
    otgr_1_week As Variant ' �������� �� ���� �� ��������� 7 ����
    order_magazine As Variant ' ������ ��������� �� ��������� 7 ����
    real_mag_7_days As Variant ' ���������� � ��������� �� ��������� 7 ����
    dvij_7_day As String ' ��/��� - �������� ������ �� ��������� 7 ����
    
    
    ' ������� �� ���� ��
    ' �������� �� ���� �� + ����� � ����
    ' ������ �� ���� ��
    ' ������ �� ���� �� �� ����������� �������
    
    
    ' �������� ����� + ����� ������� � ����
    tz_p_real As String ' �� ����������
    tz_p_otgr As String ' �� ��������
    'tz_p_progn As String ' �� ��������
    
    ' �������� ����� ��� ����� ������� ������� � ����
    tz_bp_real As String ' �� ����������
    tz_bp_otgr As String ' �� ��������
    'tz_bp_prog As String ' �� ��������
    
    'planograma_main As String ' ������� �����������
    'planograma_mgzn As String ' ������� ��������� �� �����
    
    
    
End Type

Private this As DataRow

' Method to initialize a DataRow object from a range of data
Public Sub Initialize(ByVal dataRange As Range)
    
    Dim dataArray As Variant
    ' Transpose the data range into a one-dimensional array
    dataArray = dataRange.Value
    
    ' Assign values to DataRow object attributes
    this.za = dataArray(1, 145)
    this.idKA = dataArray(1, 146)
    this.nameKA = dataArray(1, 147)
    this.contract = dataArray(1, 148)
    this.ac = dataArray(1, 149)
    this.rc = dataArray(1, 150)
    this.idProduct = dataArray(1, 151)
    this.nameProduct = dataArray(1, 152)
    this.lv = dataArray(1, 153)
    this.typeProd = dataArray(1, 154)
    ' .calendar �������� �������� ���������� ���� ��������
    '    �� ������� ������� (� "��" �� "��")
    this.calendar = Array(dataArray(1, 212), dataArray(1, 213), _
        dataArray(1, 214), dataArray(1, 215), dataArray(1, 216), _
        dataArray(1, 217), dataArray(1, 218))
    this.deliveryTime = dataArray(1, 219)
    ' �� ���� ������� "week1..7"
    '  1. ��� ����� �����
    '  2. ���� ������ �� ������������ ����� �� ��������� ������
    '  3. ���� ������ ���������� ����
    '  4. ���� ��������� ���������� ����
    '  5. ���� ������ ����� �����
    '  6. ���� ��������� ����� �����
    '  7. ���������� (� ��.) ��� ������ �� ����� �����
    '  8. ���������� (� ��.) ����������� �������� ������
    '  9. ���������� �� ����� ������� (� ���.)
    ' 10. ���������� �� ����� ������� (� ��.)
    ' 13. ���������� ������� �������� �������� (� ���.)
    ' 14. ���������� ������� �������� �������� (� ��.)
    this.week1 = Array( _
                    dataArray(1, 41), _
                    dataArray(1, 42), _
                    dataArray(1, 43), _
                    dataArray(1, 44), _
                    dataArray(1, 45), _
                    dataArray(1, 46), _
                    dataArray(1, 47), _
                    dataArray(1, 48), _
                    dataArray(1, 49), _
                    dataArray(1, 50), _
                    dataArray(1, 53), _
                    dataArray(1, 54))
    this.week2 = Array(dataArray(1, 55), dataArray(1, 56), dataArray(1, 57), dataArray(1, 58), dataArray(1, 59), dataArray(1, 60), dataArray(1, 61), dataArray(1, 62), dataArray(1, 63), dataArray(1, 64), dataArray(1, 67), dataArray(1, 68))
    this.week3 = Array(dataArray(1, 69), dataArray(1, 70), dataArray(1, 71), dataArray(1, 72), dataArray(1, 73), dataArray(1, 74), dataArray(1, 75), dataArray(1, 76), dataArray(1, 77), dataArray(1, 78), dataArray(1, 81), dataArray(1, 82))
    'current Week 4
    this.week4 = Array(dataArray(1, 83), dataArray(1, 84), dataArray(1, 85), dataArray(1, 86), dataArray(1, 87), dataArray(1, 88), dataArray(1, 89), dataArray(1, 90), dataArray(1, 91), dataArray(1, 92), dataArray(1, 95), dataArray(1, 96))
    this.week5 = Array(dataArray(1, 97), dataArray(1, 98), dataArray(1, 99), dataArray(1, 100), dataArray(1, 101), dataArray(1, 102), dataArray(1, 103), dataArray(1, 104), dataArray(1, 105), dataArray(1, 106), dataArray(1, 109), dataArray(1, 110))
    this.week6 = Array(dataArray(1, 111), dataArray(1, 112), dataArray(1, 113), dataArray(1, 114), dataArray(1, 115), dataArray(1, 116), dataArray(1, 117), dataArray(1, 118), dataArray(1, 119), dataArray(1, 120), dataArray(1, 123), dataArray(1, 124))
    this.week7 = Array(dataArray(1, 125), dataArray(1, 126), dataArray(1, 127), dataArray(1, 128), dataArray(1, 129), dataArray(1, 130), dataArray(1, 131), dataArray(1, 132), dataArray(1, 133), dataArray(1, 134), dataArray(1, 137), dataArray(1, 138))


End Sub

' Property Get method for "za" attribute
Public Property Get za() As String
    za = this.za
End Property

' Property Get method for "idKA" attribute
Public Property Get idKA() As String
    idKA = this.idKA
End Property

' Property Get method for "nameKA" attribute
Public Property Get nameKA() As String
    nameKA = this.nameKA
End Property

' Property Get method for "contract" attribute
Public Property Get contract() As String
    contract = this.contract
End Property

' Property Get method for "ac" attribute
Public Property Get ac() As String
    ac = this.ac
End Property

' Property Get method for "rc" attribute
Public Property Get rc() As String
    rc = this.rc
End Property

' Property Get method for "idProduct" attribute
Public Property Get idProduct() As String
    idProduct = this.idProduct
End Property

' Property Get method for "nameProduct" attribute
Public Property Get nameProduct() As String
    nameProduct = this.nameProduct
End Property

' Property Get method for "lv" attribute
Public Property Get lv() As String
    lv = this.lv
End Property

' Property Get method for "lv" attribute
Public Property Get calendar() As Variant
    calendar = this.calendar
End Property

' Property Get method for "deliveryTime" attribute
Public Property Get deliveryTime() As String
    deliveryTime = this.deliveryTime
End Property

' Property Get method for "week1" attribute
Public Property Get week1() As Variant
    week1 = this.week1
End Property

' Property Get method for "week2" attribute
Public Property Get week2() As Variant
    week2 = this.week2
End Property

' Property Get method for "week3" attribute
Public Property Get week3() As Variant
    week3 = this.week3
End Property

' Property Get method for "week4" attribute
Public Property Get week4() As Variant
    week4 = this.week4
End Property

' Property Get method for "week5" attribute
Public Property Get week5() As Variant
    week5 = this.week5
End Property

' Property Get method for "week6" attribute
Public Property Get week6() As Variant
    week6 = this.week6
End Property

'' Property Get method for "week7" attribute
'Public Property Get week7() As Variant
'    week7 = this.week7
'End Property
