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
    za As String ' ЗА
    idKA As String ' Код поставщика
    nameKA As String 'Наименование поставщика
    contract As String ' Номер активного контракта
    ac As String ' Адресная цепочка
    rc As String ' Код РЦ
    idProduct As String ' Код продукта
    nameProduct As String ' Наименование продукта
    lv As Integer ' действующее ЛВ по ЗА
    typeProd As String ' тип товара (Регулярка, IN - OUT, Сезонный, Распродажа, ТМЦ)
    calendar As Variant ' массив с графиком заказов
    deliveryTime As String ' время доставки
    week1 As Variant ' массив с впр по 1 неделе
    week2 As Variant ' массив с впр по 2 неделе
    week3 As Variant ' массив с впр по 3 неделе
    week4 As Variant ' массив с впр по 4 неделе
    week5 As Variant ' массив с впр по 5 неделе
    week6 As Variant ' массив с впр по 6 неделе
    week7 As Variant ' массив с впр по 7 неделе
'    td As Variant ' массив с данными для товар дня (дата, количество к заказу в кор.)
'    sc As Variant ' массив с данными для слип чек (дата, количество к заказу в кор.)
    
    ' [ STATUS: DEBUG]
 
    oos As Variant ' 3 значения: блок ООС (0 - %ООС, 1 - кол_во маг_ов с ООС, 2 - Норматив на ТТ)
    ztrk As Variant ' 4 значения: блок затарка (0 - шт/шоубокс, 1 - шт/кор, 2 - кор/слой, 3 - кор/плт)
    normativ As Variant ' 2 значения: (0 - норматив в днях, 1 - норматив в руб.)
    avrg_otgr_4_week As Variant ' 4 значения6
    otgr_1_week As Variant ' отгрузки по дням за последние 7 дней
    order_magazine As Variant ' заказы магазинов за последние 7 дней
    real_mag_7_days As Variant ' реализация в магазинах за последние 7 дней
    dvij_7_day As String ' ДА/НЕТ - движение товара за последние 7 дней
    
    
    ' остатки по всем ЛВ
    ' оставтки по всем ЛВ + объем в пути
    ' резерв по всем ЛВ
    ' резерв по всем ЛВ по сегодняшним заказам
    
    
    ' товарный запас + объем заказов в пути
    tz_p_real As String ' по реализации
    tz_p_otgr As String ' по отгрузке
    'tz_p_progn As String ' по прогнозу
    
    ' товарный запас без учета объемов заказов в пути
    tz_bp_real As String ' по реализации
    tz_bp_otgr As String ' по отгрузке
    'tz_bp_prog As String ' по прогнозу
    
    'planograma_main As String ' текущая планограмма
    'planograma_mgzn As String ' остатки магазинов на вчера
    
    
    
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
    ' .calendar содержит значения количества дней поставки
    '    из графика заказов (с "пн" по "вс")
    this.calendar = Array(dataArray(1, 212), dataArray(1, 213), _
        dataArray(1, 214), dataArray(1, 215), dataArray(1, 216), _
        dataArray(1, 217), dataArray(1, 218))
    this.deliveryTime = dataArray(1, 219)
    ' из чего состоит "week1..7"
    '  1. тип промо акции
    '  2. дата заказа на предзагрузку промо по указанной неделе
    '  3. дата начала закупочной цены
    '  4. дата окончания закупочной цены
    '  5. дата начала промо акции
    '  6. дата окончания промо акции
    '  7. количество (в шт.) для закупа на промо акцию
    '  8. количество (в шт.) минимальной выкладки товара
    '  9. количество по плану отгузки (в кор.)
    ' 10. количество по плану отгузки (в шт.)
    ' 13. количество которое осталось заказать (в кор.)
    ' 14. количество которое осталось заказать (в шт.)
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
