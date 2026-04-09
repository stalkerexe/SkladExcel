Attribute VB_Name = "я_Public"
Option Explicit

Public id() As Variant
Public opl() As Variant
Public skid() As Variant

Public iOpl As String
Public iSkid As Double

Public iSm As Double

Public iFiles As String

Public sSpr As String
Public sSpr_new As String

Public Const rwPr_zkz As Integer = 4
Public Const rwPr_mj As Integer = 5
Public Const rwPr_doc As Integer = 7
Public Const rwPr_dt As Integer = 9

Public Const rwZv_zkz As Integer = 4
Public Const rwZv_mj As Integer = 5
Public Const rwZv_adr As Integer = 6
Public Const rwZv_tlf As Integer = 7
Public Const rwZv_dt As Integer = 8
Public Const rwZv_dt2 As Integer = 9

Public sOsn As String
Public textS As String

Public ctr As Control
Public iCtr As Control

Public iOperation As String
Public iOperation2 As String
Public shNm As String

Public iRz As Integer
Public iWrapText As String
Public iMsh As Double
Public iHg As Double

Public r24 As Long
Public n7 As Long: Public n2 As Long: Public n As Long

Public i As Long: Public ii As Long: Public j As Long

Public row1 As Long: Public row2 As Long
Public r7 As Long: Public r9 As Long
Public iCol As Long
Public ind As Long
Public rw As Long
Public iRow As Long
Public iRow_zk As Long

Public sheetNm As String

Public sDoc As String
Public sDocN As String
Public sDocDt As String

Public sComm As String
Public dic_sk As Variant
Public dic As Variant

Public sPrinter As String
Public sSk As String

Public Const sPass As String = 999
Public Const parol_sheet As String = "Ai987"

Public flag As Byte: Public flag_find As Byte: Public flag_open As Byte: Public flag_hidden As Boolean: Public flag_arh As Byte
Public sGr As String: Public sNm As String: Public sCod As String: Public sEd As String
Public sCol As Double: Public sBr As Double
Public sCnZ As Variant: Public sCnR As Variant: Public sCn As Variant: Public sSm As Variant
Public sID As String
Public sFolder As String
Public sFiles As String
Public vbDirectory As String
Public cm As Integer
Public str_ As String
Public comm() As Variant
Public marker As String
Public nomer As String
Public sZkz As String: Public sMj As String: Public sAdr As String: Public sTlf As String: Public sMail As String
Public sDt As Date: Public sDt2 As Date
Public summ As Double
Public summA As Double

Public Const rwZv As Integer = 13
Public Const rwzvSm As Integer = 9
Public Const iBox1 As Integer = 13
Public Const iBox2 As Integer = 15

' ==========================================
' === КОНСТАНТЫ НОМЕРОВ КОЛОНОК ===
' ==========================================

' === РАСХОД (Лист3) - zv префикс ===
Public Const zvNN As Integer = 1          ' Номер
Public Const zvNm As Integer = 2          ' Наименование
Public Const zvCod As Integer = 3         ' Код товара
Public Const zvEd As Integer = 4          ' Единица измерения
Public Const zvCnZ As Integer = 5         ' Цена закупки
Public Const zvCol As Integer = 6         ' Кол-во
Public Const zvCn As Integer = 7          ' Цена единица
Public Const zvSm As Integer = 8          ' Сумма
Public Const zvOst As Integer = 9         ' Остаток на складе
Public Const zvSk As Integer = 10         ' Скидка
Public Const zvCnR As Integer = 11        ' Цена розница
Public Const zvComm As Integer = 12       ' Комментарий

' === ПРИХОД (Лист4) - pr префикс ===
Public Const prNN As Integer = 1          ' Номер
Public Const prNm As Integer = 2          ' Наименование
Public Const prCod As Integer = 3         ' Код товара
Public Const prEd As Integer = 4          ' Единица измерения
Public Const prCnZ As Integer = 5         ' Цена за единицу
Public Const prCol As Integer = 6         ' Кол-во
Public Const prCn As Integer = 7          ' Цена единица
Public Const prSm As Integer = 8          ' Сумма
Public Const prGr As Integer = 9          ' Группа/Категория
Public Const prCnR As Integer = 10        ' Цена розница
Public Const prOst As Integer = 11        ' Остаток
Public Const prComm As Integer = 12       ' Комментарий
Public Const prDoc As Integer = 13       '
Public Const prDocN As Integer = 14       '
Public Const prDocDt As Integer = 15       '



' === ЗАКАЗЫ (Лист6) - zk префикс ===
Public Const zkNN As Integer = 1          ' Номер
Public Const zkNom As Integer = 2         ' Номер заказа
Public Const zkNm As Integer = 3          ' Наименование
Public Const zkCod As Integer = 4         ' Код товара
Public Const zkEd As Integer = 5          ' Единица измерения
Public Const zkCnZ As Integer = 6         ' Цена за единицу
Public Const zkCol As Integer = 7         ' Кол-во
Public Const zkCn As Integer = 8          ' Цена единица
Public Const zkSm As Integer = 9          ' Сумма
Public Const zkDt1 As Integer = 10        ' Дата заказа
Public Const zkDt2 As Integer = 11        ' Дата отгрузки
Public Const zkOst As Integer = 12        ' Остаток
Public Const zkComm As Integer = 13       ' Комментарий

' === ПРИХОД ЗАКАЗОВ (Лист9) - pzk префикс ===
Public Const pzkNN As Integer = 1         ' Номер
Public Const pzkNom As Integer = 2        ' Номер закупки
Public Const pzkNm As Integer = 3         ' Наименование
Public Const pzkCod As Integer = 4        ' Код товара
Public Const pzkEd As Integer = 5         ' Единица измерения
Public Const pzkCnZ As Integer = 6        ' Цена за единицу
Public Const pzkCol As Integer = 7        ' Кол-во
Public Const pzkCn As Integer = 8         ' Цена единица
Public Const pzkSm As Integer = 9         ' Сумма
Public Const pzkPsv As Integer = 10       ' Поставщик
Public Const pzkDoc As Integer = 11       ' Документ основания
Public Const pzkDocN As Integer = 12      ' Номер документа
Public Const pzkDt As Integer = 13        ' Дата
Public Const pzkOsn As Integer = 14       ' Основание
Public Const pzkComm As Integer = 15      ' Комментарий
Public Const pzkMj As Integer = 16        ' Место хранения
Public Const pzkDocDt As Integer = 17     ' Дата документа

' === СКЛАД (Лист7) - sk префикс ===
Public Const skNN As Integer = 1          ' Номер
Public Const skNm As Integer = 2          ' Наименование
Public Const skCod As Integer = 3         ' Код товара
Public Const skEd As Integer = 4          ' Единица измерения
Public Const skCnZ As Integer = 5         ' Цена за единицу
Public Const skCol As Integer = 6         ' Кол-во на складе
Public Const skCn As Integer = 7          ' Цена единица
Public Const skSm As Integer = 8          ' Сумма
Public Const skGr As Integer = 9          ' Группа/Статус
Public Const skOst As Integer = 10        ' Остаток
Public Const skComm As Integer = 11       ' Комментарий

' === АРХИВ - arh префикс ===
Public Const arhNN As Integer = 1         ' Номер
Public Const arhNom As Integer = 2        ' Номер архива
Public Const arhNm As Integer = 3         ' Наименование
Public Const arhCod As Integer = 4        ' Код товара
Public Const arhEd As Integer = 5         ' Единица измерения
Public Const arhCnZ As Integer = 6        ' Цена закупки
Public Const arhCol As Integer = 7        ' Кол-во
Public Const arhCn As Integer = 8         ' Цена единица
Public Const arhCnR As Integer = 9        ' Цена розница
Public Const arhSm As Integer = 10        ' Сумма
Public Const arhDt As Integer = 11        ' Дата
Public Const arhZkz As Integer = 12       ' Заказчик
Public Const arhMj As Integer = 13        ' Место хранения
Public Const arhDoc As Integer = 14       ' Документ
Public Const arhSmA As Integer = 15       ' Сумма архива
Public Const arhComm As Integer = 16      ' Комментарий
Public Const arhSk As Integer = 17        ' Скидка архива
Public Const arhTlf As Integer = 18        ' Телефон
Public Const arhAdr As Integer = 19        ' Адрес
Public Const arhDt2 As Integer = 20       ' Дата 2

