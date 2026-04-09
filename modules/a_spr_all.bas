Attribute VB_Name = "a_spr_all"
Option Explicit

Public Sub load_mjj_all()
    On Error Resume Next

    ReDim mj(1 To 10, 1 To 1)

    mj(1, 1) = "Андреев Д.В."
    mj(2, 1) = "Кузнецов Д.В."
    mj(3, 1) = "Сидоров И.О."
    mj(4, 1) = "Брызгалин Р.Б."
    mj(5, 1) = "Тимошевич Ю.П."
    mj(6, 1) = "Беликов А.С."
    mj(7, 1) = "Забродин А.С."
    mj(8, 1) = "Черкашенинов А.Н."
    mj(9, 1) = "Анискин С.В."
    mj(10, 1) = "Ищенко Е.А."

End Sub

Public Sub load_zkz_all()
    On Error GoTo fallback_data

    Dim ws As Worksheet
    Dim r As Long
    Dim cnt As Long
    Dim lastRow As Long

    Set ws = get_spr_sheet(True)
    If ws Is Nothing Then GoTo fallback_data

    lastRow = ws.Cells(ws.Rows.Count, bzZkz).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, bzZkz).value)) <> "" Then cnt = cnt + 1
    Next r

    If cnt = 0 Then GoTo fallback_data

    ReDim zkz(1 To cnt, 1 To 1)
    cnt = 0

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, bzZkz).value)) <> "" Then
            cnt = cnt + 1
            zkz(cnt, 1) = Trim(CStr(ws.Cells(r, bzZkz).value))
        End If
    Next r

    Exit Sub

fallback_data:
    ReDim zkz(1 To 7, 1 To 1)

    zkz(1, 1) = "ООО АГРОФИРМА «ПЯТИГОРЬЕ»"
    zkz(2, 1) = "ООО ГК «АЛЬФА-СПК-ДЖИТЕЙЧ»"
    zkz(3, 1) = "ООО Компания «Карат»"
    zkz(4, 1) = "ООО Контракт - Авто"
    zkz(5, 1) = "ООО НПФ «ТРЭКОЛ"""
    zkz(6, 1) = "ООО ТСК «ВОСТОК - СТРОЙМАРКЕТ»"
    zkz(7, 1) = "ООО Фирма «ТеплоЦель»"
End Sub

Public Sub load_zkz_contacts_all()
    On Error GoTo no_data

    Dim ws As Worksheet
    Dim r As Long
    Dim cnt As Long
    Dim lastRow As Long

    Set ws = get_spr_sheet(True)
    If ws Is Nothing Then GoTo no_data

    lastRow = ws.Cells(ws.Rows.Count, bzZkz).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, bzZkz).value)) <> "" Then cnt = cnt + 1
    Next r

    If cnt = 0 Then GoTo no_data

    ReDim zkz(1 To cnt, 1 To 1)
    ReDim adr(1 To cnt, 1 To 1)
    ReDim tlf(1 To cnt, 1 To 1)

    cnt = 0
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, bzZkz).value)) <> "" Then
            cnt = cnt + 1
            zkz(cnt, 1) = Trim(CStr(ws.Cells(r, bzZkz).value))
            adr(cnt, 1) = Trim(CStr(ws.Cells(r, bzAdr).value))
            tlf(cnt, 1) = Trim(CStr(ws.Cells(r, bzTlf).value))
        End If
    Next r

    Exit Sub

no_data:
    ReDim zkz(1 To 1, 1 To 1)
    ReDim adr(1 To 1, 1 To 1)
    ReDim tlf(1 To 1, 1 To 1)

    zkz(1, 1) = ""
    adr(1, 1) = ""
    tlf(1, 1) = ""
End Sub

Public Function get_spr_sheet(ByVal forCounterparty As Boolean) As Worksheet
    On Error Resume Next

    If forCounterparty Then
        Set get_spr_sheet = find_sheet(Array("Справочник_контрагентов", "Контрагенты", "База_заказчиков", "Справочник", "spr"))
    Else
        Set get_spr_sheet = find_sheet(Array("Справочник_поставщиков", "Поставщики", "База_поставщиков", "Справочник", "spr"))
    End If

    If get_spr_sheet Is Nothing Then
        Set get_spr_sheet = find_sheet(Array("Контрагенты", "База_заказчиков", "Поставщики", "Справочник", "spr"))
    End If
End Function

Private Function find_sheet(ByVal names As Variant) As Worksheet
    On Error Resume Next

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Set find_sheet = ThisWorkbook.Sheets(CStr(names(i)))
        If Not find_sheet Is Nothing Then Exit Function
    Next i
End Function

Public Sub load_doc_all()
    On Error Resume Next

    ReDim doc(1 To 4, 1 To 1)

    doc(1, 1) = "счет"
    doc(2, 1) = "счет-фактура"
    doc(3, 1) = "накладная"
    doc(4, 1) = "тов-трансп.наклад"

End Sub
