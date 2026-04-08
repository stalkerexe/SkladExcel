Attribute VB_Name = "справочники_сохранение"
Option Explicit

Public Function save_psv_to_spr(ByVal psvName As String) As Boolean
    On Error GoTo err_h

    Dim ws As Worksheet
    Dim newRow As Long

    psvName = Trim(psvName)
    If psvName = "" Then Err.Raise 1001, "save_psv_to_spr", "Не заполнено название поставщика."

    Set ws = get_spr_sheet(False)
    If ws Is Nothing Then Err.Raise 1002, "save_psv_to_spr", "Не найден лист справочника поставщиков."

    If has_duplicate(ws, bzPsv, psvName) Then Err.Raise 1003, "save_psv_to_spr", "Поставщик уже существует в справочнике."

    newRow = ws.Cells(ws.Rows.Count, bzPsv).End(xlUp).Row + 1
    If newRow < 2 Then newRow = 2

    ws.Cells(newRow, bzPsv).Value = psvName

    save_psv_to_spr = True
    Exit Function

err_h:
    save_psv_to_spr = False
    show_save_error "Поставщик", Err.Description
End Function

Public Function save_zkz_to_spr(ByVal zkzName As String, ByVal zkzAdr As String, ByVal zkzTlf As String, ByVal zkzMail As String) As Boolean
    On Error GoTo err_h

    Dim ws As Worksheet
    Dim newRow As Long

    zkzName = Trim(zkzName)
    zkzAdr = Trim(zkzAdr)
    zkzTlf = Trim(zkzTlf)
    zkzMail = Trim(zkzMail)

    If zkzName = "" Then Err.Raise 1101, "save_zkz_to_spr", "Не заполнено название контрагента."

    If ThisWorkbook.Sheets("setting").Range("b41").Value = 1 And zkzTlf = "" Then
        Err.Raise 1102, "save_zkz_to_spr", "Не заполнен обязательный телефон контрагента."
    End If

    Set ws = get_spr_sheet(True)
    If ws Is Nothing Then Err.Raise 1103, "save_zkz_to_spr", "Не найден лист справочника контрагентов."

    If has_duplicate(ws, bzZkz, zkzName) Then Err.Raise 1104, "save_zkz_to_spr", "Контрагент уже существует в справочнике."

    newRow = ws.Cells(ws.Rows.Count, bzZkz).End(xlUp).Row + 1
    If newRow < 2 Then newRow = 2

    ws.Cells(newRow, bzZkz).Value = zkzName
    ws.Cells(newRow, bzAdr).Value = zkzAdr
    ws.Cells(newRow, bzTlf).NumberFormat = "@"
    ws.Cells(newRow, bzTlf).Value = zkzTlf
    ws.Cells(newRow, bzMail).Value = zkzMail

    save_zkz_to_spr = True
    Exit Function

err_h:
    save_zkz_to_spr = False
    show_save_error "Контрагент", Err.Description
End Function

Public Sub refresh_vvod_forms_sources()
    On Error Resume Next

    If vvodPr.Visible Then vvodPr.refresh_psv_source
    If vvodZv.Visible Then vvodZv.refresh_zkz_source
End Sub

Public Sub show_save_error(ByVal entityName As String, ByVal detail As String)
    MsgBox "Ошибка сохранения [" & entityName & "]: " & detail, vbExclamation, "Справочник"
End Sub

Private Function has_duplicate(ByVal ws As Worksheet, ByVal col As Long, ByVal valueToFind As String) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim v As String

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For r = 2 To lastRow
        v = Trim(CStr(ws.Cells(r, col).Value))
        If UCase(v) = UCase(valueToFind) Then
            has_duplicate = True
            Exit Function
        End If
    Next r
End Function
