Attribute VB_Name = "справочники_crud"
Option Explicit

Private Enum eDictAction
    daView = 1
    daAdd = 2
    daEdit = 3
    daDelete = 4
End Enum

Private Const DICT_HEADER_ROW As Long = 1
Private Const DICT_DATA_START_ROW As Long = 2
Private Const DICT_START_COL As Long = 2

Public Sub open_dict_counterparty()
    open_dict_workflow "Контрагенты", Array("Контрагент", "Адрес", "Телефон", "Email")
End Sub

Public Sub open_dict_supplier()
    open_dict_workflow "Поставщики", Array("Поставщик", "Адрес", "Телефон", "Email")
End Sub

Public Sub open_dict_manager()
    open_dict_workflow "Менеджеры", Array("ФИО", "Телефон", "Email", "Комментарий")
End Sub

Public Sub open_dict_warehouse()
    open_dict_workflow "Склады", Array("Склад", "Адрес", "Ответственный", "Комментарий")
End Sub

Public Sub open_dict_doc_types()
    open_dict_workflow "Типы_документов", Array("Тип документа", "Префикс", "Комментарий")
End Sub

Public Sub open_dict_units()
    open_dict_workflow "ЕдИзм", Array("Ед. изм.", "Кратко", "Комментарий")
End Sub

Private Sub open_dict_workflow(ByVal dictSheetName As String, ByVal headers As Variant)
    On Error GoTo errHandler

    Dim ws As Worksheet
    Set ws = ensure_dict_sheet(dictSheetName, headers)

    ws.Activate

    Dim actionNo As Long
    actionNo = ask_dict_action(dictSheetName)

    Select Case actionNo
        Case daView
            ws.Activate
        Case daAdd
            dict_add_row ws
        Case daEdit
            dict_edit_row ws
        Case daDelete
            dict_delete_row ws
    End Select

    Exit Sub

errHandler:
    MsgBox "Ошибка открытия справочника '" & dictSheetName & "': " & Err.Description, vbExclamation, "Справочники"
End Sub

Private Function ensure_dict_sheet(ByVal dictSheetName As String, ByVal headers As Variant) As Worksheet
    Set ensure_dict_sheet = resolve_dict_sheet(dictSheetName)

    If ensure_dict_sheet Is Nothing Then
        Set ensure_dict_sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ensure_dict_sheet.Name = dictSheetName
        MsgBox "Создан новый справочник: " & dictSheetName, vbInformation, "Справочники"
    End If

    normalize_dict_columns ensure_dict_sheet, headers
    If Trim(CStr(ensure_dict_sheet.Cells(DICT_HEADER_ROW, DICT_START_COL).Value)) = "" Then
        init_headers ensure_dict_sheet, headers
    End If
End Function

Private Function resolve_dict_sheet(ByVal dictSheetName As String) As Worksheet
    On Error Resume Next

    Select Case dictSheetName
        Case "Контрагенты"
            Set resolve_dict_sheet = get_spr_sheet(True)
        Case "Поставщики"
            Set resolve_dict_sheet = get_spr_sheet(False)
        Case Else
            Set resolve_dict_sheet = ThisWorkbook.Worksheets(dictSheetName)
    End Select
End Function

Private Sub normalize_dict_columns(ByVal ws As Worksheet, ByVal headers As Variant)
    If Trim(CStr(ws.Cells(DICT_HEADER_ROW, DICT_START_COL).Value)) <> "" Then Exit Sub
    If Trim(CStr(ws.Cells(DICT_HEADER_ROW, DICT_START_COL - 1).Value)) = "" Then Exit Sub
    If Trim$(CStr(ws.Cells(DICT_HEADER_ROW, DICT_START_COL - 1).Value)) <> CStr(headers(LBound(headers))) Then Exit Sub

    ws.Columns(DICT_START_COL - 1).Insert Shift:=xlToRight
End Sub

Private Sub init_headers(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(DICT_HEADER_ROW, i - LBound(headers) + DICT_START_COL).Value = CStr(headers(i))
        ws.Cells(DICT_HEADER_ROW, i - LBound(headers) + DICT_START_COL).Font.Bold = True
        ws.Columns(i - LBound(headers) + DICT_START_COL).ColumnWidth = 24
    Next i
    ws.Rows(DICT_HEADER_ROW).Interior.Color = RGB(221, 235, 247)
End Sub

Private Function ask_dict_action(ByVal dictName As String) As Long
    Dim prompt As String
    Dim response As String

    prompt = "Справочник: " & dictName & vbCrLf & _
             "1 - Просмотр" & vbCrLf & _
             "2 - Добавить" & vbCrLf & _
             "3 - Редактировать" & vbCrLf & _
             "4 - Удалить" & vbCrLf & vbCrLf & _
             "Оставьте пустым для простого просмотра."

    response = Trim$(InputBox(prompt, "Справочники", ""))
    If response = "" Then
        ask_dict_action = daView
        Exit Function
    End If

    If IsNumeric(response) Then
        ask_dict_action = CLng(response)
        If ask_dict_action < daView Or ask_dict_action > daDelete Then ask_dict_action = daView
    Else
        ask_dict_action = daView
    End If
End Function

Private Sub dict_add_row(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, DICT_START_COL).End(xlUp).Row
    If lastRow < DICT_HEADER_ROW Then lastRow = DICT_HEADER_ROW

    Dim colCount As Long
    colCount = ws.Cells(DICT_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = DICT_START_COL To colCount
        ws.Cells(lastRow + 1, c).Value = InputBox("Введите значение для поля '" & ws.Cells(DICT_HEADER_ROW, c).Value & "'", "Добавление")
    Next c

    ws.Activate
    ws.Cells(lastRow + 1, DICT_START_COL).Select
End Sub

Private Sub dict_edit_row(ByVal ws As Worksheet)
    Dim rowNo As Long
    rowNo = CLng(Val(InputBox("Введите номер строки для редактирования", "Редактирование", CStr(DICT_DATA_START_ROW))))
    If rowNo < DICT_DATA_START_ROW Then Exit Sub

    Dim colCount As Long
    colCount = ws.Cells(DICT_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = DICT_START_COL To colCount
        ws.Cells(rowNo, c).Value = InputBox("Новое значение для поля '" & ws.Cells(DICT_HEADER_ROW, c).Value & "'", "Редактирование", CStr(ws.Cells(rowNo, c).Value))
    Next c

    ws.Activate
    ws.Cells(rowNo, DICT_START_COL).Select
End Sub

Private Sub dict_delete_row(ByVal ws As Worksheet)
    Dim rowNo As Long
    rowNo = CLng(Val(InputBox("Введите номер строки для удаления", "Удаление", CStr(DICT_DATA_START_ROW))))
    If rowNo < DICT_DATA_START_ROW Then Exit Sub

    If MsgBox("Удалить строку " & rowNo & "?", vbYesNo + vbQuestion, "Удаление") = vbYes Then
        ws.Rows(rowNo).Delete Shift:=xlUp
    End If

    ws.Activate
End Sub
