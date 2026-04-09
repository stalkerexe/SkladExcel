Attribute VB_Name = "справочники_crud"
Option Explicit

Private Enum eDictAction
    daView = 1
    daAdd = 2
    daEdit = 3
    daDelete = 4
End Enum

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
    On Error GoTo ErrHandler

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

ErrHandler:
    MsgBox "Ошибка открытия справочника '" & dictSheetName & "': " & Err.Description, vbExclamation, "Справочники"
End Sub

Private Function ensure_dict_sheet(ByVal dictSheetName As String, ByVal headers As Variant) As Worksheet
    On Error Resume Next
    Set ensure_dict_sheet = ThisWorkbook.Worksheets(dictSheetName)
    On Error GoTo 0

    If ensure_dict_sheet Is Nothing Then
        Set ensure_dict_sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ensure_dict_sheet.Name = dictSheetName
        init_headers ensure_dict_sheet, headers
        MsgBox "Создан новый справочник: " & dictSheetName, vbInformation, "Справочники"
    ElseIf Trim(CStr(ensure_dict_sheet.Cells(1, 1).value)) = "" Then
        init_headers ensure_dict_sheet, headers
    End If
End Function

Private Sub init_headers(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i - LBound(headers) + 1).value = CStr(headers(i))
        ws.Cells(1, i - LBound(headers) + 1).Font.Bold = True
        ws.Columns(i - LBound(headers) + 1).ColumnWidth = 24
    Next i
    ws.Rows(1).Interior.Color = RGB(221, 235, 247)
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    Dim colCount As Long
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To colCount
        ws.Cells(lastRow + 1, c).value = InputBox("Введите значение для поля '" & ws.Cells(1, c).value & "'", "Добавление")
    Next c

    ws.Activate
    ws.Cells(lastRow + 1, 1).Select
End Sub

Private Sub dict_edit_row(ByVal ws As Worksheet)
    Dim rowNo As Long
    rowNo = CLng(Val(InputBox("Введите номер строки для редактирования", "Редактирование", "2")))
    If rowNo < 2 Then Exit Sub

    Dim colCount As Long
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To colCount
        ws.Cells(rowNo, c).value = InputBox("Новое значение для поля '" & ws.Cells(1, c).value & "'", "Редактирование", CStr(ws.Cells(rowNo, c).value))
    Next c

    ws.Activate
    ws.Cells(rowNo, 1).Select
End Sub

Private Sub dict_delete_row(ByVal ws As Worksheet)
    Dim rowNo As Long
    rowNo = CLng(Val(InputBox("Введите номер строки для удаления", "Удаление", "2")))
    If rowNo < 2 Then Exit Sub

    If MsgBox("Удалить строку " & rowNo & "?", vbYesNo + vbQuestion, "Удаление") = vbYes Then
        ws.Rows(rowNo).Delete Shift:=xlUp
    End If

    ws.Activate
End Sub
