﻿Attribute VB_Name = "a_demo"
Option Explicit

Private Const SKLAD_SHEET As String = "my_set"
Private Const SKLAD_COLUMN As Long = 27 'Колонка AA
Private Const SKLAD_FIRST_ROW As Long = 2

Public Sub open_sklad()
On Error GoTo ErrHandler

If Not EnsureSkladSelected(Form_sklads.ListBox1, "open_sklad") Then Exit Sub

sSk = CStr(Form_sklads.ListBox1.Value)
Unload Form_sklads
DoEvents
Call sklad_show
Exit Sub

ErrHandler:
ReportVbaError "open_sklad", Err.Number, Err.Description, "Склад"
End Sub

Public Sub добавить_склад()
On Error GoTo ErrHandler

Dim newName As String
newName = NormalizeSkName(InputBox("Введите название нового склада:", "Добавить склад"))
If newName = "" Then Exit Sub

Call load_sk
If SkladExists(newName) Then
    MsgBox "Склад с таким названием уже существует.", 48, "Добавить склад"
    Exit Sub
End If

AppendSkladToStore newName
Call load_sk
RefreshSkladListUI newName
Exit Sub

ErrHandler:
ReportVbaError "добавить_склад", Err.Number, Err.Description, "Склад"
End Sub

Public Sub rename_sk()
On Error GoTo ErrHandler

If Not EnsureSkladSelected(Form_sklads.ListBox1, "rename_sk") Then Exit Sub

Dim oldName As String
Dim newName As String

oldName = NormalizeSkName(CStr(Form_sklads.ListBox1.Value))
newName = NormalizeSkName(InputBox("Новое имя склада:", "Переименовать склад", oldName))

If newName = "" Then Exit Sub
If StrComp(oldName, newName, vbTextCompare) = 0 Then Exit Sub

Call load_sk
If SkladExists(newName) Then
    MsgBox "Склад с таким названием уже существует.", 48, "Переименовать склад"
    Exit Sub
End If

If Not UpdateSkladNameInStore(oldName, newName) Then
    MsgBox "Склад не найден в справочнике.", 48, "Переименовать склад"
    Exit Sub
End If

ReplaceWarehouseInDocs oldName, newName

If StrComp(sSk, oldName, vbTextCompare) = 0 Then sSk = newName

Call load_sk
RefreshSkladListUI newName
MsgBox "Склад переименован.", 64, "Склад"
Exit Sub

ErrHandler:
ReportVbaError "rename_sk", Err.Number, Err.Description, "Склад"
End Sub

Public Sub delete_sk()
On Error GoTo ErrHandler

If Not EnsureSkladSelected(Form_sklads.ListBox1, "delete_sk") Then Exit Sub

Dim oldName As String
oldName = NormalizeSkName(CStr(Form_sklads.ListBox1.Value))

Dim movesCount As Long
movesCount = CountWarehouseMoves(oldName)

If movesCount > 0 Then
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "По складу найдено движений: " & movesCount & "." & vbCrLf & vbCrLf & _
        "Да — мигрировать движения на другой склад и удалить." & vbCrLf & _
        "Нет — запретить удаление." & vbCrLf & _
        "Отмена — выйти.", _
        vbYesNoCancel + vbQuestion, "Удаление склада")

    If answer = vbCancel Then Exit Sub
    If answer = vbNo Then
        MsgBox "Удаление запрещено: у склада есть движения.", 48, "Удаление склада"
        Exit Sub
    End If

    Dim targetName As String
    targetName = AskMigrationTarget(oldName)
    If targetName = "" Then Exit Sub

    ReplaceWarehouseInDocs oldName, targetName
    If StrComp(sSk, oldName, vbTextCompare) = 0 Then sSk = targetName
End If

If Not DeleteSkladFromStore(oldName) Then
    MsgBox "Склад не найден в справочнике.", 48, "Удаление склада"
    Exit Sub
End If

Call load_sk
RefreshSkladListUI ""
MsgBox "Склад удалён.", 64, "Склад"
Exit Sub

ErrHandler:
ReportVbaError "delete_sk", Err.Number, Err.Description, "Склад"
End Sub

Private Sub RefreshSkladListUI(ByVal selectedName As String)
On Error Resume Next
Call RefreshWarehouseSelectors(selectedName)
End Sub

Private Function EnsureSkladSelected(ByVal lb As Object, ByVal procName As String) As Boolean
If lb Is Nothing Then
    MsgBox procName & ": список складов недоступен.", vbCritical, "Склад"
    Exit Function
End If

If lb.ListIndex = -1 Then
    MsgBox "Выберите склад!", 64, "Склад"
    Exit Function
End If

EnsureSkladSelected = True
End Function

Private Function NormalizeSkName(ByVal value As String) As String
NormalizeSkName = Trim(Replace(Replace(value, vbCr, " "), vbLf, " "))
End Function

Private Function SkladExists(ByVal nameToFind As String) As Boolean
Dim i As Long
For i = 0 To dic_sk.Count - 1
    If StrComp(CStr(dic_sk.Item(i)), nameToFind, vbTextCompare) = 0 Then
        SkladExists = True
        Exit Function
    End If
Next
End Function

Private Sub AppendSkladToStore(ByVal skladName As String)
Dim ws As Worksheet
If Not RequireSheet(SKLAD_SHEET, ws, "AppendSkladToStore") Then Exit Sub

With ws
    Dim lastRow As Long
    lastRow = .Cells(.Rows.Count, SKLAD_COLUMN).End(xlUp).Row
    If lastRow < SKLAD_FIRST_ROW Then lastRow = SKLAD_FIRST_ROW - 1

    .Cells(lastRow + 1, SKLAD_COLUMN).Value = skladName
End With
End Sub

Private Function UpdateSkladNameInStore(ByVal oldName As String, ByVal newName As String) As Boolean
Dim ws As Worksheet
If Not RequireSheet(SKLAD_SHEET, ws, "UpdateSkladNameInStore") Then Exit Function

With ws
    Dim lastRow As Long
    Dim i As Long

    lastRow = .Cells(.Rows.Count, SKLAD_COLUMN).End(xlUp).Row
    If lastRow < SKLAD_FIRST_ROW Then Exit Function

    For i = SKLAD_FIRST_ROW To lastRow
        If StrComp(NormalizeSkName(CStr(.Cells(i, SKLAD_COLUMN).Value)), oldName, vbTextCompare) = 0 Then
            .Cells(i, SKLAD_COLUMN).Value = newName
            UpdateSkladNameInStore = True
            Exit Function
        End If
    Next
End With
End Function

Private Function DeleteSkladFromStore(ByVal oldName As String) As Boolean
Dim ws As Worksheet
If Not RequireSheet(SKLAD_SHEET, ws, "DeleteSkladFromStore") Then Exit Function

With ws
    Dim lastRow As Long
    Dim i As Long

    lastRow = .Cells(.Rows.Count, SKLAD_COLUMN).End(xlUp).Row
    If lastRow < SKLAD_FIRST_ROW Then Exit Function

    For i = SKLAD_FIRST_ROW To lastRow
        If StrComp(NormalizeSkName(CStr(.Cells(i, SKLAD_COLUMN).Value)), oldName, vbTextCompare) = 0 Then
            .Range(.Cells(i, SKLAD_COLUMN), .Cells(lastRow, SKLAD_COLUMN)).Delete Shift:=xlUp
            DeleteSkladFromStore = True
            Exit Function
        End If
    Next
End With
End Function

Private Function CountWarehouseMoves(ByVal skName As String) As Long
CountWarehouseMoves = 0
CountWarehouseMoves = CountMatchesInSheetColumn("Расход", zvSk, skName)
CountWarehouseMoves = CountWarehouseMoves + CountMatchesInSheetColumn("Приход", prSk, skName)
CountWarehouseMoves = CountWarehouseMoves + CountMatchesInSheetColumn("arh_zkk", arhSk, skName)
CountWarehouseMoves = CountWarehouseMoves + CountMatchesInSheetColumn("arh_prr", arhSk, skName)
CountWarehouseMoves = CountWarehouseMoves + CountMatchesInSheetColumn("arh_vzz", arhSk, skName)
End Function

Private Function CountMatchesInSheetColumn(ByVal sheetName As String, ByVal colIndex As Long, ByVal skName As String) As Long
On Error GoTo ErrHandler

Dim ws As Worksheet
If Not RequireSheet(sheetName, ws, "CountMatchesInSheetColumn") Then Exit Function

With ws
    Dim lastRow As Long
    Dim i As Long
    lastRow = .Cells(.Rows.Count, colIndex).End(xlUp).Row
    For i = 1 To lastRow
        If StrComp(NormalizeSkName(CStr(.Cells(i, colIndex).Value)), skName, vbTextCompare) = 0 Then
            CountMatchesInSheetColumn = CountMatchesInSheetColumn + 1
        End If
    Next
End With

Exit Function
ErrHandler:
ReportVbaError "CountMatchesInSheetColumn", Err.Number, Err.Description, "Склад"
End Function

Private Sub ReplaceWarehouseInDocs(ByVal oldName As String, ByVal newName As String)
ReplaceWarehouseInSheetColumn "Расход", zvSk, oldName, newName
ReplaceWarehouseInSheetColumn "Приход", prSk, oldName, newName
ReplaceWarehouseInSheetColumn "arh_zkk", arhSk, oldName, newName
ReplaceWarehouseInSheetColumn "arh_prr", arhSk, oldName, newName
ReplaceWarehouseInSheetColumn "arh_vzz", arhSk, oldName, newName
End Sub

Private Sub ReplaceWarehouseInSheetColumn(ByVal sheetName As String, ByVal colIndex As Long, ByVal oldName As String, ByVal newName As String)
On Error GoTo ErrHandler

Dim ws As Worksheet
If Not RequireSheet(sheetName, ws, "ReplaceWarehouseInSheetColumn") Then Exit Sub

With ws
    Dim lastRow As Long
    Dim i As Long
    lastRow = .Cells(.Rows.Count, colIndex).End(xlUp).Row
    For i = 1 To lastRow
        If StrComp(NormalizeSkName(CStr(.Cells(i, colIndex).Value)), oldName, vbTextCompare) = 0 Then
            .Cells(i, colIndex).Value = newName
        End If
    Next
End With

Exit Sub
ErrHandler:
ReportVbaError "ReplaceWarehouseInSheetColumn", Err.Number, Err.Description, "Склад"
End Sub

Private Function AskMigrationTarget(ByVal oldName As String) As String
Call load_sk

If dic_sk.Count <= 1 Then
    MsgBox "Нет доступного склада для миграции движений.", 48, "Удаление склада"
    Exit Function
End If

Dim msg As String
Dim i As Long
msg = "Введите склад для миграции движений:" & vbCrLf
For i = 0 To dic_sk.Count - 1
    If StrComp(CStr(dic_sk.Item(i)), oldName, vbTextCompare) <> 0 Then
        msg = msg & "- " & CStr(dic_sk.Item(i)) & vbCrLf
    End If
Next

Dim candidate As String
candidate = NormalizeSkName(InputBox(msg, "Миграция движений"))
If candidate = "" Then Exit Function

If StrComp(candidate, oldName, vbTextCompare) = 0 Then
    MsgBox "Нельзя мигрировать на удаляемый склад.", 48, "Удаление склада"
    Exit Function
End If

If Not SkladExists(candidate) Then
    MsgBox "Указанный склад не найден в справочнике.", 48, "Удаление склада"
    Exit Function
End If

AskMigrationTarget = candidate
End Function
