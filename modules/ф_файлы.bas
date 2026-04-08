Attribute VB_Name = "ф_файлы"
Option Explicit

Private Const SKLAD_SHEET As String = "my_set"
Private Const SKLAD_COLUMN As Long = 27 'Колонка AA
Private Const SKLAD_FIRST_ROW As Long = 2

Public Sub load_sk()
On Error GoTo fallback

Set dic_sk = CreateObject("Scripting.Dictionary")

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim skName As String

Set ws = ThisWorkbook.Sheets(SKLAD_SHEET)
lastRow = ws.Cells(ws.Rows.Count, SKLAD_COLUMN).End(xlUp).Row

If lastRow < SKLAD_FIRST_ROW Then
    SeedDefaultWarehouses ws
    lastRow = ws.Cells(ws.Rows.Count, SKLAD_COLUMN).End(xlUp).Row
End If

For i = SKLAD_FIRST_ROW To lastRow
    skName = Trim(CStr(ws.Cells(i, SKLAD_COLUMN).Value))
    If skName <> "" Then
        If Not DictionaryContainsValue(dic_sk, skName) Then
            dic_sk.Add dic_sk.Count, skName
        End If
    End If
Next

If dic_sk.Count = 0 Then SeedDefaultsToDictionary
Exit Sub

fallback:
Set dic_sk = CreateObject("Scripting.Dictionary")
SeedDefaultsToDictionary
End Sub

Private Sub SeedDefaultWarehouses(ByVal ws As Worksheet)
ws.Cells(SKLAD_FIRST_ROW, SKLAD_COLUMN).Value = "Материалы"
ws.Cells(SKLAD_FIRST_ROW + 1, SKLAD_COLUMN).Value = "Металлопрокат"
ws.Cells(SKLAD_FIRST_ROW + 2, SKLAD_COLUMN).Value = "Спецодежда"
End Sub

Private Sub SeedDefaultsToDictionary()
dic_sk.Add dic_sk.Count, "Материалы"
dic_sk.Add dic_sk.Count, "Металлопрокат"
dic_sk.Add dic_sk.Count, "Спецодежда"
End Sub

Private Function DictionaryContainsValue(ByVal d As Object, ByVal valueToFind As String) As Boolean
Dim i As Long
For i = 0 To d.Count - 1
    If StrComp(CStr(d.Item(i)), valueToFind, vbTextCompare) = 0 Then
        DictionaryContainsValue = True
        Exit Function
    End If
Next
End Function

Public Sub sh_frm_Skidka()
        On Error Resume Next
        Call unload_mn_vid
        Unload frm_Oplata
        frm_Skidka.Show
End Sub

Public Sub sh_frm_Oplata()
        On Error Resume Next
        Call unload_mn_vid
        Unload frm_Skidka
        frm_Oplata.Show
End Sub
