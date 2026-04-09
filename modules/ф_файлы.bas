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

For i = SKLAD_FIRST_ROW To lastRow
    skName = Trim(CStr(ws.Cells(i, SKLAD_COLUMN).value))
    If skName <> "" Then
        If Not DictionaryContainsValue(dic_sk, skName) Then
            dic_sk.Add dic_sk.Count, skName
        End If
    End If
Next
Exit Sub

fallback:
Set dic_sk = CreateObject("Scripting.Dictionary")
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

Public Sub LoadSkToControl(ByVal ctr As Object, Optional ByVal selectedName As String = "")
On Error Resume Next

If ctr Is Nothing Then Exit Sub

Call load_sk

If selectedName = "" Then selectedName = CStr(ctr.value)
ctr.Clear

Dim i As Long
For i = 0 To dic_sk.Count - 1
    ctr.AddItem dic_sk.Item(i)
Next

If ctr.ListCount = 0 Then Exit Sub

For i = 0 To ctr.ListCount - 1
    If StrComp(CStr(ctr.List(i)), selectedName, vbTextCompare) = 0 Then
        ctr.ListIndex = i
        Exit Sub
    End If
Next

ctr.ListIndex = 0
End Sub

Public Function IsUserFormLoaded(ByVal formName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If StrComp(frm.Name, formName, vbTextCompare) = 0 Then
        IsUserFormLoaded = True
        Exit Function
    End If
Next
End Function

Public Sub RefreshWarehouseSelectors(Optional ByVal selectedName As String = "")
On Error Resume Next

If IsUserFormLoaded("Form_sklads") Then LoadSkToControl Form_sklads.ListBox1, selectedName
If IsUserFormLoaded("frm_sk") Then LoadSkToControl frm_sk.ListBox1, selectedName
If IsUserFormLoaded("frm_ZVK") Then LoadSkToControl frm_ZVK.comb_sk, selectedName
If IsUserFormLoaded("Praise") Then
    LoadSkToControl Praise.comb_sk, selectedName
    LoadSkToControl Praise.comb_sk_set, selectedName
End If
End Sub

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
