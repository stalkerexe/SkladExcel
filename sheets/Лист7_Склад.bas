' Component: Лист7  [Склад]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit

Dim r As Long
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error Resume Next
If Target.Count > 1 Then Exit Sub
If Target.Row < 5 Then Exit Sub
r = Cells(Rows.Count, skNm).End(xlUp).Row
i = Target.Row
If Not Intersect(Target, Range(Cells(5, skNm), Cells(r, skComm))) Is Nothing Then
Cancel = True
If Cells(i, skGr) = "" Then
Rows(i).Select
Call add_box
End If
End If
End Sub
Private Sub Worksheet_Deactivate()
On Error Resume Next
Unload frm_Show
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Call unload_mn_mn
Unload frm_sk
End Sub
