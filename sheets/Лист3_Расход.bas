' Component: Лист3  [Расход]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit

Dim r As Long

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error Resume Next
If Target.Count > 1 Then Exit Sub
If Target.Row < rwZv Then Exit Sub
r = Cells(Rows.Count, zvNm).End(xlUp).Row
i = Target.Row
If Not Intersect(Target, Range(Cells(rwZv, zvNm), Cells(r, zvSm))) Is Nothing Then
Cancel = True
Call Add_MyContextMenu_zv
Application.CommandBars("MyContextMenu_zv").ShowPopup
End If
End Sub

Private Sub Add_MyContextMenu_zv()
On Error Resume Next
Application.CommandBars("MyContextMenu_zv").Delete
With Application.CommandBars.Add(Name:="MyContextMenu_zv", Position:=msoBarPopup, Temporary:=True)
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 21
.Caption = "Удалить позицию"
.OnAction = "del_poz_zv"
End With
End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
Dim ra As Integer
If Target.Count > 1 Then Exit Sub
If Target.Row < rwZv Then Exit Sub
r = Cells(Rows.Count, zvNm).End(xlUp).Row
ra = Target.Row
If Not Intersect(Target, Range(Cells(rwZv, zvCol), Cells(r, zvCnR))) Is Nothing Then
Call проверка_остатков(ra)
End If
End Sub

Private Function проверка_остатков(i As Integer)
On Error Resume Next
If Sheets("setting").Range("i4") = 1 Then
If Cells(i, zvCol) > Cells(i, zvOst) Then
Cells(i, zvCol) = Cells(i, zvOst)
Cells(i, zvCol).Select
MsgBox "Превышен лимит остатков склада!" & Chr(10) & Chr(10) & _
"   На складе осталось: " & Cells(i, zvOst) & " шт", 64, "Лимит"
Exit Function
End If
End If
r = Cells(Rows.Count, zvNm).End(xlUp).Row
Cells(i, zvSm) = Cells(i, zvCol) * Cells(i, zvCnR)
Cells(rwzvSm, zvSm) = Application.Sum(Range(Cells(rwZv, zvSm), Cells(r + 4, zvSm)))
End Function


Private Sub Worksheet_Deactivate()
On Error Resume Next
Call unload_mn_vid
DoEvents
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        On Error Resume Next
        Call unload_mn_mn
        If ThisWorkbook.Sheets("Расход").Shapes("mn_vid").Visible = True Then Call unload_mn_vid
        
        If Target.Count > 1 Then Exit Sub
        
        'skid
        If Not Intersect(Target, Cells(rwZv_mj, zvOst)) Is Nothing Then
            Call frm_Skidka.Show
            GoTo 99
        End If
        
        'oplata
        If Not Intersect(Target, Cells(rwZv_mj, zvSm)) Is Nothing Then
            Call frm_Oplata.Show
            GoTo 99
        End If
99
End Sub


