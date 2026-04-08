' Component: Лист4  [Приход]
' Type: Document
Option Explicit

Dim r As Long

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error Resume Next
If Target.Count > 1 Then Exit Sub
If Target.Row < rwZv Then Exit Sub
i = Target.Row

        If Not Intersect(Target, Cells(i, prGr)) Is Nothing Then
            If Cells(i, prNN) <> "" Then
                Cancel = True
                Call sh_gr
                GoTo 99
            End If
        End If

If Not Intersect(Target, Range(Cells(i, prNm), Cells(i, prSm))) Is Nothing Then
If Cells(i, prNN) <> "" Then
Cancel = True
Call Add_MyContextMenu_pr
Application.CommandBars("MyContextMenu_pr").ShowPopup
End If
End If
99
End Sub

Private Sub sh_gr()
On Error Resume Next
        With frm_Gr
            .Show
            .CheckBox1.Value = False
        End With
End Sub

Private Sub Add_MyContextMenu_pr()
On Error Resume Next
Application.CommandBars("MyContextMenu_pr").Delete
With Application.CommandBars.Add(Name:="MyContextMenu_pr", Position:=msoBarPopup, Temporary:=True)
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 21
.Caption = "Удалить позицию"
.OnAction = "del_poz_pr"
End With
End With
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
Dim ra As Integer
If Target.Count > 1 Then Exit Sub
If Target.Row < rwZv Then Exit Sub
r = Cells(Rows.Count, prNm).End(xlUp).Row
ra = Target.Row
If Not Intersect(Target, Range(Cells(rwZv, prCol), Cells(r, prCnR))) Is Nothing Then
Call проверка_остатков(ra)
End If
End Sub

Private Function проверка_остатков(i As Integer)
On Error Resume Next
r = Cells(Rows.Count, prNm).End(xlUp).Row
Cells(i, prSm) = Cells(i, prCol) * Cells(i, prCnZ)
Cells(rwzvSm, prSm) = Application.Sum(Range(Cells(rwZv, prSm), Cells(r + 4, prSm)))
End Function

Private Sub Worksheet_Deactivate()
On Error Resume Next
Call unload_mn_vid_pr
DoEvents
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Call unload_mn_mn
If ThisWorkbook.ActiveSheet.Shapes("mn_vid_pr").Visible = True Then Call unload_mn_vid_pr
End Sub

