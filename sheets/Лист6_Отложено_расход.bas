' Component: Лист6  [Отложено_расход]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit

Dim r As Long

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error Resume Next
If Target.Count > 1 Then Exit Sub
r = Cells(Rows.Count, zkNm).End(xlUp).Row
i = Target.Row
If Not Intersect(Target, Range(Cells(4, zkDt1 - 1), Cells(r, zkSm))) Is Nothing Then
If Cells(i, "a") <> "" Then
Cancel = True
Add_MyContextMenu
Application.CommandBars("MyContextMenu").ShowPopup
End If
End If
If Not Intersect(Target, Cells(i, zkComm)) Is Nothing Then
If Cells(i, zkNm) <> "" Then
Cancel = True
End If
End If
End Sub

Private Sub Add_MyContextMenu()
On Error Resume Next
Application.CommandBars("MyContextMenu").Delete
With Application.CommandBars.Add(Name:="MyContextMenu", Position:=msoBarPopup, Temporary:=True)
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 162
.Caption = "Редактировать"
.OnAction = "zv_pedactirov_"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 4
.Caption = "Печать"
.OnAction = "printZk"
End With


With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 3160
.Caption = "Отгрузить"
.OnAction = "otgr_zk"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 21
.Caption = "Удалить заказ"
.OnAction = "delete_zv"
End With
End With
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Call unload_mn_mn
If Target.Row < 5 Then Exit Sub
i = Target.Row
If ActiveCell.Column = zkComm Then
If Cells(i, 1) = "" Then
If Cells(i, zkNm) <> "" Then
Call sh_frm_Comm
Cancel = True
End If
End If
End If
End Sub
Private Sub sh_frm_Comm()
On Error Resume Next
iRow = ActiveCell.Row
With frm_Comm
.Show
.tb_row.value = iRow
.tb_mk.Text = Cells(iRow - 1, 1)
.tb_comm.Text = ActiveCell.value
.tb_sheet.Text = "rs"
End With
End Sub



