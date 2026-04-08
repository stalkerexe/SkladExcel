' Component: Лист9  [Отложено_приход]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
        On Error Resume Next
        
        If Target.Count > 1 Then Exit Sub
        If Target.Row < 5 Then Exit Sub
        
        i = Target.Row
        
        If Not Intersect(Target, Range(Cells(i, pzkNom), Cells(i, pzkComm))) Is Nothing Then
            If Cells(i, "a") <> "" Then
                Cancel = True
                Add_MyContextMenu
                Application.CommandBars("MyContextMenu").ShowPopup
            End If
            GoTo 99
        End If
        
        On Error Resume Next
        If Not Intersect(Target, Cells(i, pzkComm)) Is Nothing Then
            If Cells(i, pzkNm) <> "" Then
            Cancel = True
            End If
        End If
        
        On Error Resume Next
        If Not Intersect(Target, Cells(i, pzkOsn)) Is Nothing Then
            Cancel = True
        End If

99
End Sub


Private Sub Add_MyContextMenu()
On Error Resume Next
Application.CommandBars("MyContextMenu").Delete
With Application.CommandBars.Add(Name:="MyContextMenu", Position:=msoBarPopup, Temporary:=True)
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 162
.Caption = "Редактировать"
.OnAction = "zv_pedactirov_pr"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 4
.Caption = "Печать"
.OnAction = "printZk_pr"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 3160
.Caption = "Приходовать"
.OnAction = "prZk"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 21
.Caption = "Удалить приход"
.OnAction = "prDelete"
End With
End With
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        On Error Resume Next
        Call unload_mn_mn
        If Target.Row < 5 Then Exit Sub
        
        i = Target.Row
        
        If ActiveCell.Column = pzkComm Then
        If Cells(i, 1) = "" Then
        If Cells(i, pzkNm) <> "" Then
            iRow = i
            Call sh_frm_Comm
            Cancel = True
        End If
        End If
        End If
        
        
        If ActiveCell.Column = pzkOsn Then
        If Cells(i, 1) = "" Then
        If Cells(i, pzkNm) <> "" Then
            iRow = i
            Call sh_frm_Doc
            Cancel = True
        End If
        End If
        End If
        
End Sub

Private Sub sh_frm_Comm()
On Error Resume Next
With frm_Comm
.Show
.tb_row.Value = iRow
.tb_mk.text = Cells(iRow - 1, 1)
.tb_comm.text = ActiveCell.Value
.tb_sheet.text = "pr"
End With
End Sub

Private Sub sh_frm_Doc()
On Error Resume Next
iRow = ActiveCell.Row - 1
With frm_Doc
.Show
.tb_row.Value = iRow
.tb_doc.text = Cells(iRow, pzkDoc)
.tb_docN.text = Cells(iRow, pzkDocN)
.tb_dt1.text = Cells(iRow, pzkDt)
.tb_mk.text = Cells(iRow, 1)
.tb_sheet.text = "pr"
End With
End Sub

