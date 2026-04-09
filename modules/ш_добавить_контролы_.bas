Attribute VB_Name = "ш_добавить_контролы_"
Option Explicit

Dim r As Long

Public Sub добавить_контролы()
On Error Resume Next
iSize = 10
controls_all_delete
Call arr_box
Call контролы_накладной
Call Frame_height
Call Frame_width
Call sum_box
frm_Show.SpinButton.Visible = False
frm_Show.ico_del.Visible = False
DoEvents
End Sub
Private Sub контролы_накладной()
On Error Resume Next
With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r + 3, zvNm)), "<>")
End With
rows_Cnr = iCol
If iCol = 0 Then Exit Sub
ReDim arr_Ctr(1 To rows_Cnr)
nom_Cnr = 1
For i = LBound(nm) To UBound(nm)
If nm(i, 1) <> "" Then
Call cnrt_add_nk
nom_Cnr = nom_Cnr + 1
End If
Next
End Sub
Private Sub arr_box()
On Error Resume Next
With ThisWorkbook.Sheets("корзина")
r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 1
nn = Range(.Cells(rwZv, zvNN), .Cells(r7, zvNN)).value
nm = Range(.Cells(rwZv, zvNm), .Cells(r7, zvNm)).value
cod = Range(.Cells(rwZv, zvCod), .Cells(r7, zvCod)).value
col = Range(.Cells(rwZv, zvCol), .Cells(r7, zvCol)).value
sk = Range(.Cells(rwZv, zvSk), .Cells(r7, zvSk)).value
End With
End Sub
Private Sub cnrt_add_nk()
On Error Resume Next
top1 = 0

nmCntr = "nNN" & nom_Cnr
txCntr = nn(i, 1)
leftCntr = 0
gnCntr = fmTextAlignCenter
widCntr = frm_Show.lb_nn.Width + 1
flag_hidden = True
Call add_cntrl_nk

nmCntr = "nNm" & nom_Cnr
txCntr = nm(i, 1)
leftCntr = frm_Show.lb_nm.Left
gnCntr = fmTextAlignLeft
widCntr = frm_Show.lb_nm.Width + 1
flag_hidden = True
Call add_cntrl_nk

nmCntr = "nCod" & nom_Cnr
txCntr = cod(i, 1)
leftCntr = frm_Show.lb_cod.Left
gnCntr = fmTextAlignLeft
widCntr = frm_Show.lb_cod.Width + 1
flag_hidden = True
If ThisWorkbook.Sheets("setting").Range("b6") = 0 Then widCntr = 0
Call add_cntrl_nk

nmCntr = "nCol" & nom_Cnr
txCntr = col(i, 1)
leftCntr = frm_Show.lb_col.Left
gnCntr = fmTextAlignCenter
widCntr = frm_Show.lb_col.Width + 1
flag_hidden = False
Call add_cntrl_nk
nmCntr = "nSk" & nom_Cnr
txCntr = sk(i, 1)
leftCntr = frm_Show.lb_sk.Left
gnCntr = fmTextAlignCenter
widCntr = frm_Show.lb_sk.Width
flag_hidden = True
If ThisWorkbook.Sheets("setting").Range("h4") = 0 Then widCntr = 0
Call add_cntrl_nk
Set arr_Ctr(nom_Cnr).nNm = frm_Show.Controls("nNm" & nom_Cnr)
Set arr_Ctr(nom_Cnr).nCol = frm_Show.Controls("nCol" & nom_Cnr)
Set arr_Ctr(nom_Cnr).nSk = frm_Show.Controls("nSk" & nom_Cnr)
Set arr_Ctr(nom_Cnr).nNN = frm_Show.Controls("nNN" & nom_Cnr)
Set arr_Ctr(nom_Cnr).nCod = frm_Show.Controls("nCod" & nom_Cnr)
End Sub
Private Sub add_cntrl_nk()
On Error Resume Next
With frm_Show.Frame_nk.Controls.Add("Forms.TextBox.1")
.Name = nmCntr
.Text = txCntr
.Height = hgCntr
.Top = top1 + (nom_Cnr - 1) * (hgCntr - 1)
.Left = leftCntr
.Width = widCntr
.BorderStyle = 1
.TextAlign = gnCntr
.Locked = flag_hidden
.Font.Size = iSize
.Font.Name = "Times New Roman"
.MousePointer = 1
If nmCntr = "nCol" & nom_Cnr Then .MousePointer = 3
End With
End Sub
Private Sub Frame_height()
On Error Resume Next
With frm_Show
.Frame_nk.Height = (hgCntr - 1) * rows_Cnr + 1
If .Frame_nk.Height > .Frame_nk_all.Height Then .ScrollBar1.Width = 12
.ScrollBar1.Min = 0
.ScrollBar1.Max = Val(.Frame_nk.Height - .Frame_nk_all.Height)
If .ScrollBar1.Width > 0 Then .ScrollBar1.value = .ScrollBar1.Max
End With
End Sub
Public Sub Frame_width()
On Error Resume Next
With frm_Show
If ThisWorkbook.Sheets("setting").Range("h4") = 1 Then
.ico_del.Left = .lb_sk.Left + .lb_sk.Width + 9
Else
.ico_del.Left = .lb_sk.Left + 9
End If
.Frame_nk.Width = .ico_del.Left + .ico_del.Width + 8
.Frame_nk_all.Width = .Frame_nk.Width
.ScrollBar1.Left = .Frame_nk_all.Left + .Frame_nk_all.Width + 2
.Width = .Frame_nk_all.Width + .ScrollBar1.Width + 27
End With
End Sub
Private Sub controls_all_delete()
On Error Resume Next
For Each ctr In frm_Show.Frame_nk.Controls
If ctr.Name = "ico_del" Then GoTo 33
If ctr.Name = "SpinButton" Then GoTo 33
frm_Show.Frame_nk.Controls.Remove ctr.Name
33
Next
frm_Show.ScrollBar1.Width = 0
End Sub
Public Sub sum_box()
On Error Resume Next
summ = ThisWorkbook.Sheets("корзина").Cells(rwzvSm, zvSm)
frm_Show.tb_sm.value = Format(summ, "#,##0.00")
ThisWorkbook.Sheets("Склад").Cells(3, iBox2) = ThisWorkbook.Sheets("корзина").Cells(rwzvSm, zvSm)
With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r + 7, zvNm)), "<>")
End With
ThisWorkbook.Sheets("Склад").Cells(3, iBox1) = iCol
End Sub
Public Sub clear_color()
On Error Resume Next

For Each ctr In frm_Show.Frame_nk.Controls
If ctr.Name = "ico_del" Then GoTo 33
If ctr.Name = "SpinButton" Then GoTo 33
ctr.BackColor = &H80000005
33
Next
End Sub

