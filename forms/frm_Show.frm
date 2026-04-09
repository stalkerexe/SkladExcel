VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Show 
   Caption         =   "Корзина"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   OleObjectBlob   =   "frm_Show.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim r As Long
Private Sub OK_Click()
On Error Resume Next
With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r + 3, zvNm)), "<>")
End With
If iCol = 0 Then
MsgBox "   В корзине нет товара!   ", 64, "Оформить заказ"
Exit Sub
End If
If MsgBox("   Оформить накладную?   ", vbOKCancel + vbQuestion, "Оформить заказ") = vbCancel Then Exit Sub
Call оформить_заказ
End Sub
Private Sub OK_pr_Click()
On Error Resume Next
With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r + 3, zvNm)), "<>")
End With
If iCol = 0 Then
MsgBox "   В корзине нет товара!   ", 64, "Оформить заказ"
Exit Sub
End If
If MsgBox("   Оформить накладную?   ", vbOKCancel + vbQuestion, "Приход") = vbCancel Then Exit Sub
Call оформить_заказ_pr
End Sub
Private Sub CB_clear_Click()
On Error Resume Next
If MsgBox("Очистить корзину?", vbOKCancel + vbQuestion, "Очистить") = vbCancel Then Exit Sub
Call controls_all_delete
Call clear_box
Call clear_color
SpinButton.Visible = False
ico_del.Visible = False
tb_sm.value = "0"
End Sub
Private Sub controls_all_delete()
On Error Resume Next

For Each ctr In Me.Frame_nk.Controls
If ctr.Name = "ico_del" Then GoTo 33
If ctr.Name = "SpinButton" Then GoTo 33
Me.Frame_nk.Controls.Remove ctr.Name
33
Next
frm_Show.ScrollBar1.Width = 0
End Sub
Private Sub clear_zv()
On Error Resume Next
With ThisWorkbook.Sheets("Расход")
r = .UsedRange.Rows.Count + .UsedRange.Row - 1
.Range(.Cells(rwZv, 2), .Cells(r + 44, 2)).EntireRow.Delete
.Cells(rwzvSm, zvSm) = ""
End With
End Sub
Private Sub CheckBox1_Click()
On Error Resume Next
Call sk_hide
ico_del.Visible = False
SpinButton.Visible = False
Call clear_color
End Sub
Private Sub sk_hide()
On Error Resume Next

For Each ctr In Me.Frame_nk.Controls
If ctr.Left = lb_sk.Left Then
If CheckBox1.value = True Then
ctr.Width = lb_sk.Width: lb_sk.Visible = True: ThisWorkbook.Sheets("setting").Range("h4") = 1
Else
ctr.Width = 0: lb_sk.Visible = False: ThisWorkbook.Sheets("setting").Range("h4") = 0
End If
End If
Next
Call Frame_width
End Sub
Private Sub ico_del_Click()
On Error Resume Next
nom_Cnr = Val(tb_nom.value)
iRow = nom_Cnr + rwZv - 1
ico_del.Visible = False
SpinButton.Visible = False
Call del_poz_box
Call добавить_контролы
Call clear_color
End Sub
Private Sub SpinButton_SpinDown()
On Error Resume Next
nom_Cnr = Val(tb_nom.value)
iRow = nom_Cnr + rwZv - 1
With ThisWorkbook.Sheets("корзина")
.Cells(iRow, zvCol) = .Cells(iRow, zvCol) - 1
If .Cells(iRow, zvCol) <= 0 Then .Cells(iRow, zvCol) = 0
sCol = .Cells(iRow, zvCol)
Controls("nCol" & nom_Cnr).value = sCol
End With
End Sub
Private Sub SpinButton_SpinUp()
On Error Resume Next
nom_Cnr = Val(tb_nom.value)
iRow = nom_Cnr + rwZv - 1
With ThisWorkbook.Sheets("корзина")
.Cells(iRow, zvCol) = .Cells(iRow, zvCol) + 1
sCol = .Cells(iRow, zvCol)
Controls("nCol" & nom_Cnr).value = sCol
End With
End Sub
Private Sub ScrollBar1_Change()
Frame_nk.Top = -ScrollBar1.value
End Sub
Private Sub ScrollBar1_Scroll()
Frame_nk.Top = -ScrollBar1.value
End Sub
Private Sub UserForm_Click()
ico_del.Visible = False
SpinButton.Visible = False
Call clear_color
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
If ThisWorkbook.Sheets("setting").Range("b6") = 0 Then
lb_cod.Width = 0
lb_col.Left = lb_cod.Left
lb_sk.Left = lb_col.Left + lb_col.Width
CheckBox1.Left = lb_sk.Left + 2
End If
If ThisWorkbook.Sheets("setting").Range("b8") = 0 Then tb_sm.Visible = False
ScrollBar1.Top = Frame_nk_all.Top
ScrollBar1.Height = Frame_nk_all.Height
If ThisWorkbook.Sheets("setting").Range("h4") = 1 Then CheckBox1.value = True: lb_sk.Visible = True
If ThisWorkbook.Sheets("setting").Range("h4") = 0 Then lb_sk.Visible = False
tb_sm.Left = Frame_nk_all.Left + lb_col.Left
End Sub
Private Sub NO_Click()
Unload Me
End Sub
Private Sub Frame_nk_all_Click()
On Error Resume Next
SpinButton.Visible = False
ico_del.Visible = False
Call clear_color
End Sub
