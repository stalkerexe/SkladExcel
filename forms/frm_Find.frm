VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Find 
   Caption         =   "Поиск"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   OleObjectBlob   =   "frm_Find.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim w(): Dim cc()
Dim iCntr As Integer

Private Sub comb_gr_Change()
On Error Resume Next
ind = comb_gr.ListIndex
If ind = -1 Then Exit Sub
rw = comb_gr.List(ind, 0) + 4
Range(Cells(rw, 1), Cells(rw, 33)).Select
ActiveWindow.ScrollRow = rw - 1
End Sub
Private Sub tb_1_Change()
On Error Resume Next
Call find_
If Val(comb_find.ListCount) > 0 Then
comb_find.DropDown
End If
End Sub
Private Sub tb_2_Change()
On Error Resume Next
Call find_
If Val(comb_find.ListCount) > 0 Then
comb_find.DropDown
End If
End Sub
Private Sub find_()
On Error Resume Next
iCntr = VBA.Right(Me.ActiveControl.Name, 1)
For n = 1 To 2
If n <> iCntr Then
Me.Controls("tb_" & n) = ""
End If
Next
comb_find.clear
str_ = Me.ActiveControl.Text
If Len(str_) = 0 Then comb_find.SetFocus: Controls("tb_" & iCntr).SetFocus: Exit Sub
If str_ = "" Then
comb_find.SetFocus
Controls("tb_" & iCntr).SetFocus
Exit Sub
End If
ReDim cc(LBound(c) To UBound(c), 1 To 3)
iCol = 0
For i = LBound(c) To UBound(c)
If Len(str_) = 1 Then
If VBA.UCase(VBA.Left(c(i, iCntr + 1), 1)) = VBA.UCase(str_) Then
cc(i, 1) = c(i, 1)
cc(i, 2) = c(i, 2)
cc(i, 3) = c(i, 3)
iCol = iCol + 1
End If
Else
If InStr(1, VBA.UCase(c(i, iCntr + 1)), VBA.UCase(str_)) > 0 Then
cc(i, 1) = c(i, 1)
cc(i, 2) = c(i, 2)
cc(i, 3) = c(i, 3)
iCol = iCol + 1
End If
End If
Next
If iCol = 0 Then
comb_find.clear
comb_find.SetFocus
Controls("tb_" & iCntr).SetFocus
DoEvents
Exit Sub
End If
ReDim w(1 To iCol + 1, 1 To 3)
j = 1
For i = LBound(c) To UBound(c)
If cc(i, 1) <> "" Then
w(j, 1) = cc(i, 1)
w(j, 2) = cc(i, 2)
w(j, 3) = cc(i, 3)
j = j + 1
End If
Next
comb_find.clear
comb_find.SetFocus
Controls("tb_" & iCntr).SetFocus
DoEvents
comb_find.List = w
End Sub
Private Sub comb_find_Click()
On Error Resume Next
ind = comb_find.ListIndex
If ind = -1 Then Exit Sub
Me.Caption = "   " & comb_find.List(ind, 1) & "          " & comb_find.List(ind, 2)
tb_row.Text = comb_find.List(ind, 0)
rw = Val(tb_row.Value)
If rw = 0 Then Exit Sub
Range(Cells(rw, 1), Cells(rw, 33)).Select
ActiveWindow.ScrollRow = rw - 1
End Sub
Private Sub tb_1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
If Val(comb_find.ListCount) > 0 Then comb_find.DropDown
End Sub
Private Sub tb_2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
If Val(comb_find.ListCount) > 0 Then comb_find.DropDown
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
With ActiveSheet.Shapes("cmb_f")
Me.Top = .Top + .Height + 20
Me.Left = .Left
End With
Call load_sk__
Call cod_YN
With comb_find
.Left = tb_1.Left
.Top = tb_2.Top
.Width = tb_1.Width + tb_2.Width
.ZOrder 1
End With
If Sheets("setting").Range("b6") = 1 Then
comb_find.ColumnWidths = "0;" & tb_1.Width & ";" & tb_2.Width - 2
Else
comb_find.ColumnWidths = "0;0;" & tb_2.Width - 2
End If
With comb_gr
.Left = lb_nm.Left
.Top = lb_nm.Top
.Width = lb_nm.Width + 13
.ZOrder 1
End With
End Sub
Private Sub cod_YN()
On Error Resume Next
If Sheets("setting").Range("b6") = 0 Then
lb_cod.Width = 0: tb_1.Width = 0
lb_nm.Left = lb_cod.Left: tb_2.Left = tb_1.Left
lb_nm.Width = Me.Width
tb_2.Width = Me.Width
End If
End Sub
Private Sub load_sk__()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
r7 = .Cells(Rows.Count, skNm).End(xlUp).Row + 2
gr = .Range(.Cells(5, skGr), .Cells(r7, skGr)).Value
cod = .Range(.Cells(5, skCod), .Cells(r7, skCod)).Value
nm = .Range(.Cells(5, skNm), .Cells(r7, skNm)).Value
iCol = Application.CountIf(.Range(.Cells(5, skNm), .Cells(r7, skNm)), "<>")
End With
If iCol = 0 Then Exit Sub
Call parse_arr_sk
End Sub
Private Sub parse_arr_sk()
On Error Resume Next
ReDim c(LBound(nm) To UBound(nm), 1 To 3)
j = 1
n = 0
For i = LBound(nm) To UBound(nm)
If nm(i, 1) <> "" Then
c(j, 1) = i + 4
c(j, 2) = cod(i, 1)
c(j, 3) = nm(i, 1)
If gr(i, 1) <> "" Then
c(j, 2) = nm(i, 1) & " -----------------------------------------------------------"
c(j, 3) = "---------------------------------"
comb_gr.AddItem ""
comb_gr.List(n, 0) = j
comb_gr.List(n, 1) = nm(i, 1)
n = n + 1
End If
j = j + 1
End If
Next i
End Sub
Private Sub lb_nm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
Me.comb_gr.DropDown
tb_1.Text = ""
tb_2.Text = ""
End Sub

