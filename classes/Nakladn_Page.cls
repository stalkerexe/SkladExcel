VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Nakladn_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents nNm As MSForms.TextBox
Attribute nNm.VB_VarHelpID = -1
Public WithEvents nCol As MSForms.TextBox
Attribute nCol.VB_VarHelpID = -1
Public WithEvents nSk As MSForms.TextBox
Attribute nSk.VB_VarHelpID = -1
Public WithEvents nNN As MSForms.TextBox
Attribute nNN.VB_VarHelpID = -1
Public WithEvents nCod As MSForms.TextBox
Attribute nCod.VB_VarHelpID = -1


Private Sub nCol_Change()
On Error Resume Next
nom_Cnr = Mid(nCol.Name, 5)
iRow = nom_Cnr + rwZv - 1
With ThisWorkbook.Sheets("корзина")
.Cells(iRow, zvCol) = Replace(nCol, ",", ".")
End With
Call sum_box
End Sub
Private Sub nCol_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
nom_Cnr = Mid(nCol.Name, 5)
frm_Show.tb_nom.Value = nom_Cnr
frm_Show.SpinButton.Visible = True
frm_Show.SpinButton.Top = nNm.Top - 2
frm_Show.SpinButton.Left = nCol.Left - frm_Show.SpinButton.Width
frm_Show.SpinButton.ZOrder 0
Call sh_ico_del
Call color_row
End Sub
Private Sub nNN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
nom_Cnr = Mid(nNN.Name, 4)
frm_Show.tb_nom.Value = nom_Cnr
Call sh_ico_del
frm_Show.SpinButton.Visible = False
Call color_row
End Sub
Private Sub nNm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
nom_Cnr = Mid(nNm.Name, 4)
frm_Show.tb_nom.Value = nom_Cnr
Call sh_ico_del
frm_Show.SpinButton.Visible = False
Call color_row
End Sub
Private Sub nCod_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
nom_Cnr = Mid(nCod.Name, 5)
frm_Show.tb_nom.Value = nom_Cnr
Call sh_ico_del
frm_Show.SpinButton.Visible = False
Call color_row
End Sub
Private Sub nSk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
nom_Cnr = Mid(nSk.Name, 4)
frm_Show.tb_nom.Value = nom_Cnr
Call sh_ico_del
frm_Show.SpinButton.Visible = False
Call color_row
End Sub
Private Sub sh_ico_del()
On Error Resume Next
frm_Show.ico_del.Visible = True
frm_Show.ico_del.Top = nNm.Top
frm_Show.ico_del.Left = nSk.Left + nSk.Width + 9
End Sub
Private Sub color_row()
On Error Resume Next
Call clear_color
nNN.BackColor = &H80000018
nNm.BackColor = &H80000018
nCol.BackColor = &H80000018
nCod.BackColor = &H80000018
nSk.BackColor = &H80000018
End Sub

