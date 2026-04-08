VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_zvk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents nNm As MSForms.TextBox
Attribute nNm.VB_VarHelpID = -1
Public WithEvents nNN As MSForms.TextBox
Attribute nNN.VB_VarHelpID = -1
Public WithEvents nCod As MSForms.TextBox
Attribute nCod.VB_VarHelpID = -1
Public WithEvents nEd As MSForms.TextBox
Attribute nEd.VB_VarHelpID = -1
Public WithEvents nCn As MSForms.TextBox
Attribute nCn.VB_VarHelpID = -1
Public WithEvents nSm As MSForms.TextBox
Attribute nSm.VB_VarHelpID = -1

Public WithEvents nCol As MSForms.TextBox
Attribute nCol.VB_VarHelpID = -1
Public WithEvents nSk As MSForms.TextBox
Attribute nSk.VB_VarHelpID = -1

Public WithEvents nCol_vz As MSForms.TextBox
Attribute nCol_vz.VB_VarHelpID = -1
Public WithEvents nSk_vz As MSForms.TextBox
Attribute nSk_vz.VB_VarHelpID = -1



Private Sub nNN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nNm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nCod_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nEd_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nCn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nSm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nCol_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub

Private Sub nSk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
End Sub


Private Sub nCol_vz_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        On Error Resume Next
        Call find_nom_Cnr
        Call SpinButton_show
End Sub

Private Sub nSk_vz_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call find_nom_Cnr
        Call sh_comb
End Sub



Private Sub SpinButton_show()
        On Error Resume Next
        With frm_ZVK
            .SpinButton.Visible = True
            .SpinButton.Top = nCol_vz.Top - 2
.SpinButton.Left = nSk_vz.Left
            .SpinButton.ZOrder 0
        End With
End Sub


Private Sub sh_comb()
        On Error Resume Next
        With frm_ZVK.comb_sk
            .Visible = True
            .ListIndex = -1
.Top = nSk_vz.Top
            .Left = nSk_vz.Left
            .Width = nSk_vz.Width + 13
.BackColor = &H80000018
            .DropDown
        End With
End Sub


Private Sub find_nom_Cnr()
        nom_Cnr = CInt(nNN.Value)
        frm_ZVK.tb_nom.Value = nom_Cnr
        Call visible_NO
        Call color_row
End Sub

Private Sub visible_NO()
        With frm_ZVK
            .comb_sk.Visible = False
            .SpinButton.Visible = False
            .Frame_set.Visible = False
        End With
End Sub



Private Sub color_row()
        On Error Resume Next
        
        Call visible_menu_zvk_NO
        Call clear_color_vz
        
        nNN.BackColor = &H80000018
        nNm.BackColor = &H80000018
        nCod.BackColor = &H80000018
        nEd.BackColor = &H80000018
        nCn.BackColor = &H80000018
        nSm.BackColor = &H80000018
        
        nCol.BackColor = &H80000018
        nSk.BackColor = &H80000018
        
        nCol_vz.BackColor = &H80000018
        nSk_vz.BackColor = &H80000018
        
End Sub

Private Sub visible_menu_zvk_NO()
        On Error Resume Next
        frm_ZVK.Frame_menu.Visible = False
End Sub

