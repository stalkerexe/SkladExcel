VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ZVK 
   ClientHeight    =   8265.001
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   17265
   OleObjectBlob   =   "frm_ZVK.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ZVK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CheckBox_perenos_Click()

        If CheckBox_perenos.Value = True Then
            Call perenos_yes
            ThisWorkbook.Sheets("setting").Range("f12") = 1
        Else
            Call perenos_no
            ThisWorkbook.Sheets("setting").Range("f12") = 0
        End If
        
        Frame_set.Visible = False
        
End Sub











Private Sub ico_close_set_Click()
        Frame_set.Visible = False
End Sub





Private Sub CheckBox_vz_Click()
        On Error Resume Next

        Call col_Controls

        If CheckBox_vz.Value = True Then
                
            For i = 1 To iColCtr
                Controls("nCol_vz" & i).Value = Controls("nCol" & i).Value
            Next
            
        Else
            
            For i = 1 To iColCtr
                Controls("nCol_vz" & i).Value = ""
            Next
        
        End If
End Sub


Private Sub comb_sk_Click()
        On Error Resume Next

        nom_Cnr = Val(tb_nom.Value)
        
        i = comb_sk.ListIndex
        If i = -1 Then Exit Sub
        
        Controls("nSk_vz" & nom_Cnr).Value = comb_sk.Value
        
End Sub








Private Sub ico_set_Click()
        On Error Resume Next
        
        Call visible_NO
        
        With Frame_set
            If .Visible = False Then
                .Visible = True
                .Left = Frame_zg.Left + ico_set.Left - 10
                .Top = Frame_zg.Top + ico_set.Top + ico_set.Height + 2
                .ZOrder 0
            Else
                .Visible = False
            End If
        End With
End Sub

Private Sub OK_arh_Click()
        Call unload_menu_button
        frm_msg.Show
End Sub



Private Sub OK_Click()

        SpinButton.Visible = False
        
        Call proverka_vz
        
        If iCol = 0 Then
            MsgBox "Не выбраны позиции для возврата!", 64, "Возврат на склад"
            Exit Sub
        End If
        
        frm_msg.Show
        
End Sub





Private Sub OK_otmena_Click()
        Call unload_menu_button
        frm_msg.Show
End Sub





Private Sub cmb_Print_Click()
        Call unload_menu_button
        marker = tb_mk.Text:  If marker = "" Then Exit Sub
        Call printZVK
End Sub




Private Sub OK_vz_Click()
        On Error Resume Next

        Call unload_menu_button
        
        Me.Height = frmHg2
        Me.Width = frmWd2
        Me.Left = Me.Left - Me.Frame_nk_all_vz.Width
        Frame_button.Visible = True
        
End Sub















Private Sub UserForm_Initialize()
        Call forma
        Call load_sklads
End Sub

Private Sub UserForm_Click()
        Call visible_NO
End Sub

Private Sub forma()
        On Error Resume Next
        
        ScrollBar1.Width = 0
Me.Height = frmHg2
        Me.Width = frmWd1
        
        Me.Frame_button.BackColor = Me.BackColor
        Me.Frame_dann.BackColor = Me.BackColor
        Me.Frame_nk.BackColor = Me.BackColor
        Me.Frame_nk_all.BackColor = Me.BackColor
        Me.Frame_nk_vz.BackColor = Me.BackColor
        Me.Frame_nk_all_vz.BackColor = Me.BackColor
        Me.Frame_set.BackColor = Me.BackColor
        Me.Frame_zg.BackColor = Me.BackColor
        Me.Frame_zg2.BackColor = Me.BackColor
        Me.Frame_menu.BackColor = Me.BackColor
        
        Me.Frame_menu.Width = OK_arh.Width
        

End Sub

Private Sub load_sklads()
        On Error Resume Next
        Call LoadSkToControl(comb_sk)
End Sub






Private Sub Frame_nk_all_Click()
        Call visible_NO
End Sub

Private Sub Frame_nk_all_vz_Click()
        Call visible_NO
End Sub

Private Sub Frame_dann_Click()
        Call visible_NO
End Sub

Private Sub Frame_nk_vz_Click()
        Call visible_NO
End Sub

Private Sub tb_Mnj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call visible_NO
End Sub

Private Sub Frame_nk_Click()
        Call visible_NO
End Sub

Private Sub tb_Dt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call visible_NO
End Sub

Private Sub tb_Zkz_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call visible_NO
End Sub


Private Sub visible_NO()
        comb_sk.Visible = False
        SpinButton.Visible = False
        Frame_set.Visible = False
        Call unload_menu_button
End Sub


Private Sub NO2_Click()

        With Me
        
            .Width = frmWd1
            
            .Left = .Left + .Frame_nk_all_vz.Width
            
            Frame_button.Visible = False
            
            .Repaint
        
        End With

End Sub




Private Sub NO_Click()
        Unload Me
End Sub






















Private Sub ScrollBar1_Change()
Frame_nk.Top = -ScrollBar1.Value
Frame_nk_vz.Top = -ScrollBar1.Value
End Sub

Private Sub ScrollBar1_Scroll()
Frame_nk.Top = -ScrollBar1.Value
Frame_nk_vz.Top = -ScrollBar1.Value
End Sub



Private Sub SpinButton_SpinDown()
        On Error Resume Next
        nom_Cnr = Val(tb_nom.Value)
        If Controls("nCol_vz" & nom_Cnr).Value = "" Then sCol = 0: Exit Sub
        sCol = Controls("nCol_vz" & nom_Cnr).Value
        If sCol <= 0 Then sCol = 0: Exit Sub
33
        sCol = sCol - 1
        Controls("nCol_vz" & nom_Cnr).Value = sCol
End Sub

Private Sub SpinButton_SpinUp()
        On Error Resume Next
        nom_Cnr = Val(tb_nom.Value)
        If Controls("nCol_vz" & nom_Cnr).Value = "" Then sCol = 0: GoTo 33
        sCol = Controls("nCol_vz" & nom_Cnr).Value
        If sCol >= Controls("nCol" & nom_Cnr).Value Then sCol = Controls("nCol" & nom_Cnr).Value: Exit Sub
33
        sCol = sCol + 1
        Controls("nCol_vz" & nom_Cnr).Value = sCol
        If sCol > Controls("nCol" & nom_Cnr).Value Then Controls("nCol_vz" & nom_Cnr).Value = Controls("nCol" & nom_Cnr).Value: Exit Sub
End Sub




Private Sub OK_menu_Click()
        Call show_menu_button
End Sub

Private Sub show_menu_button()
        On Error Resume Next
        With Me.Frame_menu
            .Height = 0
            .Visible = True
            .SpecialEffect = 0
            .ZOrder 0
            .Top = OK_menu.Top + OK_menu.Height + 2
            .Left = OK_menu.Left
        End With
        Call do_show_menu_button
End Sub

Private Sub do_show_menu_button()
        On Error Resume Next
        
        If tb_what.Text = "Отгрузка" Then
            iHg = 80
        Else
            iHg = 60
        End If
        
        If tb_what.Text = "Возврат" Then iHg = 40
        
        Me.Frame_menu.Height = iHg

End Sub

Private Sub unload_menu_button()
        On Error Resume Next
        Me.Frame_menu.Visible = False
End Sub

