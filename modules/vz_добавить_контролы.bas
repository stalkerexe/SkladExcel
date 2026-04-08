Attribute VB_Name = "vz_добавить_контролы"
Option Explicit

Private Const hgCntr = 18


Public Sub добавить_контролы_vz()

        iColCtr = iCol
        If iColCtr = 0 Then Exit Sub
        
        iVid = zvSelect.comb_vid.Value

        Call контролы_накладной
        Call arr_controls_vz
        Call Frame_height

End Sub

Private Sub контролы_накладной()
        On Error Resume Next
        
        iSize = 9

        ReDim arr_Ctr_zvk(1 To iColCtr)
        
        nom_Cnr = 1
        
        For i = LBound(nm) To UBound(nm)
            If nn(i, 1) <> "" Then
            
                Call dann_poz
                Call controls_add
                
                nom_Cnr = nom_Cnr + 1
                
            End If
        Next
        
End Sub

Private Sub dann_poz()
        On Error Resume Next
        sNN = nn(i, 1)
        sNm = nm(i, 1)
        sCod = cod(i, 1)
        sEd = ed(i, 1)
        sSm = sm(i, 1)
        sCol = col(i, 1)
        sSk = sk(i, 1)
        
        If iVid = "Приход" Then sCn = cnZ(i, 1)
        If iVid = "Отгрузка" Then sCn = cnR(i, 1)
        If iVid = "Возврат" Then sCn = cnR(i, 1)
        
End Sub

Private Sub arr_controls_vz()
        On Error Resume Next

        ReDim arr_Ctr_zvk(1 To nom_Cnr)

For i = 1 To iColCtr
            Set arr_Ctr_zvk(i).nNN = frm_ZVK.Controls("nNN" & i)
            Set arr_Ctr_zvk(i).nNm = frm_ZVK.Controls("nNm" & i)
            Set arr_Ctr_zvk(i).nCod = frm_ZVK.Controls("nCod" & i)
            Set arr_Ctr_zvk(i).nEd = frm_ZVK.Controls("nEd" & i)
            Set arr_Ctr_zvk(i).nCn = frm_ZVK.Controls("nCn" & i)
            Set arr_Ctr_zvk(i).nSm = frm_ZVK.Controls("nSm" & i)
            Set arr_Ctr_zvk(i).nCol = frm_ZVK.Controls("nCol" & i)
            Set arr_Ctr_zvk(i).nSk = frm_ZVK.Controls("nSk" & i)
            
            Set arr_Ctr_zvk(i).nCol_vz = frm_ZVK.Controls("nCol_vz" & i)
            Set arr_Ctr_zvk(i).nSk_vz = frm_ZVK.Controls("nSk_vz" & i)
            
        Next
End Sub

Private Sub controls_add()
        On Error Resume Next

        nmCntr = "nNN" & nom_Cnr
        txCntr = sNN
        leftCntr = 0
        gnCntr = fmTextAlignCenter
        widCntr = frm_ZVK.lb_nn.Width + 1
        flag_hidden = True
        Call add_cntrl_nk
        
        nmCntr = "nNm" & nom_Cnr
        txCntr = sNm
        leftCntr = frm_ZVK.lb_nm.Left
        gnCntr = fmTextAlignLeft
        widCntr = frm_ZVK.lb_nm.Width + 1 + 13
        flag_hidden = True
        Call add_cntrl_nk
        
        nmCntr = "nCod" & nom_Cnr
txCntr = sCod
        leftCntr = frm_ZVK.lb_cod.Left
        gnCntr = fmTextAlignLeft
        widCntr = frm_ZVK.lb_cod.Width + 1
        flag_hidden = True
        If ThisWorkbook.Sheets("setting").Range("b6") = 0 Then widCntr = 0
        Call add_cntrl_nk

        nmCntr = "nEd" & nom_Cnr
        txCntr = sEd
        leftCntr = frm_ZVK.lb_ed.Left
        gnCntr = fmTextAlignLeft
        widCntr = frm_ZVK.lb_ed.Width + 1
        flag_hidden = True
        Call add_cntrl_nk
        
        nmCntr = "nCn" & nom_Cnr
        txCntr = VBA.Format(sCn, "0.00")
        leftCntr = frm_ZVK.lb_cn.Left
        gnCntr = fmTextAlignCenter
        widCntr = frm_ZVK.lb_cn.Width + 1
        flag_hidden = True
        If ThisWorkbook.Sheets("setting").Range("b8") = 0 Then widCntr = 0
        Call add_cntrl_nk
        
        nmCntr = "nSm" & nom_Cnr
        txCntr = VBA.Format(sSm, "0.00")
        leftCntr = frm_ZVK.lb_sm.Left
        gnCntr = fmTextAlignCenter
        widCntr = frm_ZVK.lb_sm.Width + 1
        flag_hidden = True
        If ThisWorkbook.Sheets("setting").Range("b8") = 0 Then widCntr = 0
        Call add_cntrl_nk
        
        
        nmCntr = "nCol" & nom_Cnr
        txCntr = sCol
        leftCntr = frm_ZVK.lb_col.Left
        gnCntr = fmTextAlignCenter
        widCntr = frm_ZVK.lb_col.Width + 1
        flag_hidden = True
        Call add_cntrl_nk
        
        nmCntr = "nSk" & nom_Cnr
        txCntr = sSk
        leftCntr = frm_ZVK.lb_sk.Left
        gnCntr = fmTextAlignLeft
        widCntr = frm_ZVK.lb_sk.Width + 1
        flag_hidden = True
        Call add_cntrl_nk
        
        
        nmCntr = "nCol_vz" & nom_Cnr
        txCntr = ""
        leftCntr = frm_ZVK.lb_col_vz.Left
        gnCntr = fmTextAlignCenter
        widCntr = frm_ZVK.lb_col_vz.Width + 1
        flag_hidden = False
        Call add_cntrl_nk_vz
        
        nmCntr = "nSk_vz" & nom_Cnr
        txCntr = sSk
        leftCntr = frm_ZVK.lb_sk_vz.Left
        gnCntr = fmTextAlignLeft
        widCntr = frm_ZVK.lb_sk_vz.Width + 1
        flag_hidden = True
        Call add_cntrl_nk_vz
        
End Sub


Private Sub add_cntrl_nk()
        On Error Resume Next
        With frm_ZVK.Frame_nk.Controls.Add("Forms.TextBox.1")
            .Name = nmCntr
            .text = txCntr
            .Height = hgCntr
            .Top = (nom_Cnr - 1) * (hgCntr - 1): If .Top = 0 Then .Top = 1
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

Private Sub add_cntrl_nk_vz()
        On Error Resume Next
        With frm_ZVK.Frame_nk_vz.Controls.Add("Forms.TextBox.1")
            .Name = nmCntr
            .text = txCntr
            .Height = hgCntr
            .Top = (nom_Cnr - 1) * (hgCntr - 1): If .Top = 0 Then .Top = 1
            .Left = leftCntr
            .Width = widCntr
            .BorderStyle = 1
            .TextAlign = gnCntr
            .Locked = flag_hidden
            .Font.Size = iSize
            .Font.Name = "Times New Roman"
            .MousePointer = 1
            
            
            If nmCntr = "nCol_vz" & nom_Cnr Then
                .MousePointer = 3
                If sCol = 0 Then
                    .Enabled = False
                    frm_ZVK.Controls("nCol_vz" & nom_Cnr).Value = 0
                    frm_ZVK.Controls("nCol_vz" & nom_Cnr).Enabled = False
                End If
            End If
            
        End With
End Sub












Private Sub Frame_height()
        On Error Resume Next
        With frm_ZVK
        
            .Frame_nk.Height = nom_Cnr * (hgCntr - 1)
            If .Frame_nk.Height > .Frame_nk_all.Height Then .ScrollBar1.Width = 12
            .ScrollBar1.Min = 0
            .ScrollBar1.Max = Val(.Frame_nk.Height - .Frame_nk_all.Height)
            
            .Frame_nk_vz.Height = .Frame_nk.Height
            
        End With
End Sub














