Attribute VB_Name = "архив_load"
Option Explicit
Dim sZg As String



Public Sub load_nk_from_arh()
        On Error Resume Next
        
        Call dann_arh
        If ind = -1 Then Exit Sub
        If marker = "" Then Exit Sub
        
        Call to_form
        Call doScreenOn
        
        Erase c
        DoEvents
        If row1 = 0 Then MsgBox "Эта накладная не найдена в архиве", 64, "Накладная"
End Sub

Private Sub dann_arh()
On Error Resume Next
        With zvSelect
            iGod = .comb_year.Value
            iPapka = .comb_vid.Value
            iVid = .comb_vid.Value
        End With

        With zvSelect.ListBox1
            ind = .ListIndex
            If ind = -1 Then Exit Sub
            marker = .List(.ListIndex, 0)
            sDt = .List(.ListIndex, 3)
        End With
        
        Call find_path_vid
        
        
        Call find_zg
                
End Sub


Private Sub find_zg()
        If iVid = "Приход" Then sZg = "Приходная накладная"
        If iVid = "Отгрузка" Then sZg = "Расходная накладная"
        If iVid = "Возврат" Then sZg = "Накладная возврата"
        If iVid = "Перемещение" Then sZg = "Накладная перемещения"
End Sub


Private Sub to_form()
        On Error Resume Next

        Call find_mk_arh
        If row1 = 0 Then GoTo 99
        
        Call dann
        Call arr_arh
99
        Call form_show
        Call добавить_контролы_vz
        
        If ThisWorkbook.Sheets("setting").Range("f12") = 1 Then
            frm_ZVK.CheckBox_perenos.Value = True
        End If
        
End Sub

Private Sub form_show()
        On Error Resume Next

        Waite.Label2.Caption = "Загрузка данных...":  DoEvents
        
        With frm_ZVK
            .Show
            .lb_doc.Visible = True
            .tb_ind.Text = ind
            .tb_date.Text = sDt
            .tb_what.Text = iVid
            .tb_year.Text = iGod
            .Top = zvSelect.Top + 15
            .Left = zvSelect.Left
        End With
        
        With frm_ZVK
            .lb_vid_nk.Caption = sZg & " №"
            .tb_nomer.Caption = VBA.Format(nomer, "00000")
            .tb_Zkz.Text = sZkz
            .tb_Mnj.Text = sMj
            .tb_Dt.Text = sDt
            .tb_doc.Text = sOsn: If sOsn = "" Then .lb_doc.Visible = False
            .tb_mk.Text = marker
            .tb_sm.Text = VBA.Format(summ, "#,##0.00")
        End With
        
        If iVid = "Отгрузка" Then
            With frm_ZVK
                .lb_doc.Visible = False
                .tb_doc.Text = ""
            End With
        End If
        
        Call hidden_clm_zvk
        
End Sub

Private Sub dann()
        On Error Resume Next
        
        iRow = row1
        
        If iVid = "Приход" Then Call dann_arh_pr
        If iVid = "Отгрузка" Then Call dann_arh_rs
        If iVid = "Возврат" Then Call dann_arh_vz
        
End Sub


Private Sub arr_arh()
        On Error Resume Next
        
        iRow = row1
        
        If iVid = "Приход" Then Call arr_arh_pr
        If iVid = "Отгрузка" Then Call arr_arh_rs
        
End Sub


