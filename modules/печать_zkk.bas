Attribute VB_Name = "печать_zkk"
Option Explicit

Private Const shRow As Integer = 13
Private Const cmEnd = 8




Public Sub prnt_zk()
        Call doScreenOff
        Call do_blank
        Call doScreenOn
End Sub

Private Sub do_blank()
        Call do_vid
        Call diap_this
        Call copy_to_blank
        Call print_paper
        ThisWorkbook.Sheets(nmBlank).Visible = 2
End Sub

Private Sub do_vid()
        If frm_print.TextBox1.Text = 3 Then iVid = "Отгрузка": shNm = "Отложено_расход"
        If frm_print.TextBox1.Text = 4 Then iVid = "Приход": shNm = "Отложено_приход"
End Sub


Private Sub copy_to_blank()
        On Error Resume Next
        
        If iVid = "Приход" Then
            Call dann_zk_pr
            Call arr_zk_this_pr
            Call copy_to_blank_pr
        End If
        
        If iVid = "Отгрузка" Then
            Call dann_zk_rs
            Call arr_zk_this
            Call copy_to_blank_rs
        End If
        
End Sub


Private Sub diap_this()
    On Error Resume Next
    
    iRow = Val(frm_print.tb_row.Value)
    row1 = iRow + 1
    
    Call find_row2_this
    
End Sub

