Attribute VB_Name = "печать_zvk"
Option Explicit


Public Sub print_ZVK_arh()
        Call doScreenOff
        Call do_print_ZVK
        Call doScreenOn
End Sub

Private Sub do_print_ZVK()
        On Error Resume Next
        Call dann_zvk
Call find_path_vid
        Call open_arh_file
        Call copy_to_blank
        Call do_print_paper
        ThisWorkbook.Sheets(nmBlank).Visible = 2
End Sub

Private Sub open_arh_file()
        On Error Resume Next

        Call find_mk_arh
        If row1 = 0 Then
            fg_open_arh = 1
            Exit Sub
        End If
        
        Call dann_arh

End Sub

Private Sub copy_to_blank()
        If iVid = "Приход" Then Call copy_to_blank_pr
        If iVid = "Отгрузка" Then Call copy_to_blank_rs
End Sub

Private Sub do_print_paper()
        Call print_paper
End Sub

Private Sub dann_arh()
        On Error Resume Next
        iRow = row1
        
        If iVid = "Приход" Then
            Call dann_arh_pr
            Call arr_arh_pr
        End If
        
        If iVid = "Отгрузка" Then
            Call dann_arh_rs
            Call arr_arh_rs
        End If
        
End Sub


      

