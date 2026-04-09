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
        If frm_print.TextBox1.Text = 4 Then iVid = "Приход":   shNm = "Отложено_приход"
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


' ИСПРАВЛЕНИЕ #12: добавлена валидация iRow.
' В оригинале Val(frm_print.tb_row.Value) не проверялось на 0.
' При пустом tb_row iRow = 0, row1 = 1 — процедура пыталась
' распечатать строку заголовка как накладную без ошибки.
' Теперь при iRow < 5 показываем понятное сообщение и выходим.
Private Sub diap_this()
    On Error Resume Next

    iRow = Val(frm_print.tb_row.value)

    If iRow < 5 Then
        MsgBox "Не выбрана накладная для печати!" & Chr(10) & _
               "Дважды щёлкните по строке накладной перед нажатием кнопки печати.", _
               48, "Печать"
        Exit Sub
    End If

    row1 = iRow + 1

    Call find_row2_this

End Sub


