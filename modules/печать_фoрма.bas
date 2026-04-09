Attribute VB_Name = "печать_фoрма"
Option Explicit

Public Sub printZv()
        On Error Resume Next
        Call unload_mn_vid
        DoEvents
        r7 = Cells(Rows.Count, zvNm).End(xlUp).Row
        If r7 < rwZv Then
        MsgBox "Нет позиций для печати в накладной!", 64, "Печать"
        Exit Sub
        End If
        frm_print.Show
        frm_print.TextBox1.Text = 1
End Sub

Public Sub printPr()
        On Error Resume Next
        Call unload_mn_vid_pr
        DoEvents
        r7 = Cells(Rows.Count, prNm).End(xlUp).Row
        If r7 < rwZv Then
        MsgBox "Нет позиций для печати в приходе!", 64, "Печать"
        Exit Sub
        End If
        frm_print.Show
        frm_print.TextBox1.Text = 2
End Sub

Public Sub printZk()
        On Error Resume Next
        With frm_print
            .Show
            .TextBox1.Text = 3
            .tb_row.value = ActiveCell.Row
        End With
End Sub

Public Sub printZk_pr()
        On Error Resume Next
        With frm_print
            .Show
            .TextBox1.Text = 4
            .tb_row.value = ActiveCell.Row
        End With
End Sub

Public Sub printSk()
        On Error Resume Next
        frm_print.Show
        frm_print.TextBox1.Text = 7
End Sub

Public Sub printZVK()
        On Error Resume Next
        With frm_print
            .Show
            .TextBox1.Text = 5
            .StartUpPosition = 0
            .Top = frm_ZVK.Top + frm_ZVK.OK_menu.Top
            .Left = frm_ZVK.Left + frm_ZVK.OK_menu.Left - .Width
        End With
End Sub


