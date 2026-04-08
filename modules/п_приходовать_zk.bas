Attribute VB_Name = "п_приходовать_zk"
Option Explicit



Public Sub prZk()
        On Error Resume Next
        iRow = ActiveCell.Row
        Range(Cells(iRow, 3), Cells(iRow, 12)).Select
        If MsgBox("Оприходовать накладную № " & Cells(iRow, zkNom) & ": " & Chr(34) & Cells(iRow, zkNm) & Chr(34) & "?", vbOKCancel + vbQuestion, "Приход") = vbCancel Then Exit Sub
        Call prZk_do
End Sub

Private Sub prZk_do()
        Call doScreenOff
        Call do_otg
        Call doScreenOn
End Sub

Private Sub do_otg()
        On Error Resume Next
        
        Call this_row_mk:   Waite.Label2.Caption = "diap_zk_this...": DoEvents
        Call diap_zk_this:  Waite.Label2.Caption = "doOst...": DoEvents
        Call doOst:         Waite.Label2.Caption = "сохранить_накладную...": DoEvents
        Call svPr:          Waite.Label2.Caption = "delete_zk_in_file...": DoEvents
        Call delete_zk_in_file:  Waite.Label2.Caption = "clearBf...": DoEvents
        Call clearBf:       Waite.Label2.Caption = "обновить данные...": DoEvents

        Call do_sklad_obnovitt:   Waite.Label2.Caption = "обновить_данные...": DoEvents

        Erase mk: iOperation = "": iOperation2 = ""
        
End Sub

Private Sub this_row_mk()
        iRow = ActiveCell.Row
        marker = Cells(iRow, 1)
        shNm = "Отложено_приход"
        iOperation = "zk_pr"
End Sub



Private Sub doOst()
        iOperation = "pr"
        row1 = iRow + 1
        Call arr_zk_this_pr
        Call ost_skds
End Sub



Private Sub svPr()
        iVid = "pr"
        Call dann
        Call save_nk
End Sub


Private Sub dann()
        On Error Resume Next
        iRow = row1 - 1
        Call dann_zk_pr
End Sub





