Attribute VB_Name = "расход_отгрузка_zk"
Option Explicit


Public Sub otgr_zk()
        On Error Resume Next
        iRow = ActiveCell.Row
        Rows(iRow).Select
        If MsgBox("Отгрузить заказ № " & Cells(iRow, zkNom) & ": " & Chr(34) & Cells(iRow, zkNm) & Chr(34) & "?", vbOKCancel + vbQuestion, "Отгрузка") = vbCancel Then Exit Sub
        Call do_otgr_zk
End Sub

Private Sub do_otgr_zk()
        Call doScreenOff
        Call do_otg
        Call doScreenOn
End Sub

Private Sub do_otg()
        On Error Resume Next
        Call this_row_mk:         Waite.Label2.Caption = "diap_zk_this...": DoEvents
        Call diap_zk_this:        Waite.Label2.Caption = "doOst...": DoEvents
        Call doOst:               Waite.Label2.Caption = "сохранить_накладную...": DoEvents
        Call svOtg:               Waite.Label2.Caption = "delete_zk_in_file...": DoEvents
        Call delete_zk_in_file:   Waite.Label2.Caption = "обновить_склад...": DoEvents
        
        Call do_sklad_obnovitt:   Waite.Label2.Caption = "обновить_данные...": DoEvents
        Erase mk: iOperation = "": iOperation2 = ""
End Sub

Private Sub this_row_mk()
        iRow = ActiveCell.Row
        marker = Cells(iRow, 1)
        shNm = "Отложено_расход"
End Sub



Private Sub doOst()
        iOperation = "zv"
        row1 = row1 + 1
        Call arr_zk_this
        Call ost_skds
End Sub

Private Sub svOtg()
        iVid = "ot"
        Call dann
        Call save_nk
End Sub



Private Sub dann()
        On Error Resume Next
        iRow = row1 - 1
        Call dann_zk_rs
End Sub



