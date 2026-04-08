Attribute VB_Name = "расход_удалить"
Option Explicit


Public Sub delete_zv()
        On Error Resume Next
        iRow = ActiveCell.Row
        Range(Cells(iRow, 3), Cells(iRow, 12)).Select
        If MsgBox("Удалить заказ № " & Cells(iRow, zkNom) & ": " & Chr(34) & Cells(iRow, zkNm) & Chr(34) & "?", vbOKCancel + vbQuestion, "Удаление") = vbCancel Then Exit Sub
        Call zvDelete
End Sub


Private Sub zvDelete()
        Call doScreenOff
        Call do_delete
        Call doScreenOn
End Sub

Private Sub do_delete()
        On Error Resume Next
        
        Call this_row_mk:              Waite.Label2.Caption = "delete_zk_in_file...": DoEvents
        Call delete_zk_in_file:        Waite.Label2.Caption = "обновить реестр...": DoEvents
        
        Call erase_arr_zk_this
        Erase mk: iOperation = "": iOperation2 = ""
        
End Sub

Private Sub this_row_mk()
        iRow = ActiveCell.Row
        marker = Cells(iRow, 1)
        shNm = "Отложено_расход"
        iOperation = "br_delete_zk"
End Sub


