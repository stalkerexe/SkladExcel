Attribute VB_Name = "расход_ред"
Option Explicit


Public Sub zv_pedactirov_()
        On Error Resume Next
        iRow = ActiveCell.Row
        Rows(iRow).Select
        If MsgBox("Редактировать накладную № " & Cells(iRow, zkNom) & ": " & Chr(34) & Cells(iRow, zkNm) & Chr(34) & "?", vbOKCancel + vbQuestion, "Редактировать") = vbCancel Then Exit Sub
        Call do_zv_pedactirov_
End Sub


Private Sub do_zv_pedactirov_()
        Call doScreenOff
        Call do_red
        Call doScreenOn
End Sub

Private Sub do_red()
        On Error Resume Next
        
        Call режим_редактирования_on_pr("Расход")
        
        Call this_row_mk:       Waite.Label2.Caption = "diap_zk_this...": DoEvents
        Call diap_zk_this:      Waite.Label2.Caption = "copy_zk...": DoEvents
        Call copy_zk:           Waite.Label2.Caption = "delete_zk_in_file...": DoEvents
        Call delete_zk_in_file: Waite.Label2.Caption = "clearBf...": DoEvents
        Call clearBf:           Waite.Label2.Caption = "обновить...": DoEvents
        Call copy_to_box:       Waite.Label2.Caption = "завершение...": DoEvents
        
        Erase mk: iOperation = "": iOperation2 = ""
        
        Sheets("Расход").Select
        Range("a1").Select

End Sub

Private Sub this_row_mk()
        iRow = ActiveCell.Row
        marker = Cells(iRow, 1)
        shNm = "Отложено_расход"
End Sub

Private Sub copy_zk()
        Call clear_zv
        Call clear_box
        Call dann_zk_rs
        Call copy_dann
        Call copy_nk
        Call format_zv
End Sub

Private Sub copy_dann()
        With ThisWorkbook.Sheets("Расход")
            .Range("a1") = marker
            .Range("d2") = nomer
            
            .Cells(rwZv_zkz, 4).Value = sZkz
            .Cells(rwZv_adr, 4).Value = sAdr
            .Cells(rwZv_tlf, 4).Value = sTlf
            .Cells(rwZv_mj, 4).Value = sMj
            .Cells(rwZv_dt, 4).Value = sDt
            .Cells(rwZv_dt2, 4).Value = sDt2

.Cells(rwZv_mj, zvSm).Value = iOpl
.Cells(rwZv_mj, zvOst).Value = iSkid
            
            .Cells(1, zvComm) = sComm
            .Cells(rwzvSm, zvSm) = summ
        End With
End Sub

Private Sub copy_nk()

        row1 = row1 + 1
        With ThisWorkbook.Sheets("Отложено_расход")
            Range(Cells(row1, zkNm), Cells(row2, zkSm)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvNm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Range(Cells(row1, zkNN), Cells(row2, zkNN)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvNN).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Range(Cells(row1, zkSk), Cells(row2, zkSk)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Range(Cells(row1, zkCnZ), Cells(row2, zkCnZ)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvCnZ).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Range(Cells(row1, zkCn), Cells(row2, zkCn)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvCn).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Range(Cells(row1, zkOst), Cells(row2, zkOst)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, zvOst).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

            Range(Cells(row1, zkID), Cells(row2, zkID)).Copy
            ThisWorkbook.Sheets("Расход").Cells(rwZv, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

            Application.CutCopyMode = False
        End With
        
End Sub

Private Sub format_zv()
        On Error Resume Next
        row1 = rwZv
        row2 = ThisWorkbook.Sheets("Расход").Cells(Rows.Count, zvNm).End(xlUp).Row
        Call format_zv_
End Sub




Private Sub clear_zv()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            r24 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r24 + 44).EntireRow.Delete
            .Cells(rwzvSm, zvSm) = ""
            .Range("a1") = ""
            .Range("d2") = ""
        End With
End Sub

Private Sub copy_to_box()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("Расход")
            r24 = .Cells(Rows.Count, zvNm).End(xlUp).Row
            .Range(.Cells(rwZv, 1), .Cells(r24, 100)).Copy
            ThisWorkbook.Sheets("корзина").Cells(rwZv, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        End With
        
        With ThisWorkbook.Sheets("корзина")
            For i = rwZv To r24
                iRowBox = i
                Call formula_in_box
            Next
        End With
        
        Call sum_box

End Sub
