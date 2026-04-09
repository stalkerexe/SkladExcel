Attribute VB_Name = "п_приход_ред"
Option Explicit


Public Sub zv_pedactirov_pr()
On Error Resume Next
iRow = ActiveCell.Row
Range(Cells(iRow, 3), Cells(iRow, 12)).Select
If MsgBox("Редактировать накладную № " & Cells(iRow, pzkNom) & ": " & Chr(34) & Cells(iRow, pzkNm) & Chr(34) & "?", vbOKCancel + vbQuestion, "Редактировать") = vbCancel Then Exit Sub
Call zvRed_pr
End Sub




Private Sub zvRed_pr()
        Call doScreenOff
        Call do_red
        Call doScreenOn
End Sub

Private Sub do_red()
        On Error Resume Next
        
        Call режим_редактирования_on_pr("Приход")
        
        Call this_row_mk:    Waite.Label2.Caption = "diap_zk_this...": DoEvents
        Call diap_zk_this:   Waite.Label2.Caption = "бронь_удалить...": DoEvents
        Call clear_pr:       Waite.Label2.Caption = "copy_zk...": DoEvents
        Call copy_zk:        Waite.Label2.Caption = "clearBf...": DoEvents
        Call delete_pr_in_file: Waite.Label2.Caption = "clearBf...": DoEvents
        Call clearBf:        Waite.Label2.Caption = "обновить...": DoEvents
        Erase mk: iOperation = "": iOperation2 = ""
        
        Sheets("Приход").Select
        Range("a1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1

End Sub

Private Sub this_row_mk()
        iRow = ActiveCell.Row
        marker = Cells(iRow, 1)
        shNm = "Отложено_приход"
End Sub

Private Sub copy_zk()
        Call clear_pr
        Call clear_box
        Call dann_zk_pr
        Call copy_dann
        Call copy_nk
        Call format_pr
End Sub

Private Sub copy_dann()
        On Error Resume Next
        With ThisWorkbook.Sheets("Приход")
            .Range("a1") = marker
            .Range("d2") = nomer
            
            .Cells(rwPr_zkz, 4).value = sZkz
            .Cells(rwPr_mj, 4).value = sMj
            .Cells(rwPr_doc, 4).value = sOsn
            .Cells(rwPr_dt, 4).value = sDt
            
            .Cells(rwzvSm, prSm) = summ
            .Cells(1, prComm) = sComm
            
            .Cells(1, prDoc) = sDoc
            .Cells(1, prDocN) = sDocN
            .Cells(1, prDocDt) = sDocDt
            
        End With
End Sub

Private Sub copy_nk()
        On Error Resume Next

        row1 = row1 + 1
        
        With ThisWorkbook.Sheets("Отложено_приход")
        
            .Range(Cells(row1, pzkNm), Cells(row2, pzkCnZ)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prNm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(Cells(row1, pzkSm), Cells(row2, pzkSm)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prSm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(Cells(row1, pzkCnR), Cells(row2, pzkCnR)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prCnR).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(Cells(row1, pzkNN), Cells(row2, pzkNN)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prNN).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(Cells(row1, pzkSk), Cells(row2, pzkSk)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(Cells(row1, pzkGr), Cells(row2, pzkGr)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, prGr).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            .Range(Cells(row1, pzkID), Cells(row2, pzkID)).Copy
            ThisWorkbook.Sheets("Приход").Cells(rwZv, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Application.CutCopyMode = False

        End With
        
End Sub

Private Sub format_pr()
        On Error Resume Next
        row1 = rwZv
        row2 = ThisWorkbook.Sheets("Приход").Cells(Rows.Count, prNm).End(xlUp).Row
        Call format_pr_
End Sub






Private Sub clear_pr()
On Error Resume Next
With ThisWorkbook.Sheets("Приход")
r24 = .UsedRange.Rows.Count + .UsedRange.Row - 1
.Range("a" & rwZv & ":a" & r24 + 44).EntireRow.Delete
.Range("a1") = ""
.Range("d2") = ""
End With
End Sub





