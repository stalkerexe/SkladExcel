Attribute VB_Name = "п_приход_save"
Option Explicit

Public Sub svPrihod()

        Call unload_mn_vid_pr: DoEvents

        With ThisWorkbook.Sheets("Приход")
        r7 = .Cells(Rows.Count, prNm).End(xlUp).Row
        If r7 < rwZv Then
        MsgBox "          Нет позиций в накладной!" & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "Нажмите кнопку <Добавить позицию> и двойным кликом выберите позиции", 64, ""
        Exit Sub
        End If
        End With
        
        sZkz = Cells(rwPr_zkz, 4).Value
        sDt = VBA.CDate(Cells(rwPr_dt, 4).Value)
            
        If MsgBox("     Отложить накладную?               " & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "   Контрагент: " & sZkz & VBA.Chr(10) & _
        "   Дата: " & sDt, vbOKCancel + vbQuestion, "Приход") = vbCancel Then Exit Sub
        
        Call doScreenOff
        Call do_sv
        Call doScreenOn
End Sub

Private Sub do_sv()

        Call dann:            Waite.Label2.Caption = "Копирование данных...": DoEvents
        Call copy_to_zkz:     Waite.Label2.Caption = "clear_this...": DoEvents
        Call clear_this:      Waite.Label2.Caption = "обновить...": DoEvents
        Call erase_arr_zv
        Call erase_arr_sk

        Call do_obnov_pr
End Sub

Private Sub dann()
        On Error Resume Next
        marker = "c" & VBA.Now
        Call dann_pr
End Sub

Private Sub copy_to_zkz()
        On Error Resume Next
        Call copy_to_zkz_dann
        Call copy_to_zkz_nk
End Sub

Private Sub copy_to_zkz_dann()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_приход")
            n7 = .Cells(Rows.Count, pzkNm).End(xlUp).Row + 2: If n7 < 5 Then n7 = 5
            .Cells(n7, 1) = marker
            .Cells(n7, pzkNom) = nomer
            .Cells(n7, pzkDt) = sDt
            .Cells(n7, pzkPsv) = sZkz
            .Cells(n7, pzkMj) = sMj
            .Cells(n7, pzkNm) = sZkz
            .Cells(n7, pzkCol) = sMj
            .Cells(n7, pzkSm) = summ
            
            .Cells(n7, pzkDoc) = sDoc
            .Cells(n7, pzkDocN) = "'" & sDocN
            .Cells(n7, pzkDocDt) = sDocDt
            
            .Cells(n7 + 1, pzkOsn) = sOsn
            .Cells(n7 + 1, pzkComm) = sComm
            
            .Cells(n7 + 1, pzkOsn).WrapText = False

        End With
End Sub

Private Sub copy_to_zkz_nk()
        On Error Resume Next
        n7 = n7 + 1
        With ThisWorkbook.Sheets("Приход")
            r7 = .Cells(Rows.Count, prNm).End(xlUp).Row
            
            .Range(.Cells(rwZv, prNm), .Cells(r7, prCnZ)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkNm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, prSm), .Cells(r7, prSm)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkSm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, prNN), .Cells(r7, prNN)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkNN).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, prSk), .Cells(r7, prSk)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, prCnR), .Cells(r7, prCnR)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkCnR).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            .Range(.Cells(rwZv, prGr), .Cells(r7, prGr)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkGr).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            .Range(.Cells(rwZv, 1), .Cells(r7, 1)).Copy
            ThisWorkbook.Sheets("Отложено_приход").Cells(n7, pzkID).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Application.CutCopyMode = False
            
        End With
End Sub





Private Sub clear_this()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("Приход")
            If .Cells(9, zvOst) = "" Then
                Call nom_nk(3)
                .Range("d2") = nomer
            End If
        End With

        Call режим_редактирования_off_pr("Приход")
        
        
        With ThisWorkbook.Sheets("Приход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r7 + 44).EntireRow.Delete
            .Range("a1").Value = ""
            
            .Cells(rwzvSm, prSm).Value = ""
            .Cells(rwPr_doc, 4).Value = ""
            .Cells(rwPr_doc, prCol).Value = ""
            
            .Cells(1, prDoc).Value = ""
            .Cells(1, prDocN).Value = ""
            .Cells(1, prComm).Value = ""
        End With
        
        Call clear_box

End Sub


