Attribute VB_Name = "расход_save"
Option Explicit

Public Sub svZvk()

        Call unload_mn_vid: DoEvents
        With ThisWorkbook.Sheets("Расход")
        r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row
        If r7 < rwZv Then
        MsgBox "          Нет позиций в накладной!" & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "Нажмите кнопку <Добавить позицию> и двойным кликом выберите позиции", 64, ""
        Exit Sub
        End If
        End With
        
        sZkz = Cells(rwZv_zkz, 4).Value
        sDt = Cells(rwZv_dt, 4).Value
        
        If MsgBox("     Отложить накладную?               " & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "   Кому: " & sZkz & VBA.Chr(10) & _
        "   Дата: " & sDt, vbOKCancel + vbQuestion, "Расход") = vbCancel Then Exit Sub
        
        Call doScreenOff
        Call do_sv
        Call doScreenOn
End Sub

Private Sub do_sv()
        On Error Resume Next
        Call this_row_mk
        Call dann_zv:           Waite.Label2.Caption = "Копирование данных...": DoEvents
        Call copy_to_zkz:       Waite.Label2.Caption = "clear_this...": DoEvents
        Call clear_this:        Waite.Label2.Caption = "обновить_склад...": DoEvents
        
        Call clear_peremenn
        Call erase_arr_zv
        Call erase_arr_sk
        Call do_obnov
        Erase mk: iOperation = "": iOperation2 = ""
End Sub

Private Sub this_row_mk()
        marker = "c" & VBA.Now
        iOperation = "sv_zv"
End Sub

Private Sub copy_to_zkz()
        On Error Resume Next
        Call copy_to_zkz_dann
        Call copy_to_zkz_nk
End Sub

Private Sub copy_to_zkz_dann()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_расход")
            n7 = .Cells(Rows.Count, zkNm).End(xlUp).Row + 2: If n7 < 5 Then n7 = 5
            .Cells(n7, 1) = marker
            .Cells(n7, zkNom).Value = nomer
            .Cells(n7, zkDt1).Value = sDt
            .Cells(n7, zkDt2).Value = sDt2
            .Cells(n7, zkSm).Value = summ
            .Cells(n7, zkZkz).Value = sZkz
            .Cells(n7, zkAdr).Value = sAdr
            .Cells(n7, zkTlf).NumberFormat = "@"
            .Cells(n7, zkTlf).Value = sTlf
            .Cells(n7, zkMj).Value = sMj
            .Cells(n7 + 1, zkComm) = sComm
            
            .Cells(n7, zkOpl).Value = iOpl
            .Cells(n7, zkComm).Value = iOpl
            
            .Cells(n7, zkSkid).Value = iSkid
            If iSkid > 0 Then
                .Cells(n7, zkCnR).NumberFormat = "@"
                .Cells(n7, zkCnR).Value = iSkid & "%"
            End If

            
            Dim sZkzL As String
            If VBA.Len(sZkz) > 33 Then sZkz = VBA.Left(sZkz, 33)
            If VBA.Len(sAdr) > 33 Then sAdr = VBA.Left(sAdr, 33)

            If ThisWorkbook.Sheets("setting").Range("b40") = 0 Then sAdr = ""
            If ThisWorkbook.Sheets("setting").Range("b41") = 0 Then sTlf = ""
            
            .Cells(n7, zkNm) = sZkz & "   " & sAdr & "   " & sTlf & "   " & sMj
            
        End With
End Sub


Private Sub copy_to_zkz_nk()
        On Error Resume Next
        n7 = n7 + 1
        With ThisWorkbook.Sheets("Расход")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row
            .Range(.Cells(rwZv, zvNm), .Cells(r7, zvSm)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkNm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, zvNN), .Cells(r7, zvNN)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkNN).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, zvSk), .Cells(r7, zvSk)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range(.Cells(rwZv, zvCnZ), .Cells(r7, zvCnZ)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkCnZ).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
.Range(.Cells(rwZv, zvCn), .Cells(r7, zvCn)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkCn).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
.Range(.Cells(rwZv, 1), .Cells(r7, 1)).Copy
            ThisWorkbook.Sheets("Отложено_расход").Cells(n7, zkID).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Application.CutCopyMode = False
        End With
End Sub




Private Sub clear_this()
        On Error Resume Next

        With ThisWorkbook.Sheets("Расход")
            If .Cells(9, zvOst) = "" Then
                Call nom_nk(2)
                .Range("d2") = nomer
            End If
        End With

        Call режим_редактирования_off_pr("Расход")


        With ThisWorkbook.Sheets("Расход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r7 + 44).EntireRow.Delete
            .Cells(rwzvSm, zvSm) = ""
            .Range("a1") = ""
            .Cells(1, zvComm) = ""
            
.Cells(rwZv_mj, zvOst) = ""
.Cells(rwZv_mj, zvSm) = ""
        End With
        
        Call clear_box
End Sub
