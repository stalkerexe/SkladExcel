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

        sZkz = Cells(rwZv_zkz, 4).value
        sDt = Cells(rwZv_dt, 4).value

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

        ' ИСПРАВЛЕНИЕ #9: усечение строк теперь происходит ДО записи в ячейку.
        ' В оригинале усечение было ПОСЛЕ .Cells(n7, zkNm) = sZkz & ...,
        ' то есть в ячейку попадало полное (неусечённое) значение.
        If VBA.Len(sZkz) > 33 Then sZkz = VBA.Left(sZkz, 33)
        If VBA.Len(sAdr) > 33 Then sAdr = VBA.Left(sAdr, 33)

        With ThisWorkbook.Sheets("Отложено_расход")
            n7 = .Cells(Rows.Count, zkNm).End(xlUp).Row + 2: If n7 < 5 Then n7 = 5
            .Cells(n7, 1) = marker
            .Cells(n7, zkNom).value = nomer
            .Cells(n7, zkDt1).value = sDt
            .Cells(n7, zkDt2).value = sDt2
            .Cells(n7, zkSm).value = summ
            .Cells(n7, zkZkz).value = sZkz
            .Cells(n7, zkAdr).value = sAdr
            .Cells(n7, zkTlf).NumberFormat = "@"
            .Cells(n7, zkTlf).value = sTlf
            .Cells(n7, zkMj).value = sMj
            .Cells(n7 + 1, zkComm) = sComm

            .Cells(n7, zkOpl).value = iOpl
            .Cells(n7, zkComm).value = iOpl

            .Cells(n7, zkSkid).value = iSkid
            If iSkid > 0 Then
                .Cells(n7, zkCnR).NumberFormat = "@"
                .Cells(n7, zkCnR).value = iSkid & "%"
            End If

            If ThisWorkbook.Sheets("setting").Range("b40") = 0 Then sAdr = ""
            If ThisWorkbook.Sheets("setting").Range("b41") = 0 Then sTlf = ""

            .Cells(n7, zkNm) = sZkz & "   " & sAdr & "   " & sTlf & "   " & sMj

        End With
End Sub


Private Sub copy_to_zkz_nk()
        On Error Resume Next

        ' ИСПРАВЛЕНИЕ #10: в оригинале Cells(...) внутри With ThisWorkbook.Sheets("Отложено_расход")
        ' использовался без точки — обращение шло к ActiveSheet, а не к листу "Расход".
        ' При активном листе отличном от "Расход" данные копировались из неверного места.
        ' Теперь все диапазоны-источники явно обращаются к листу "Расход".

        Dim wsRash As Worksheet
        Dim wsZk   As Worksheet
        Set wsRash = ThisWorkbook.Sheets("Расход")
        Set wsZk = ThisWorkbook.Sheets("Отложено_расход")

        n7 = n7 + 1
        r7 = wsRash.Cells(wsRash.Rows.Count, zvNm).End(xlUp).Row

        wsRash.Range(wsRash.Cells(rwZv, zvNm), wsRash.Cells(r7, zvSm)).Copy
        wsZk.Cells(n7, zkNm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        wsRash.Range(wsRash.Cells(rwZv, zvNN), wsRash.Cells(r7, zvNN)).Copy
        wsZk.Cells(n7, zkNN).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        wsRash.Range(wsRash.Cells(rwZv, zvSk), wsRash.Cells(r7, zvSk)).Copy
        wsZk.Cells(n7, zkSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        wsRash.Range(wsRash.Cells(rwZv, zvCnZ), wsRash.Cells(r7, zvCnZ)).Copy
        wsZk.Cells(n7, zkCnZ).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        wsRash.Range(wsRash.Cells(rwZv, zvCn), wsRash.Cells(r7, zvCn)).Copy
        wsZk.Cells(n7, zkCn).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        wsRash.Range(wsRash.Cells(rwZv, 1), wsRash.Cells(r7, 1)).Copy
        wsZk.Cells(n7, zkID).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Application.CutCopyMode = False
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


