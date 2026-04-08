Attribute VB_Name = "п_приходовать_zv"
Option Explicit


Public Sub prihod_()

        Call unload_mn_vid_pr: DoEvents
        
        With Sheets("Приход")
        r7 = .Cells(Rows.Count, prNm).End(xlUp).Row
        If r7 < rwZv Then
        MsgBox "          Нет позиций в накладной!" & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "Нажмите кнопку <Добавить позицию> и двойным кликом выберите позиции", 64, "Приход"
        Exit Sub
        End If
        End With
        
        sZkz = Cells(rwPr_zkz, 4).Value
        sDt = VBA.CDate(Cells(rwPr_dt, 4).Value)
        
        If MsgBox("     Приходовать накладную?               " & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "   Контрагент: " & sZkz & VBA.Chr(10) & _
        "   Дата:     " & sDt, vbOKCancel + vbQuestion, "Приход") = vbCancel Then Exit Sub
        
        Call doScreenOff
        Call do_otg
        Call doScreenOn
End Sub

Private Sub do_otg()

        Call doOst:        Waite.Label2.Caption = "сохранить накладную...": DoEvents
        Call svPr:         Waite.Label2.Caption = "очистить данные...": DoEvents

        Call clearBf:      Waite.Label2.Caption = "clear_this...": DoEvents
        Call clear_this:   Waite.Label2.Caption = "завершение...": DoEvents

        Call do_sklad_obnovitt:   Waite.Label2.Caption = "обновить_данные...": DoEvents

        Call erase_arr_zv
        Call erase_arr_sk
        
        Sheets("Приход").Select
End Sub

Private Sub dann()
        On Error Resume Next
        marker = "c" & VBA.Now
        Call dann_pr
End Sub



Private Sub doOst()
        iOperation = "pr"
        Call arr_pr
        Call ost_skds
End Sub
 
 

Private Sub svPr()
        On Error Resume Next
        iVid = "pr"
        Call dann
        Call save_nk
End Sub


Private Sub clear_this()
        On Error Resume Next

        With ThisWorkbook.Sheets("Приход")
            If .Cells(9, zvOst) <> "Режим_редактирования" Then
                Call nom_nk(3)
                .Range("d2") = nomer
            End If
        End With
        
        Call режим_редактирования_off_pr("Приход")


        With ThisWorkbook.Sheets("Приход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r7 + 44).EntireRow.Delete
            .Range("a1") = ""
            .Cells(1, prComm) = ""
.Cells(9, zvOst) = ""
        End With
        
End Sub





