Attribute VB_Name = "расход_отгрузка_zv"
Option Explicit


Public Sub otgZv()
        Call unload_mn_vid: DoEvents
        With Sheets("Расход")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row
            If r7 < rwZv Then
                MsgBox "          Нет позиций в накладной!" & VBA.Chr(10) & _
                "---------------------------------------------------" & VBA.Chr(10) & _
                "Нажмите кнопку <Добавить позицию> и двойным кликом выберите позиции", 64, "Расход"
                Exit Sub
            End If
        End With
        
        sZkz = Cells(rwZv_zkz, 4).Value
        sDt = Cells(rwZv_dt, 4).Value
        
        If MsgBox("     Отгрузить накладную?               " & VBA.Chr(10) & _
        "---------------------------------------------------" & VBA.Chr(10) & _
        "   Кому: " & sZkz & VBA.Chr(10) & _
        "   Дата:     " & sDt, vbOKCancel + vbQuestion, "Расход") = vbCancel Then Exit Sub
        
        Call doScreenOff
        Call do_otg
        Call doScreenOn
End Sub

Private Sub do_otg()

        Call doOst:        Waite.Label2.Caption = "сохранить накладную...": DoEvents
        Call svOtg:        Waite.Label2.Caption = "очистить данные...": DoEvents
        
        Call clearBf:      Waite.Label2.Caption = "clear_this...": DoEvents
        Call clear_this:   Waite.Label2.Caption = "обновить_склад...": DoEvents
        
        Call do_sklad_obnovitt:   Waite.Label2.Caption = "завершение...": DoEvents
        
        Call erase_arr_zv
        Call erase_arr_sk
        Erase mk: iOperation = "": iOperation2 = ""
        
        Sheets("Расход").Select
        
End Sub

Private Sub dann()
        On Error Resume Next
        marker = "c" & VBA.Now
        Call dann_zv
End Sub



Private Sub doOst()
        iOperation = "zv"
        Call arr_zv
        Call ost_skds
End Sub


 
 

Private Sub svOtg()
        On Error Resume Next
        iVid = "ot"
        Call dann
        Call save_nk
End Sub



Private Sub clear_this()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("Расход")
            If .Cells(9, zvOst) <> "Режим_редактирования" Then
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
.Cells(9, zvOst) = ""
            
.Cells(rwZv_mj, zvOst) = ""
.Cells(rwZv_mj, zvSm) = ""
        End With

        Call clear_box
End Sub



