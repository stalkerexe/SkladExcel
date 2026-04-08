Attribute VB_Name = "clmns_hidden____"


Public Sub clmns_hidden()
        Call doScreenOff
        Call do_clmns_hidden
        Call doScreenOn
        Unload frm_Set: DoEvents
        MsgBox "   Выполнено!     ", 64, "Опции"
End Sub

Private Sub do_clmns_hidden()
        Call hidden_clm
End Sub

Private Sub hidden_clm()
        On Error Resume Next
        
        If ThisWorkbook.Sheets("setting").Range("b6") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Приход").Cells(1, prCod).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Расход").Cells(1, zvCod).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Отложено_расход").Cells(1, zkCod).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Отложено_приход").Cells(1, pzkCod).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Склад").Cells(1, skCod).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b8") = 1 Then
            flag_hidden = False
        Else
            flag_hidden = True
        End If
        
        With ThisWorkbook.Sheets("Приход")
        .Cells(1, prCnZ).EntireColumn.Hidden = flag_hidden
        .Cells(1, prCnR).EntireColumn.Hidden = flag_hidden
        .Cells(1, prSm).EntireColumn.Hidden = flag_hidden
        End With
        
        With ThisWorkbook.Sheets("Расход")
        .Cells(1, zvCnR).EntireColumn.Hidden = flag_hidden
        .Cells(1, zvSm).EntireColumn.Hidden = flag_hidden
        End With
        
        With ThisWorkbook.Sheets("Отложено_расход")
        .Cells(1, zkCnR).EntireColumn.Hidden = flag_hidden
        .Cells(1, zkSm).EntireColumn.Hidden = flag_hidden
        End With
        
        With ThisWorkbook.Sheets("Отложено_приход")
        .Cells(1, pzkCnZ).EntireColumn.Hidden = flag_hidden
        .Cells(1, pzkSm).EntireColumn.Hidden = flag_hidden
        End With
        
        With ThisWorkbook.Sheets("Склад")
        .Cells(1, skCnZ).EntireColumn.Hidden = flag_hidden
        .Cells(1, skCnR).EntireColumn.Hidden = flag_hidden
        .Cells(1, bxSm).EntireColumn.Hidden = flag_hidden
        End With
        
        If ThisWorkbook.Sheets("setting").Range("b9") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Отложено_расход").Cells(1, zkBr).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Отложено_приход").Cells(1, pzkBr).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Расход").Cells(1, zvBr).EntireColumn.Hidden = flag_hidden
        ThisWorkbook.Sheets("Склад").Cells(1, skBr).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b11") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Склад").Cells(1, skCr).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b35") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Приход").Cells(rwPr_doc, 2).EntireRow.Hidden = flag_hidden
        ThisWorkbook.Sheets("Отложено_приход").Cells(2, pzkOsn).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b40") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Расход").Cells(rwZv_adr, 2).EntireRow.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b41") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets("Расход").Cells(rwZv_tlf, 2).EntireRow.Hidden = flag_hidden

        Call hidden_clm_opl

        Call hidden_clm_skid

End Sub

Private Sub hidden_clm_opl_skid()
        Call hidden_clm_opl
        Call hidden_clm_skid
End Sub

Private Sub hidden_clm_opl()
        On Error Resume Next

        If ThisWorkbook.Sheets("setting").Range("b42") = 1 Then
            flag_hidden = True
        Else
            flag_hidden = False
        End If
        
        With ThisWorkbook.Sheets("Расход")
        
            .Shapes("cmb_oplata").Visible = flag_hidden
            .Shapes("cmb_oplata").Top = .Cells(rwZv_mj, zvSm).Top - 10
            .Shapes("cmb_oplata").Left = .Cells(rwZv_mj, zvSm).Left - .Shapes("cmb_oplata").Width - 4
                    
            With .Cells(rwZv_mj, zvSm)
                .Borders.Color = RGB(127, 127, 127)
                .Borders.LineStyle = flag_hidden
            End With
            
            If flag_hidden = False Then
                .Cells(rwZv_mj, zvSm) = ""
                .Cells(rwZv_mj - 1, zvSm) = ""
            Else
                .Cells(rwZv_mj - 1, zvSm) = "Способ оплаты"
            End If
            
        End With

End Sub

Private Sub hidden_clm_skid()
        On Error Resume Next

        If ThisWorkbook.Sheets("setting").Range("b43") = 1 Then
            flag_hidden = True
        Else
            flag_hidden = False
        End If
        
        With ThisWorkbook.Sheets("Расход")
        
            .Shapes("cmb_skidka").Visible = flag_hidden
            .Shapes("cmb_skidka").Top = .Cells(rwZv_mj, zvOst).Top - 4
            .Shapes("cmb_skidka").Left = .Cells(rwZv_mj, zvOst).Left + .Cells(rwZv_mj, zvOst).Width + 4
            
            With .Cells(rwZv_mj, zvOst)
                .Borders.Color = RGB(127, 127, 127)
                .Borders.LineStyle = flag_hidden
            End With
            
            If flag_hidden = False Then
                .Cells(rwZv_mj, zvOst) = ""
                .Cells(rwZv_mj - 1, zvOst) = ""
            Else
                .Cells(rwZv_mj - 1, zvOst) = "Скидка %"
            End If
            
        End With

End Sub



