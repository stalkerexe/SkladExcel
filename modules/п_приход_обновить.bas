Attribute VB_Name = "п_приход_обновить"
Option Explicit
Dim cell As Range


Public Sub do_obnov_pr()
        On Error Resume Next
        
        ThisWorkbook.Activate: Sheets("Отложено_приход").Activate
        
        Call dann_set
        Call format_
        
        shNm = "Отложено_приход"
        Call ost_sk_zk
        
        Call format_all
        Call remove_green
        
End Sub



Private Sub format_()
        On Error Resume Next
        r7 = ThisWorkbook.Sheets("Отложено_приход").Cells(Rows.Count, pzkNm).End(xlUp).Row: If r7 <= 5 Then Exit Sub
        
        For Each cell In Range(Cells(5, pzkNm), Cells(r7, pzkNm))
            rw = cell.Row: Waite.Label2.Caption = Cells(rw, pzkNm) & "...": DoEvents
            If Cells(rw, 1) <> "" Then
                Call diap_this
                Call format_zk
            End If
        Next
End Sub

Private Sub format_zk()
        Call format_1
        Call format_zg
        Call format_diap_color
        Call format_comm
        Call format_doc
End Sub

Private Sub format_1()

        Range(Cells(row1, pzkNom), Cells(row2 - 1, pzkSk)).VerticalAlignment = xlTop
        
        With Range(Cells(row1 - 1, pzkNom), Cells(row2 + 1, pzkSk))
            .Font.Name = "Times New Roman"
.Font.Size = iRz
        End With
        
        If iWrapText = 1 Then
            With Range(Cells(row1, pzkNm), Cells(row2 - 1, pzkNm))
                .WrapText = True
                .Rows.AutoFit
            End With
        End If
        
        Cells(row1, pzkComm).WrapText = False
        iHg = Cells(row1, pzkComm).RowHeight
        Cells(row1, pzkComm).WrapText = True
        
End Sub

Private Sub format_zg()
        On Error Resume Next
Range(Cells(rw, pzkNom), Cells(rw, pzkComm)).Font.Size = iRz
        Range(Cells(rw, pzkNom), Cells(rw, pzkComm)).Interior.Color = RGB(49, 132, 155)
        Range(Cells(rw, pzkOst), Cells(rw, pzkBr)).Interior.Color = RGB(49, 132, 155)
        Range(Cells(rw, pzkNom), Cells(rw, pzkEnd)).Font.Color = RGB(255, 255, 255)
        Range(Cells(rw, pzkNom), Cells(rw, pzkEnd)).Font.Italic = True
        Range(Cells(rw, pzkNom), Cells(rw, pzkEnd)).Font.Bold = True
        Range(Cells(rw, pzkNom), Cells(rw, pzkDt)).HorizontalAlignment = xlCenter
        Range(Cells(rw, pzkNom), Cells(rw, pzkEnd)).VerticalAlignment = xlCenter
End Sub

Private Sub format_diap_color()
        On Error Resume Next
        For i = row1 - 1 To row2
        
            With Range(Cells(i, pzkNom), Cells(i, pzkEnd)).Borders(xlEdgeBottom)
                .Color = RGB(255, 255, 255)
                .LineStyle = xlContinuous
            End With
            
            With Range(Cells(i, pzkOst), Cells(i, pzkBr)).Borders(xlEdgeBottom)
                .Color = RGB(255, 255, 255)
                .LineStyle = xlContinuous
            End With
            
        Next
        
        With Range(Cells(row1, pzkComm), Cells(row2 - 1, pzkComm))
            .BorderAround Weight:=xlThin
            .Borders.Color = RGB(255, 255, 255)
        End With

        Range(Cells(row1, pzkNN), Cells(row2 - 1, pzkComm)).Interior.Color = RGB(234, 241, 221)
        Range(Cells(row1, pzkOst), Cells(row2 - 1, pzkBr)).Interior.Color = RGB(234, 241, 221)

End Sub

Private Sub format_comm()
        On Error Resume Next
        
        iCol = Application.CountIf(Range(Cells(row1, pzkNm), Cells(row2, pzkNm)), "<>")
        
        Range(Cells(row1, pzkComm), Cells(row2 - 1, pzkComm)).Merge
        Cells(row1, pzkComm).VerticalAlignment = xlTop
        Cells(row1, pzkComm).Font.Size = 8
        
        If iCol = 1 Then
            Cells(row1, pzkComm).RowHeight = iHg
        End If

        If iCol > 1 Then
            Range(Cells(row1, pzkNm), Cells(row2 - 1, pzkNm)).Rows.AutoFit
        End If
        
End Sub

Private Sub format_doc()
        On Error Resume Next
        
        Range(Cells(row1, pzkOsn), Cells(row2 - 1, pzkOsn)).WrapText = True
        Range(Cells(row1, pzkOsn), Cells(row2 - 1, pzkOsn)).Merge
        Cells(row1, pzkOsn).VerticalAlignment = xlTop
        Cells(row1, pzkOsn).Font.Size = 8
        
        If iCol = 1 Then
            Cells(row1, pzkOsn).RowHeight = iHg
        End If

        If iCol > 1 Then
            Range(Cells(row1, pzkNm), Cells(row2 - 1, pzkNm)).Rows.AutoFit
        End If

End Sub

Private Sub format_all()
        On Error Resume Next
        r7 = Cells(Rows.Count, pzkNm).End(xlUp).Row: If r7 <= 5 Then Exit Sub
        Range(Cells(5, pzkNom), Cells(r7, pzkNom)).NumberFormat = "00000"
        Range(Cells(5, pzkDt), Cells(r7, pzkDt)).NumberFormat = "dd.mm.yyyy"
        
        Range(Cells(5, pzkNm), Cells(r7, pzkCod)).HorizontalAlignment = xlLeft
        Range(Cells(5, pzkComm), Cells(r7, pzkComm)).HorizontalAlignment = xlLeft
        Range(Cells(5, pzkNom), Cells(r7, pzkNN)).HorizontalAlignment = xlCenter
        Range(Cells(5, pzkOst), Cells(r7, pzkBr)).HorizontalAlignment = xlCenter
        Range(Cells(5, pzkEd), Cells(r7, pzkSm)).HorizontalAlignment = xlCenter
        
        Range(Cells(5, pzkCnZ), Cells(r7, pzkCnZ)).NumberFormat = "#,##0.00"
        Range(Cells(5, pzkSm), Cells(r7, pzkSm)).NumberFormat = "#,##0.00"
        
        Range(Cells(5, pzkComm), Cells(r7, pzkComm)).NumberFormat = "@"
        
        Range(Cells(5, pzkCod), Cells(r7, pzkCod)).InsertIndent 1
        Range(Cells(5, pzkOsn), Cells(r7, pzkOsn)).InsertIndent 1
        Range(Cells(5, pzkComm), Cells(r7, pzkComm)).InsertIndent 1
        
        Range(Cells(5, pzkNN), Cells(r7, pzkNN)).Font.Size = 9

        Range("a1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        
End Sub





Private Sub diap_this()
    On Error Resume Next
    row1 = rw + 1
    shNm = "Отложено_приход"
    Call find_row2_this
End Sub


