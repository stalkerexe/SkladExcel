Attribute VB_Name = "расход_обновить"
Option Explicit
Dim cell As Range

Public Sub do_obnov()
        On Error Resume Next
        
        ThisWorkbook.Activate: Sheets("Отложено_расход").Activate
        
        Call dann_set
        Call format_
        
        shNm = "Отложено_расход"
        Call ost_sk_zk
        
        Call format_all
        Call remove_green
        
End Sub



Private Sub format_()
        On Error Resume Next
        r7 = ThisWorkbook.Sheets("Отложено_расход").Cells(Rows.Count, zkNm).End(xlUp).Row: If r7 <= 5 Then Exit Sub
        
        For Each cell In Range(Cells(5, zkNm), Cells(r7, zkNm))
            rw = cell.Row: Waite.Label2.Caption = Cells(rw, zkNm) & "...": DoEvents
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
End Sub

Private Sub format_1()
        On Error Resume Next

        Range(Cells(row1, zkNom), Cells(row2 - 1, zkSk)).VerticalAlignment = xlTop
        
        With Range(Cells(row1 - 1, zkNom), Cells(row2 + 1, zkSk))
            .Font.Name = "Times New Roman"
.Font.Size = iRz
        End With
        
        If iWrapText = 1 Then
            With Range(Cells(row1, zkNm), Cells(row2 - 1, zkNm))
                .WrapText = True
                .Rows.AutoFit
            End With
        End If
        
        Cells(row1, zkComm).WrapText = False
        iHg = Cells(row1, zkComm).RowHeight
        Cells(row1, zkComm).WrapText = True
        
End Sub

Private Sub format_zg()
        On Error Resume Next
Range(Cells(rw, zkNom), Cells(rw, zkComm)).Font.Size = iRz
        Range(Cells(rw, zkNom), Cells(rw, zkComm)).Interior.Color = RGB(79, 129, 189)
        Range(Cells(rw, zkOst), Cells(rw, zkBr)).Interior.Color = RGB(79, 129, 189)
        Range(Cells(rw, zkNom), Cells(rw, zkEnd)).Font.Color = RGB(255, 255, 255)
        Range(Cells(rw, zkNom), Cells(rw, zkEnd)).Font.Italic = True
        Range(Cells(rw, zkNom), Cells(rw, zkEnd)).Font.Bold = True
        Range(Cells(rw, zkNom), Cells(rw, zkDt2)).HorizontalAlignment = xlCenter
        Range(Cells(rw, zkNom), Cells(rw, zkEnd)).VerticalAlignment = xlCenter
End Sub

Private Sub format_diap_color()
        On Error Resume Next

        For i = row1 - 1 To row2
            With Range(Cells(i, zkNom), Cells(i, zkEnd)).Borders(xlEdgeBottom)
                .Color = RGB(255, 255, 255)
                .LineStyle = xlContinuous
            End With
            With Range(Cells(i, zkOst), Cells(i, zkBr)).Borders(xlEdgeBottom)
                .Color = RGB(255, 255, 255)
                .LineStyle = xlContinuous
            End With
        Next
        
        With Range(Cells(row1, zkComm), Cells(row2 - 1, zkComm))
            .BorderAround Weight:=xlThin
            .Borders.Color = RGB(255, 255, 255)
        End With

        Range(Cells(row1, zkNN), Cells(row2 - 1, zkComm)).Interior.Color = RGB(234, 241, 221)
        Range(Cells(row1, zkOst), Cells(row2 - 1, zkBr)).Interior.Color = RGB(234, 241, 221)

End Sub

Private Sub format_comm()
        On Error Resume Next
        
        iCol = Application.CountIf(Range(Cells(row1, zkNm), Cells(row2, zkNm)), "<>")
        
        Range(Cells(row1, zkComm), Cells(row2 - 1, zkComm)).Merge
        Cells(row1, zkComm).VerticalAlignment = xlTop
        Cells(row1, zkComm).Font.Size = 8
        
        If iCol = 1 Then
            Cells(row1, zkComm).RowHeight = iHg
        End If

        If iCol > 1 Then
            Range(Cells(row1, zkNm), Cells(row2 - 1, zkNm)).Rows.AutoFit
        End If
        
End Sub

Private Sub format_all()
        On Error Resume Next
        r7 = Cells(Rows.Count, zkNm).End(xlUp).Row: If r7 <= 5 Then Exit Sub
        Range(Cells(5, zkNom), Cells(r7, zkNom)).NumberFormat = "00000"
        Range(Cells(5, zkDt1), Cells(r7, zkDt2)).NumberFormat = "dd.mm.yyyy"
        Range(Cells(5, zkNom), Cells(r7, zkNN)).HorizontalAlignment = xlCenter
        Range(Cells(5, zkCnR), Cells(r7, zkCnR)).NumberFormat = "#,##0.00"
        Range(Cells(5, zkSm), Cells(r7, zkSm)).NumberFormat = "#,##0.00"
        Range(Cells(5, zkComm), Cells(r7, zkComm)).NumberFormat = "@"
        Range(Cells(5, zkNom), Cells(r7, zkSk)).Font.Name = "Times New Roman"
        Range(Cells(5, zkOst), Cells(r7, zkBr)).HorizontalAlignment = xlCenter
        Range(Cells(5, zkNm), Cells(r7, zkNm)).HorizontalAlignment = xlLeft
        Range(Cells(5, zkComm), Cells(r7, zkComm)).HorizontalAlignment = xlLeft
        Range(Cells(5, zkEd), Cells(r7, zkSm)).HorizontalAlignment = xlCenter
        Range(Cells(5, zkCod), Cells(r7, zkCod)).InsertIndent 1
        Range(Cells(5, zkComm), Cells(r7, zkComm)).InsertIndent 1
        
        Range(Cells(5, zkNN), Cells(r7, zkNN)).Font.Size = 9

        Range("a1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
End Sub

Public Sub dann_set()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("setting")
        
            iRz = .Range("o24").Value
            
            iWrapText = .Range("o25").Value
            
        End With
End Sub

Private Sub diap_this()
    On Error Resume Next
    
    row1 = rw + 1
    
    shNm = "Отложено_расход"
    Call find_row2_this
    
End Sub



