Attribute VB_Name = "склад_смотреть"



Public Sub sklad_show()
        Call doScreenOff
        Call do_sklad_show
        Call doScreenOn
End Sub

Public Sub do_sklad_show()
    On Error Resume Next
    
    Sheets("Склад").Select
    
    flag_open = 0:       Waite.Label2.Caption = "clear_this": DoEvents
    
    Call clear_this:     Waite.Label2.Caption = "перебрать_sk": DoEvents
    
    Call перебрать_sk:   Waite.Label2.Caption = "format_": DoEvents
    
    Call format_:        Waite.Label2.Caption = "Завершение...": DoEvents
    
    Call clearBf

    Call shapes_left

    Application.ErrorCheckingOptions.BackgroundChecking = False
    Range("a1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
End Sub


Private Sub перебрать_sk()

On Error Resume Next
    
    For n = LBound(sk) To UBound(sk)
        If sk(n, 1) <> "" Then
        sSk = sk(n, 1): Waite.Label2.Caption = sSk: DoEvents
        If sSk <> "" Then
            
            Call arr_select_sk
            Call resize_sk_bf
            Call parse_sk
            
        End If
        End If
33
    Next
    
End Sub

Private Sub resize_sk_bf()
    On Error Resume Next

    Call clearBf

    With ThisWorkbook.Sheets("буфер")
        .Cells(1, "a").Resize(UBound(c), 9) = c
    End With

    With ThisWorkbook.Sheets("буфер")
        r7 = .Cells(Rows.Count, skNm).End(xlUp).Row + 2
        gr_sk = .Range(.Cells(1, skGr), .Cells(r7, skGr)).Value
        gr = .Range(.Cells(1, 2), .Cells(r7, 2)).Value
        cod_sk = .Range(.Cells(1, skCod), .Cells(r7, skCod)).Value
        nm_sk = .Range(.Cells(1, skNm), .Cells(r7, skNm)).Value
        ed_sk = .Range(.Cells(1, skEd), .Cells(r7, skEd)).Value
        cnZ_sk = .Range(.Cells(1, "g"), .Cells(r7, "g")).Value
        cnR_sk = .Range(.Cells(1, "h"), .Cells(r7, "h")).Value
        col_sk = .Range(.Cells(1, "f"), .Cells(r7, "f")).Value
        br_sk = .Range(.Cells(1, "i"), .Cells(r7, "i")).Value
    End With

End Sub

Private Sub parse_sk()
    On Error Resume Next

        With ThisWorkbook.Sheets("Склад")
            r7 = .Cells(Rows.Count, skNm).End(xlUp).Row + 2: If r7 <= 6 Then r7 = 6
            .Cells(r7, 1) = 1

            With .Cells(r7, skNm)
                .Value = sSk
                .Font.Bold = True
                .Font.Size = 16
                .Font.ColorIndex = 3
            End With

            With Range(.Cells(r7, skNm), .Cells(r7, skComm))
                .Merge
                .HorizontalAlignment = xlLeft
            End With
        End With

        row1 = r7 + 1

        With ThisWorkbook.Sheets("Склад")
            Range(.Cells(row1, skNm), .Cells(row1 + UBound(c), skNm)).NumberFormat = "@"
        End With

   
    With ThisWorkbook.Sheets("Склад")
        .Cells(row1, skGr).Resize(UBound(nm_sk), 1) = gr_sk
        .Cells(row1, 2).Resize(UBound(nm_sk), 1) = gr
        .Cells(row1, skCod).Resize(UBound(nm_sk), 1) = cod_sk
        .Cells(row1, skNm).Resize(UBound(nm_sk), 1) = nm_sk
        .Cells(row1, skEd).Resize(UBound(nm_sk), 1) = ed_sk
        .Cells(row1, skCnZ).Resize(UBound(nm_sk), 1) = cnZ_sk
        .Cells(row1, skCnR).Resize(UBound(nm_sk), 1) = cnR_sk
        .Cells(row1, skOst).Resize(UBound(nm_sk), 1) = col_sk
        .Cells(row1, skBr).Resize(UBound(nm_sk), 1) = br_sk
    End With
    
    With ThisWorkbook.Sheets("Склад")
        row2 = .Cells(Rows.Count, skNm).End(xlUp).Row
        Range(.Cells(row1, skSk), .Cells(row2, skSk)) = sSk
    End With


End Sub



Private Sub copy_()
On Error Resume Next

    With ThisWorkbook.Sheets("Склад")
        r = .Cells(Rows.Count, skNm).End(xlUp).Row + 2: If r <= 5 Then r = 5
            .Cells(r, 1) = 1
            With .Cells(r, skNm)
                .Value = "склад   " & sSk
                .Font.Bold = True
                .Font.Size = 16
                .Font.ColorIndex = 3
            End With
    End With


r = r + 1
With ThisWorkbook.Sheets("Склад")
r9 = Cells(Rows.Count, skNm).End(xlUp).Row + 2
Range(Cells(5, 1), Cells(r9, skComm)).Copy
.Cells(r, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
End With

End Sub


Private Sub format_()
        On Error Resume Next

        r7 = Cells(Rows.Count, skNm).End(xlUp).Row
        iCol = Application.CountIf(Range(Cells(5, skNm), Cells(r7 + 3, skNm)), "<>")
        If iCol = 0 Then Exit Sub

        With Range(Cells(5, skNm), Cells(r7, skNm))
            .WrapText = True
            .Rows.AutoFit
        End With

        For i = 5 To r7
            If Cells(i, skGr) <> "" Then
                Cells(i, skNm).Font.Bold = True
                Cells(i, skNm).Font.Size = 12

                With Range(Cells(i, skNm), Cells(i, skComm))
                    .WrapText = False
                    .RowHeight = 18
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                End With

            End If
        Next

        Range(Cells(5, skEd), Cells(r7, skCr)).HorizontalAlignment = xlCenter

        With Range(Cells(5, skCod), Cells(r7, skCod))
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With

        With Range(Cells(5, skComm), Cells(r7, skComm))
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Font.Size = 9
        End With

        Range(Cells(5, skSk), Cells(r7, skSk)).Font.Size = 9

        Range(Cells(5, skCnZ), Cells(r7, skCnZ)).NumberFormat = "#,##0.00"
        Range(Cells(5, skCnR), Cells(r7, skCnR)).NumberFormat = "#,##0.00"
        Range(Cells(5, skEd), Cells(r7, skCr)).HorizontalAlignment = xlCenter

        Call zebra

        For i = 5 To r7
        If Cells(i, skNm) <> "" And Cells(i, skGr) = "" Then
        If Cells(i, skOst) < Cells(i, skCr) Then
        Cells(i, skOst).Interior.Color = RGB(230, 185, 184)
        End If
        End If
        Next

        Range(Cells(4, skCod), Cells(r7, skComm)).AutoFilter
End Sub

Private Sub zebra()
On Error Resume Next
For i = 6 To r7 Step 2
Range(Cells(i, skCod), Cells(i, skComm)).Interior.Color = RGB(216, 216, 216)
Next
End Sub

Private Sub clear_this()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
Call AutoFilter_delete
r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
.Range("a" & 5 & ":a" & r7 + 44).EntireRow.Delete
End With
End Sub

Public Sub shapes_left()
On Error Resume Next
With ThisWorkbook.Sheets("Склад").Shapes("grCmbBox")
    .Left = Range("m3").Left - .Width + 5
End With
End Sub
