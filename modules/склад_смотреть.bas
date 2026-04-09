Attribute VB_Name = "склад_смотреть"
Option Explicit

Public Sub sklad_show()
    Call doScreenOff
    Call do_sklad_show
    Call doScreenOn
End Sub

Public Sub do_sklad_show()
    On Error GoTo ErrHandler

    Dim wsSklad As Worksheet
    If Not RequireSheet(SHEET_SKLAD, wsSklad, "do_sklad_show") Then Exit Sub

    wsSklad.Select

    flag_open = 0:       Waite.Label2.Caption = "clear_this": DoEvents

    Call clear_this:     Waite.Label2.Caption = "перебрать_sk": DoEvents

    Call перебрать_sk:   Waite.Label2.Caption = "format_": DoEvents

    Call format_:        Waite.Label2.Caption = "Завершение...": DoEvents

    Call clearBf

    Call shapes_left

    Application.ErrorCheckingOptions.BackgroundChecking = False
    wsSklad.Range("a1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Exit Sub

ErrHandler:
    ReportVbaError "do_sklad_show", Err.Number, Err.Description, "Склад"
End Sub


Private Sub перебрать_sk()
On Error GoTo ErrHandler

    For n = LBound(sk) To UBound(sk)
        If sk(n, 1) <> "" Then
            sSk = sk(n, 1): Waite.Label2.Caption = sSk: DoEvents
            If sSk <> "" Then
                Call arr_select_sk
                Call resize_sk_bf
                Call parse_sk
            End If
        End If
    Next

    Exit Sub
ErrHandler:
    ReportVbaError "перебрать_sk", Err.Number, Err.Description, "Склад"
End Sub

Private Sub resize_sk_bf()
    On Error GoTo ErrHandler

    Dim wsBuffer As Worksheet
    If Not RequireSheet("буфер", wsBuffer, "resize_sk_bf") Then Exit Sub

    Call clearBf

    With wsBuffer
        .Cells(1, "a").Resize(UBound(c), 9) = c
    End With

    With wsBuffer
        r7 = .Cells(.Rows.Count, skNm).End(xlUp).Row + 2
        gr_sk = .Range(.Cells(1, skGr), .Cells(r7, skGr)).value
        gr = .Range(.Cells(1, 2), .Cells(r7, 2)).value
        cod_sk = .Range(.Cells(1, skCod), .Cells(r7, skCod)).value
        nm_sk = .Range(.Cells(1, skNm), .Cells(r7, skNm)).value
        ed_sk = .Range(.Cells(1, skEd), .Cells(r7, skEd)).value
        cnZ_sk = .Range(.Cells(1, "g"), .Cells(r7, "g")).value
        cnR_sk = .Range(.Cells(1, "h"), .Cells(r7, "h")).value
        col_sk = .Range(.Cells(1, "f"), .Cells(r7, "f")).value
        br_sk = .Range(.Cells(1, "i"), .Cells(r7, "i")).value
    End With

    Exit Sub
ErrHandler:
    ReportVbaError "resize_sk_bf", Err.Number, Err.Description, "Склад"
End Sub

Private Sub parse_sk()
    On Error GoTo ErrHandler

    Dim wsSklad As Worksheet
    If Not RequireSheet(SHEET_SKLAD, wsSklad, "parse_sk") Then Exit Sub

    With wsSklad
        r7 = .Cells(.Rows.Count, skNm).End(xlUp).Row + 2: If r7 <= 6 Then r7 = 6
        .Cells(r7, 1) = 1

        With .Cells(r7, skNm)
            .value = sSk
            .Font.Bold = True
            .Font.Size = 16
            .Font.ColorIndex = 3
        End With

        With .Range(.Cells(r7, skNm), .Cells(r7, skComm))
            .Merge
            .HorizontalAlignment = xlLeft
        End With
    End With

    row1 = r7 + 1

    With wsSklad
        .Range(.Cells(row1, skNm), .Cells(row1 + UBound(c), skNm)).NumberFormat = "@"
    End With

    With wsSklad
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

    With wsSklad
        row2 = .Cells(.Rows.Count, skNm).End(xlUp).Row
        .Range(.Cells(row1, skSk), .Cells(row2, skSk)) = sSk
    End With

    Exit Sub
ErrHandler:
    ReportVbaError "parse_sk", Err.Number, Err.Description, "Склад"
End Sub



Private Sub copy_()
On Error GoTo ErrHandler

    Dim wsSklad As Worksheet
    If Not RequireSheet(SHEET_SKLAD, wsSklad, "copy_") Then Exit Sub

    With wsSklad
        r = .Cells(.Rows.Count, skNm).End(xlUp).Row + 2: If r <= 5 Then r = 5
        .Cells(r, 1) = 1
        With .Cells(r, skNm)
            .value = "склад   " & sSk
            .Font.Bold = True
            .Font.Size = 16
            .Font.ColorIndex = 3
        End With
    End With

    r = r + 1
    With wsSklad
        r9 = .Cells(.Rows.Count, skNm).End(xlUp).Row + 2
        .Range(.Cells(5, 1), .Cells(r9, skComm)).Copy
        .Cells(r, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    Exit Sub
ErrHandler:
    ReportVbaError "copy_", Err.Number, Err.Description, "Склад"
End Sub


Private Sub format_()
    On Error GoTo ErrHandler

    Dim wsSklad As Worksheet
    If Not RequireSheet(SHEET_SKLAD, wsSklad, "format_") Then Exit Sub

    With wsSklad
        r7 = .Cells(.Rows.Count, skNm).End(xlUp).Row
        iCol = Application.CountIf(.Range(.Cells(5, skNm), .Cells(r7 + 3, skNm)), "<>")
        If iCol = 0 Then Exit Sub

        With .Range(.Cells(5, skNm), .Cells(r7, skNm))
            .WrapText = True
            .Rows.AutoFit
        End With

        For i = 5 To r7
            If .Cells(i, skGr) <> "" Then
                .Cells(i, skNm).Font.Bold = True
                .Cells(i, skNm).Font.Size = 12

                With .Range(.Cells(i, skNm), .Cells(i, skComm))
                    .WrapText = False
                    .RowHeight = 18
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                End With
            End If
        Next

        .Range(.Cells(5, skEd), .Cells(r7, skCr)).HorizontalAlignment = xlCenter

        With .Range(.Cells(5, skCod), .Cells(r7, skCod))
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With

        With .Range(.Cells(5, skComm), .Cells(r7, skComm))
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Font.Size = 9
        End With

        .Range(.Cells(5, skSk), .Cells(r7, skSk)).Font.Size = 9

        .Range(.Cells(5, skCnZ), .Cells(r7, skCnZ)).NumberFormat = "#,##0.00"
        .Range(.Cells(5, skCnR), .Cells(r7, skCnR)).NumberFormat = "#,##0.00"
        .Range(.Cells(5, skEd), .Cells(r7, skCr)).HorizontalAlignment = xlCenter

        Call zebra

        For i = 5 To r7
            If .Cells(i, skNm) <> "" And .Cells(i, skGr) = "" Then
                If .Cells(i, skOst) < .Cells(i, skCr) Then
                    .Cells(i, skOst).Interior.Color = RGB(230, 185, 184)
                End If
            End If
        Next

        .Range(.Cells(4, skCod), .Cells(r7, skComm)).AutoFilter
    End With

    Exit Sub
ErrHandler:
    ReportVbaError "format_", Err.Number, Err.Description, "Склад"
End Sub

Private Sub zebra()
On Error GoTo ErrHandler

Dim wsSklad As Worksheet
If Not RequireSheet(SHEET_SKLAD, wsSklad, "zebra") Then Exit Sub

Dim i As Long
For i = 6 To r7 Step 2
    wsSklad.Range(wsSklad.Cells(i, skCod), wsSklad.Cells(i, skComm)).Interior.Color = RGB(216, 216, 216)
Next

Exit Sub
ErrHandler:
ReportVbaError "zebra", Err.Number, Err.Description, "Склад"
End Sub

Private Sub clear_this()
On Error GoTo ErrHandler

Dim wsSklad As Worksheet
If Not RequireSheet(SHEET_SKLAD, wsSklad, "clear_this") Then Exit Sub

With wsSklad
    Call AutoFilter_delete
    r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
    .Range("a" & 5 & ":a" & r7 + 44).EntireRow.Delete
End With

Exit Sub
ErrHandler:
ReportVbaError "clear_this", Err.Number, Err.Description, "Склад"
End Sub

Public Sub shapes_left()
On Error Resume Next
With ThisWorkbook.Sheets("Склад").Shapes("grCmbBox")
    .Left = Range("m3").Left - .Width + 5
End With
End Sub
