Attribute VB_Name = "о_otchet_format"
Option Explicit

Dim arr_wid()
Dim sZg As String
Dim iMin
Dim iMax

Private Const vNom As Integer = 2
Private Const vDt As Integer = 3
Private Const vNm As Integer = 4
Private Const vCod As Integer = 5
Private Const vCol As Integer = 6

Dim vCn As Integer: Dim vCnZ As Integer: Dim vCnR As Integer
Dim vSm As Integer: Dim vSmZ As Integer: Dim vSmR As Integer
Dim vMj As Integer: Dim vZkz As Integer: Dim vDoc As Integer
Dim vSk As Integer: Dim vOpl As Integer: Dim vSkid As Integer: Dim vPR As Integer
Dim vEnd As Integer


Public Sub format_otchet()
        Call cm_column
        Call format_otchet_all
        If iVid = "pr" Then Call format_otchet_pr
        If iVid = "ot" Then Call format_otchet_ot
End Sub

Private Sub cm_column()

        If iVid = "pr" Then
            vCn = 7
            vSm = 8
            vMj = 9
            vZkz = 10
            vDoc = 11
            vEnd = 11
        End If

        If iVid = "ot" Then
            
            vCnR = 7
            vSmR = 8
            vCnZ = 9
            vSmZ = 10
            
            vPR = 11
            vMj = 12
            vZkz = 13
            vSk = 14
            vOpl = 15
            vSkid = 16
            vEnd = 16
        End If

End Sub



Private Sub format_otchet_ot()
        Call format_otchet_do_ot
        Call format_otchet_zg_ot
End Sub

Private Sub format_otchet_do_ot()
        On Error Resume Next
        
        With wbOt.ActiveSheet
        
            r7 = .Cells(Rows.Count, vNm).End(xlUp).Row
            
            Range(.Cells(5, vNom), .Cells(r7, vEnd)).Borders.LineStyle = True
            Range(.Cells(5, vCol), .Cells(r7, vPR)).HorizontalAlignment = xlCenter
            Range(.Cells(5, vCnR), .Cells(r7, vPR)).NumberFormat = "#,##0.00"
            
            Range(.Cells(5, vCnR), .Cells(r7, vSmR)).Font.ColorIndex = 3
            Range(.Cells(5, vCnZ), .Cells(r7, vSmZ)).Font.ColorIndex = 10
            
            With Range(.Cells(5, vSk), .Cells(r7, vSk))
                .Font.Size = 9
                .IndentLevel = 1
            End With
            
            With Range(.Cells(5, vOpl), .Cells(r7, vOpl))
                .Font.Size = 9
                .IndentLevel = 1
            End With
            
            With Range(.Cells(5, vSkid), .Cells(r7, vSkid))
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With
            
        End With
        
End Sub

Private Sub format_otchet_zg_ot()

        If ThisWorkbook.Sheets("setting").Range("b8").value = 0 Then GoTo 33
        With wbOt.ActiveSheet
            
            r7 = .Cells(Rows.Count, vNm).End(xlUp).Row
            
            .Cells(2, vSmR).Formula = "=SUBTOTAL(9,h5:h" & r7 + 1 & ")"
            
            .Cells(2, vSmZ).Formula = "=SUBTOTAL(9,j5:j" & r7 + 1 & ")"
        
            .Cells(2, vPR).Formula = "=h2-j2"
            
            With Range(.Cells(2, "h"), .Cells(2, "k"))
                .Font.Name = "Times New Roman"
                .Font.Size = 10
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .NumberFormat = "#,##0.00"
            End With
        
        End With
33

        ReDim arr_wid(1 To iCmOt + 1, 2)
        
        arr_wid(1, 1) = 1:   arr_wid(1, 2) = ""
        arr_wid(2, 1) = 8:   arr_wid(2, 2) = "Номер"
        arr_wid(3, 1) = 10:  arr_wid(3, 2) = "Дата"
        arr_wid(4, 1) = 39:  arr_wid(4, 2) = "Наименование"
        arr_wid(5, 1) = 11:  arr_wid(5, 2) = "Артикул"
        arr_wid(6, 1) = 9:   arr_wid(6, 2) = "Кол - во"
        arr_wid(7, 1) = 11:  arr_wid(7, 2) = "Цена продажа"
        arr_wid(8, 1) = 15:  arr_wid(8, 2) = "Сумма продажа"
        arr_wid(9, 1) = 13:  arr_wid(9, 2) = "Цена закуп"
        arr_wid(10, 1) = 14:  arr_wid(10, 2) = "Сумма закуп"
        arr_wid(11, 1) = 16:  arr_wid(11, 2) = "Прибыль"
        arr_wid(12, 1) = 17:  arr_wid(12, 2) = "Сотрудник"
        arr_wid(13, 1) = 23: arr_wid(13, 2) = "Получатель"
        arr_wid(14, 1) = 12: arr_wid(14, 2) = "Склад"
        arr_wid(15, 1) = 12: arr_wid(15, 2) = "Способ оплаты"
        arr_wid(16, 1) = 9: arr_wid(16, 2) = "Скидка %"
        
        For cm = 1 To iCmOt + 1
            With wbOt.ActiveSheet
                .Cells(4, cm).value = arr_wid(cm, 2): .Cells(4, cm).ColumnWidth = 7: .Cells(4, cm).WrapText = True
                .Cells(4, cm).ColumnWidth = CDbl(arr_wid(cm, 1))
            End With
        Next

        With wbOt.ActiveSheet
            With Range(.Cells(4, 2), .Cells(4, iCmOt + 1))
                .Font.Name = "Times New Roman"
                .Font.Size = 10
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = True
                .Font.Bold = True
                .RowHeight = 30
                .Interior.Color = RGB(234, 241, 221)
            End With
        End With
        
        Call find_zg
        Call min_max_data
        
        With wbOt.ActiveSheet
            With .Cells(2, 4)
                .value = sZg
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .Font.Bold = True
            End With
            
            
            str_ = VBA.Format(iMin, "dd.mm.yyyy") & " - " & VBA.Format(iMax, "dd.mm.yyyy")
            With .Cells(3, 4)
                .value = str_
                .Font.Name = "Times New Roman"
                .Font.Size = 9
                .VerticalAlignment = xlTop
            End With
            
        End With
        

End Sub




Private Sub format_otchet_pr()
        Call format_otchet_do_pr
        Call format_otchet_zg_pr
End Sub

Private Sub format_otchet_zg_pr()

        ReDim arr_wid(1 To 11, 2)
        
        arr_wid(1, 1) = 2:   arr_wid(1, 2) = ""
        arr_wid(2, 1) = 9:   arr_wid(2, 2) = "Номер"
        arr_wid(3, 1) = 12:  arr_wid(3, 2) = "Дата"
        arr_wid(4, 1) = 42:  arr_wid(4, 2) = "Наименование"
        arr_wid(5, 1) = 12:  arr_wid(5, 2) = "Артикул"
        arr_wid(6, 1) = 9:   arr_wid(6, 2) = "Кол - во"
        arr_wid(7, 1) = 11:  arr_wid(7, 2) = "Цена"
        arr_wid(8, 1) = 15:  arr_wid(8, 2) = "Сумма"
        arr_wid(9, 1) = 17:  arr_wid(9, 2) = "Сотрудник"
        arr_wid(10, 1) = 25: arr_wid(10, 2) = "Поставщик"
        arr_wid(11, 1) = 25: arr_wid(11, 2) = "Документ"
        
        For cm = 1 To iCmOt + 1
            With wbOt.ActiveSheet
                .Cells(4, cm).ColumnWidth = CDbl(arr_wid(cm, 1))
                .Cells(4, cm).value = arr_wid(cm, 2)
            End With
        Next

        With wbOt.ActiveSheet
            With Range(.Cells(4, 2), .Cells(4, iCmOt + 1))
                .Font.Name = "Times New Roman"
                .Font.Size = 10
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = True
                .Font.Bold = True
                .RowHeight = 30
                .Interior.Color = RGB(242, 221, 221)
            End With
        End With
        
        Call find_zg
        Call min_max_data
        
        With wbOt.ActiveSheet
            With .Cells(2, 4)
                .value = sZg
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .Font.Bold = True
            End With
            
            str_ = VBA.Format(iMin, "dd.mm.yyyy") & " - " & VBA.Format(iMax, "dd.mm.yyyy")
            With .Cells(3, 4)
                .value = str_
                .Font.Name = "Times New Roman"
                .Font.Size = 9
                .VerticalAlignment = xlTop
            End With
            
        End With

End Sub

Private Sub format_otchet_do_pr()
        On Error Resume Next

        With wbOt.ActiveSheet
        
            r7 = .Cells(Rows.Count, vNm).End(xlUp).Row
            
            Range(.Cells(5, vNom), .Cells(r7, vEnd)).Borders.LineStyle = True
            
            Range(.Cells(5, vDoc), .Cells(r7, vDoc)).IndentLevel = 1
            Range(.Cells(5, vDoc), .Cells(r7, vDoc)).Font.Size = 9
            
            Range(.Cells(5, vCol), .Cells(r7, vSm)).HorizontalAlignment = xlCenter
            
            Range(.Cells(5, vCn), .Cells(r7, vSm)).NumberFormat = "#,##0.00"
            
        End With
        
End Sub



Private Sub format_otchet_all()
        On Error Resume Next
        
        Call row_zzz
    
        With wbOt.ActiveSheet
        
            r7 = .Cells(Rows.Count, vNm).End(xlUp).Row
            
            With Range(.Cells(5, vNom), .Cells(r7, vEnd))
                .Font.Name = "Times New Roman"
                .Font.Size = 10
                .VerticalAlignment = xlTop
            End With
            
            Range(.Cells(5, vNm), .Cells(r7, vNm)).IndentLevel = 1
            Range(.Cells(5, vCod), .Cells(r7, vCod)).IndentLevel = 1
            Range(.Cells(5, vMj), .Cells(r7, vMj)).IndentLevel = 1
            Range(.Cells(5, vZkz), .Cells(r7, vZkz)).IndentLevel = 1

            Range(.Cells(5, vZkz), .Cells(r7, vZkz)).Font.Size = 10
            Range(.Cells(5, vMj), .Cells(r7, vMj)).Font.Size = 10
            
            Range(.Cells(5, vNom), .Cells(r7, vDt)).HorizontalAlignment = xlCenter
            Range(.Cells(5, vNom), .Cells(r7, vNom)).NumberFormat = "00000"
            Range(.Cells(5, vDt), .Cells(r7, vDt)).NumberFormat = "dd.mm.yyyy"
            
            Range(.Cells(4, vNom), .Cells(r7, vEnd)).AutoFilter
        
        End With
        
        ActiveWindow.ScrollRow = 1
        Call remove_green
        
        ActiveWindow.DisplayGridlines = False
        
        ActiveWindow.Zoom = 90
        
End Sub

Private Sub row_zzz()
        On Error Resume Next

        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        With wbOt.ActiveSheet
            .Cells(1, 2).EntireRow.Insert
            .Cells(1, 2).EntireRow.Insert
            .Cells(1, 2).EntireRow.Insert
        End With
        
End Sub

Private Sub min_max_data()
        On Error Resume Next
        With wbOt.ActiveSheet
            r7 = .Cells(Rows.Count, vNm).End(xlUp).Row
            iMin = Application.Min(Range(.Cells(5, vDt), .Cells(r7, vDt)))
            iMax = Application.Max(Range(.Cells(5, vDt), .Cells(r7, vDt)))
        End With
End Sub

Private Sub find_zg()
        If iVid = "pr" Then sZg = "ЗАКУПКА ЗА ПЕРИОД"
        If iVid = "ot" Then sZg = "РЕАЛИЗОВАНО ЗА ПЕРИОД"
End Sub
