Attribute VB_Name = "п_____________________________"
Option Explicit


Public Sub format_pr_()
        On Error Resume Next

        With ThisWorkbook.Sheets("Приход")
        
            Range(.Cells(row1, prNm), .Cells(row2, prSm)).Borders.LineStyle = True
            Range(.Cells(row1, prEd), .Cells(row2, prSm)).HorizontalAlignment = xlCenter
            Range(.Cells(row1, prCod), .Cells(row2, prCod)).IndentLevel = 1
    
            Range(.Cells(row1, prCod), .Cells(row2, prCod)).NumberFormat = "@"
    
            Range(.Cells(row1, prCnZ), .Cells(row2, prCnZ)).NumberFormat = "#,##0.00"
            Range(.Cells(row1, prCnR), .Cells(row2, prCnR)).NumberFormat = "#,##0.00"
            Range(.Cells(row1, prSm), .Cells(row2, prSm)).NumberFormat = "#,##0.00"
    
            With Range(.Cells(row1, prNm), .Cells(row2, prNm))
                .WrapText = True
                .Rows.AutoFit
            End With
    
            With Range(.Cells(row1, 2), .Cells(row2, prSk))
                .Font.Name = "Times New Roman"
                .Font.Size = 11
            End With
    
            With Range(.Cells(row1, prNN), .Cells(row2, prNN))
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With
        
            With Range(.Cells(row1, prSk), .Cells(row2, prSk))
                .Font.Name = "Times New Roman"
                .Borders.LineStyle = True
                .IndentLevel = 1
                .Font.Size = 9
            End With
        
        
        End With

        Call remove_green
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1

End Sub


Public Sub format_zv_()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("Расход")
        
            Range(.Cells(row1, zvNm), .Cells(row2, zvSm)).Borders.LineStyle = True
            
            Range(.Cells(row1, zvEd), .Cells(row2, zvSm)).HorizontalAlignment = xlCenter
            
            Range(.Cells(row1, zvCnR), .Cells(row2, zvCnR)).NumberFormat = "#,##0.00"
            Range(.Cells(row1, zvSm), .Cells(row2, zvSm)).NumberFormat = "#,##0.00"
            Range(.Cells(row1, zvCod), .Cells(row2, zvCod)).IndentLevel = 1
            
            With Range(.Cells(row1, zvOst), .Cells(row2, zvBr))
                .Borders.LineStyle = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            With Range(.Cells(row1, zvNm), .Cells(row2, zvNm))
                .WrapText = True
                .Rows.AutoFit
            End With
            
            With Range(.Cells(row1, 2), .Cells(row2, zvBr))
                .Font.Name = "Times New Roman"
                .Font.Size = 11
            End With
       
            With Range(.Cells(row1, zvNN), .Cells(row2, zvNN))
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With
                   
        End With
       
        Range("a1").Select
        Call remove_green
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
       
End Sub

