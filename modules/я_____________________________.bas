Attribute VB_Name = "я_____________________________"
Option Explicit
Dim rEnd As Long



Public Sub find_row2_this()
        On Error Resume Next
        
        Call find_rEnd
        
        With ThisWorkbook.Sheets(shNm)
            mk = Range(.Cells(row1, 1), .Cells(rEnd, 1)).Value
        End With
            
        row2 = 0
        For i = LBound(mk) To UBound(mk)
            If mk(i, 1) > 0 Then
                row2 = i + row1 - 2
            GoTo 99
            End If
        Next
        
        row2 = UBound(mk) + row1 - 2
99
End Sub



Private Sub find_rEnd()
        On Error Resume Next

        If shNm = "Отложено_расход" Then cm = zkNm
        If shNm = "Отложено_приход" Then cm = pzkNm

        With ThisWorkbook.Sheets(shNm)
            r24 = .UsedRange.Rows.Count + .UsedRange.Row
            nm = Range(.Cells(4, cm), .Cells(r24, cm)).Value
        End With
            
        rEnd = 0
        For i = UBound(nm) To LBound(nm) Step -1
            If nm(i, 1) > 0 Then
                rEnd = i + 5
            GoTo 99
            End If
        Next
99
End Sub
