Attribute VB_Name = "склад_обновить"
Option Explicit

Public Sub sklad_obnovitt()
        Call doScreenOff
        Call do_sklad_obnovitt
        Call doScreenOn
End Sub

Public Sub do_sklad_obnovitt()
        On Error Resume Next
        
        ThisWorkbook.Activate
        
        Call arr_sk_this_sheet
        If iCol = 0 Then Exit Sub
        
        Call parse_sk
        Call do_sklad_show
        
End Sub

Private Sub arr_sk_this_sheet()
        On Error Resume Next
        With ThisWorkbook.Sheets("Склад")
            r7 = .Cells(Rows.Count, skNm).End(xlUp).Row + 4
            nn = Range(.Cells(5, 1), .Cells(r7, 1)).Value
            nm = Range(.Cells(5, skNm), .Cells(r7, skNm)).Value
            iCol = Application.CountIf(Range(.Cells(5, skNm), .Cells(r7, skNm)), "<>")
        End With
End Sub

Private Sub parse_sk()
        On Error Resume Next
        
        ReDim c(1 To iCol, 1 To 1)

        j = 1
        For i = LBound(nm) To UBound(nm)
            If nn(i, 1) <> "" Then
                c(j, 1) = nm(i, 1)
                j = j + 1
            End If
        Next

End Sub

