Attribute VB_Name = "о_оформить_заказ"
Option Explicit


Public Sub оформить_заказ()
        Call copy_
        Call format_zv
End Sub

Private Sub copy_()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r7 + 44).EntireRow.Delete
            .Cells(rwzvSm, zvSm) = ""
            .Range("a1") = ""
            .Range("d2") = ""
        End With
        Sheets("Расход").Activate
        With ThisWorkbook.Sheets("корзина")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row
            .Range(.Cells(rwZv, 1), .Cells(r7, 100)).Copy
            Cells(rwZv, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        End With
End Sub

Private Sub format_zv()
        On Error Resume Next
        row1 = rwZv
        row2 = ThisWorkbook.Sheets("Расход").Cells(Rows.Count, zvNm).End(xlUp).Row
        Call format_zv_
        Cells(rwZv, zvCol) = Cells(rwZv, zvCol)
End Sub


