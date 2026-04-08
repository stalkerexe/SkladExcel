Attribute VB_Name = "о_оформить_заказ_pr"
Option Explicit

Public Sub оформить_заказ_pr()
        Call copy_
        Call format_pr
End Sub

Private Sub copy_()
        On Error Resume Next
        With ThisWorkbook.Sheets("Приход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range("a" & rwZv & ":a" & r7 + 44).EntireRow.Delete
            .Cells(rwzvSm, zvSm) = ""
            .Range("a1") = ""
            .Range("d2") = ""
        End With
        Sheets("Приход").Activate
        With ThisWorkbook.Sheets("корзина")
        r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row
        .Range(.Cells(rwZv, 1), .Cells(r7, zvCol)).Copy
        Cells(rwZv, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range(.Cells(rwZv, zvCnZ), .Cells(r7, zvCnZ)).Copy
        Cells(rwZv, prCnZ).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range(.Cells(rwZv, zvCnR), .Cells(r7, zvCnR)).Copy
        Cells(rwZv, prCnR).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range(.Cells(rwZv, zvSm), .Cells(r7, zvSm)).Copy
        Cells(rwZv, prSm).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range(.Cells(rwZv, zvSk), .Cells(r7, zvSk)).Copy
        Cells(rwZv, prSk).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
.Range(.Cells(rwZv, zvGr), .Cells(r7, zvGr)).Copy
Cells(rwZv, prGr).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Application.CutCopyMode = False
        End With
End Sub

Private Sub format_pr()
        On Error Resume Next
        row1 = rwZv
        row2 = ThisWorkbook.Sheets("Приход").Cells(Rows.Count, prNm).End(xlUp).Row
        Call format_pr_
        Cells(rwZv, prCol) = Cells(rwZv, prCol)
End Sub





