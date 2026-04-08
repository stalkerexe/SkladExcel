Attribute VB_Name = "ф_корзина"
Option Explicit

Dim r As Long: Dim rr As Long
Dim cell As Range


Public iCellsOne As String
Public iCellsTwo As String
Public iRowBox As Long
Public sFormula_1 As String
Public sFormula_2 As String




Public Sub add_box()
On Error Resume Next
iRow = ActiveCell.Row
Call find_gr
Call dann
Call copy_
If frm_Show.Visible = True Then добавить_контролы
With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r + 3, zvNm)), "<>")
End With
With ThisWorkbook.Sheets("Склад")
.Cells(3, iBox1) = iCol
.Cells(3, iBox2) = ThisWorkbook.Sheets("корзина").Cells(rwzvSm, zvSm)
End With
End Sub


Private Sub arr_sk_sk()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
    gr_sk = Range(.Cells(5, skGr), .Cells(iRow, skGr)).Value
    nm_sk = Range(.Cells(5, skNm), .Cells(iRow, skNm)).Value
End With
End Sub

Private Sub find_gr()
On Error Resume Next

Call arr_sk_sk

sGr = ""

For i = UBound(gr_sk) To LBound(gr_sk) Step -1
    If gr_sk(i, 1) <> "" Then
    sGr = nm_sk(i, 1)
    GoTo 99
    End If
Next
99
End Sub




Private Sub dann()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
sID = iRow
sNm = .Cells(iRow, skNm).Value
sCod = .Cells(iRow, skCod).Value
sEd = .Cells(iRow, skEd).Value
sCol = .Cells(iRow, skOst).Value
sBr = .Cells(iRow, skBr).Value
sCnZ = .Cells(iRow, skCnZ).Value
sCnR = .Cells(iRow, skCnR).Value
sSk = .Cells(iRow, skSk).Value
End With
End Sub

Private Sub copy_()
On Error Resume Next

With ThisWorkbook.Sheets("корзина")
r = .Cells(Rows.Count, zvNm).End(xlUp).Row + 1
If r < rwZv Then r = rwZv: GoTo 22
For Each cell In Range(.Cells(rwZv, 1), .Cells(r, 1))
If .Cells(cell.Row, zvSk) = sSk Then
If CStr(.Cells(cell.Row, zvNm)) = sNm And CStr(.Cells(cell.Row, zvCod)) = sCod Then
rr = cell.Row
.Cells(rr, zvCol) = .Cells(rr, zvCol) + 1
Exit Sub
End If
End If
Next
22
.Cells(r, zvCol).NumberFormat = "#,##0.00"
.Cells(r, zvCnZ).NumberFormat = "#,##0.00"
.Cells(r, zvCnR).NumberFormat = "#,##0.00"
.Cells(r, zvSm).NumberFormat = "#,##0.00"
.Cells(r, 1) = sID
.Cells(r, zvSk) = sSk
.Cells(r, zvNm) = sNm
.Cells(r, zvCod) = sCod
.Cells(r, zvEd) = sEd
.Cells(r, zvCnR) = sCnR
.Cells(r, zvCnZ) = sCnZ
.Cells(r, zvOst) = sCol
.Cells(r, zvBr) = sBr
.Cells(r, zvCol) = 1
.Cells(r, zvGr) = sGr
End With

With ThisWorkbook.Sheets("корзина")
    .Cells(r, zvNN) = r - rwZv + 1
End With

iRowBox = r
Call formula_in_box

End Sub

Public Sub del_poz_box()
On Error Resume Next
With ThisWorkbook.Sheets("корзина")
.Cells(iRow, 2).EntireRow.Delete
r = .Cells(Rows.Count, zvNm).End(xlUp).Row
j = 1
For i = rwZv To r
.Cells(i, zvNN) = j
j = j + 1
Next
End With
End Sub


Public Sub formula_in_box()
        On Error Resume Next
        
        If iRowBox = 0 Then Exit Sub
        
        With ThisWorkbook.Sheets("корзина")
        
            iCellsOne = .Cells(iRowBox, zvCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            iCellsTwo = .Cells(iRowBox, zvCnR).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            sFormula_1 = "=" & iCellsOne & "*" & iCellsTwo
            .Cells(iRowBox, zvSm).Formula = sFormula_1
            
            r24 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 9
            iCellsOne = .Cells(rwZv, zvSm).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            iCellsTwo = .Cells(r24, zvSm).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            sFormula_2 = "=SUM(" & iCellsOne & ":" & iCellsTwo & ")"
            .Cells(rwzvSm, zvSm).Formula = sFormula_2
            
        End With

End Sub
