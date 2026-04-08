Attribute VB_Name = "p_popup"
Option Explicit


Public Sub del_poz_zv()
On Error Resume Next
iRow = ActiveCell.Row
Range(Cells(iRow, 3), Cells(iRow, 7)).Select
Call del_poz_zv_
Cells(iRow, 1).Select
End Sub

Private Sub del_poz_zv_()
On Error Resume Next
ActiveCell.EntireRow.Delete
r7 = Cells(Rows.Count, zvNm).End(xlUp).Row
j = 1
For i = rwZv To r7
Cells(i, zvNN) = j
j = j + 1
Next
Cells(rwzvSm, zvSm) = Application.Sum(Range(Cells(rwZv, zvSm), Cells(r7 + 4, zvSm)))
iCol = Application.CountIf(Range(Cells(rwZv, zvNm), Cells(r7 + 3, zvNm)), "<>")
If iCol = 0 Then Cells(rwZv_mj, zvOst) = ""
End Sub

Public Sub del_poz_pr()
On Error Resume Next
iRow = ActiveCell.Row
Range(Cells(iRow, 3), Cells(iRow, 7)).Select
Call del_poz_pr_
Cells(iRow, 1).Select
End Sub

Private Sub del_poz_pr_()
On Error Resume Next
ActiveCell.EntireRow.Delete
r7 = Cells(Rows.Count, prNN).End(xlUp).Row
j = 1
For i = rwZv To r7
Cells(i, prNN) = j
j = j + 1
Next
Cells(rwzvSm, prSm) = Application.Sum(Range(Cells(rwZv, prSm), Cells(r7 + 4, prSm)))
End Sub


