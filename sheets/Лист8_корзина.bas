' Component: Лист8  [корзина]
' Type: Document
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
If Not Intersect(Target, Cells(rwzvSm, zvSm)) Is Nothing Then
Call sum_box
End If
End Sub


