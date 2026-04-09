Attribute VB_Name = "я_____________________________"
Option Explicit
Dim rEnd As Long


Public Sub find_row2_this()
        On Error Resume Next

        Call find_rEnd

        With ThisWorkbook.Sheets(shNm)
            mk = Range(.Cells(row1, 1), .Cells(rEnd, 1)).value
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


' ИСПРАВЛЕНИЕ #8: в оригинале использовалась глобальная переменная nm для
' временного считывания данных из листа внутри find_rEnd.
' Это затирало глобальный nm, который мог содержать данные текущей накладной.
' Теперь используется локальный массив tmpNm — глобальный nm не затрагивается.
Private Sub find_rEnd()
        On Error Resume Next

        Dim tmpNm() As Variant
        Dim cmLocal As Integer

        If shNm = "Отложено_расход" Then cmLocal = zkNm
        If shNm = "Отложено_приход" Then cmLocal = pzkNm

        With ThisWorkbook.Sheets(shNm)
            r24 = .UsedRange.Rows.Count + .UsedRange.Row
            tmpNm = Range(.Cells(4, cmLocal), .Cells(r24, cmLocal)).value
        End With

        rEnd = 0
        For i = UBound(tmpNm) To LBound(tmpNm) Step -1
            If tmpNm(i, 1) > 0 Then
                rEnd = i + 5
            GoTo 99
            End If
        Next
99
End Sub

