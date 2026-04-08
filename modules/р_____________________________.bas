Attribute VB_Name = "р_____________________________"
Option Explicit

Public Sub clear_zv_all()
        On Error Resume Next
        Call unload_mn_vid:  DoEvents
        If MsgBox("Удалить все позиции из накладной?       ", vbOKCancel + vbQuestion, "Очистить") = vbCancel Then Exit Sub
        Call doScreenOff
        Call do_clear_zv_all
        Call doScreenOn
End Sub

Private Sub do_clear_zv_all()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range(.Cells(rwZv, 2), .Cells(r7 + 44, 2)).EntireRow.Delete
            .Cells(rwzvSm, zvSm) = ""
            .Range("a1") = ""
            .Cells(1, zvComm) = ""
.Cells(rwZv_mj, zvOst) = ""
        End With
        
        Call режим_редактирования_off_pr("Расход")
        
        Call clear_box
End Sub


Public Sub RemoveDuplicates_sk()
On Error Resume Next
Call clearBf
With ThisWorkbook.Sheets("буфер")
.Cells(1, "a").Resize(UBound(sk), 1) = sk
End With
With ThisWorkbook.Sheets("буфер")
r24 = .Cells(Rows.Count, 1).End(xlUp).Row + 1
.Range("a1:a" & r24).RemoveDuplicates Columns:=1, Header:=xlNo
r24 = .Cells(Rows.Count, 1).End(xlUp).Row + 1
sk_2 = .Range("a1:a" & r24).Value
End With
Application.CutCopyMode = False
End Sub


Public Sub clear_sk()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
    Call AutoFilter_delete
    r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
    .Range("a" & 5 & ":a" & r7 + 44).EntireRow.Delete
End With
End Sub




Public Sub diap_zk_this()
        Call arr_mk_zk
        Call find_mk_zk_this
        Call find_row2_zk_this
        row1 = row1 + 4
        row2 = row2 + 4
End Sub

Private Sub arr_mk_zk()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNm)
            r24 = .UsedRange.Rows.Count + .UsedRange.Row
            mk = Range(.Cells(5, 1), .Cells(r24, 1)).Value
            nm = Range(.Cells(5, zkNm), .Cells(r24, zkNm)).Value
        End With
End Sub

Private Sub find_mk_zk_this()
        On Error Resume Next
        row1 = 0
        For i = LBound(mk) To UBound(mk)
            If mk(i, 1) = marker Then
                row1 = i
                GoTo 99
            End If
        Next
99
End Sub

Public Sub find_row2_zk_this()
        On Error Resume Next

        row2 = 0
        For ii = row1 + 1 To UBound(nm)
            If mk(ii, 1) <> "" Then
                row2 = ii - 1
                GoTo 99
            End If
        Next
        
        row2 = UBound(nm)
99
End Sub





Public Sub find_mk_zk_file()
        On Error Resume Next
        
        With ThisWorkbook.Sheets(shNm)
            r24 = .UsedRange.Rows.Count + .UsedRange.Row
            mk = .Range(.Cells(3, 1), .Cells(r24, 1)).Value
        End With
        
        row1 = 0
        
        For i = LBound(mk) To UBound(mk)
            If mk(i, 1) = marker Then
                row1 = i + 2
                GoTo 99
            End If
        Next
        
        Exit Sub
99
        Call find_row2
End Sub

Private Sub find_row2()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNm)
            r9 = .Cells(Rows.Count, zkNm).End(xlUp).Row + 2
            row2 = .Cells(row1, 1).End(xlDown).Row - 1
        End With
        If row2 > r9 Then row2 = r9
End Sub




Public Sub find_mk_arh()
        On Error Resume Next
        
        With ThisWorkbook.Sheets(shNmArh)
            r24 = .UsedRange.Rows.Count + .UsedRange.Row
            mk = .Range(.Cells(3, 1), .Cells(r24, 1)).Value
        End With
        
        row1 = 0
        
        For i = LBound(mk) To UBound(mk)
            If mk(i, 1) = marker Then
                row1 = i + 2
                GoTo 99
            End If
        Next
        
        Exit Sub
99
        Call find_row2_arh
End Sub

Private Sub find_row2_arh()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r9 = .UsedRange.Rows.Count + .UsedRange.Row
            row2 = .Cells(row1, 1).End(xlDown).Row - 1
        End With
        If row2 > r9 Then row2 = r9
End Sub






Public Sub delete_zk_in_file()
        On Error Resume Next

        Call find_mk_zk_file
        If row1 = 0 Then GoTo 99
        
        Call delete_nk_zk
99
End Sub

Private Sub delete_nk_zk()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNm)
            .Range(.Cells(row1, 2), .Cells(row2, 2)).EntireRow.Delete
        End With
End Sub



Public Sub delete_pr_in_file()
        On Error Resume Next

        Call find_mk_zk_file
        If row1 = 0 Then GoTo 99
        
        Call delete_nk_zk
99
End Sub


