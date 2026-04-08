Attribute VB_Name = "о_ост_обновить_zk"
Option Explicit
Dim cc()

Public Sub ost_sk_zk()
        On Error Resume Next
        Call arr_this
        Call перебрать_sk
        Call resize_this_ost
End Sub

Private Sub arr_this()
        On Error Resume Next
        
        row1 = 5
        
        If shNm = "Отложено_расход" Then
            row2 = ThisWorkbook.Sheets(shNm).Cells(Rows.Count, zkNm).End(xlUp).Row + 1
            Call arr_zk_this
        End If
        
        If shNm = "Отложено_приход" Then
            row2 = ThisWorkbook.Sheets(shNm).Cells(Rows.Count, pzkNm).End(xlUp).Row + 1
            Call arr_zk_this_pr
        End If
        
        ReDim cc(LBound(nm) To UBound(nm), 1 To 1)
        
End Sub

Private Sub resize_this_ost()
        On Error Resume Next
        
        If shNm = "Отложено_расход" Then
            ThisWorkbook.Sheets(shNm).Cells(5, zkOst).Resize(UBound(cc), 1) = cc
        End If
        
        If shNm = "Отложено_приход" Then
            ThisWorkbook.Sheets(shNm).Cells(5, pzkOst).Resize(UBound(cc), 1) = cc
        End If

End Sub

Private Sub перебрать_sk()
        On Error Resume Next
        Call load_sk
        For n = 0 To dic_sk.Count - 1
        If dic_sk.Item(n) <> "" Then
        sSk = dic_sk.Item(n): Waite.Label2.Caption = sSk: DoEvents
        flag = 0
        Call arr_sk_
        End If
33
        Next
End Sub

Private Sub arr_sk_()
        On Error Resume Next
        
        Call find_cm_sk
        
        With ThisWorkbook.Sheets(shNm)
            r7 = .Cells(Rows.Count, skNm).End(xlUp).Row + 2
            iCol = Application.CountIf(.Range(.Cells(5, skNm), .Cells(r7, skNm)), "<>")
            Application.CutCopyMode = False
        End With
        
        If iCol = 0 Then Exit Sub
        
Call arr_select_sk

        Call parse

End Sub


Private Sub parse()
        On Error Resume Next
        For i = LBound(nm) To UBound(nm)
            If sk(i, 1) = sSk Then
            If nm(i, 1) <> "" Then
            
            For ii = LBound(c) To UBound(c)
                If c(ii, 4) = nm(i, 1) And CStr(c(ii, 3)) = CStr(cod(i, 1)) Then
                cc(i, 1) = c(ii, 6)
                GoTo 33
                End If
            Next ii
33
            End If
            End If
        Next i
End Sub

