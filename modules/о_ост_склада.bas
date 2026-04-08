Attribute VB_Name = "о_ост_склада"


Public Sub ost_skds()
        On Error Resume Next
        
        For i = LBound(nm) To UBound(nm)
            If nm(i, 1) <> "" Then
            
                    sSk = sk(i, 1)
                    rw = id(i, 1) + 1
                    
                    Call find_cm_sk
                    Call sk_ost_do
            
            End If
        Next i

End Sub

Private Sub sk_ost_do()
        On Error Resume Next
        With ThisWorkbook.Sheets("sk_123")
            If iOperation = "zv" Then .Cells(rw, cm) = .Cells(rw, cm) - CDbl(col(i, 1))
            If iOperation = "pr" Then .Cells(rw, cm) = .Cells(rw, cm) + CDbl(col(i, 1))
            If iOperation = "vz" Then .Cells(rw, cm) = .Cells(rw, cm) + CDbl(col(i, 1))
        End With
End Sub


Public Sub find_cm_sk()
        If sSk = "Материалы" Then cm = 2
        If sSk = "Металлопрокат" Then cm = 4
        If sSk = "Спецодежда" Then cm = 6
End Sub


