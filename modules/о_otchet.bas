Attribute VB_Name = "о_otchet"
Public iCmOt As Integer
Public wbOt As Workbook


Public Sub otchet_do()
        Call doScreenOff
        Call do_otchet_do
        Call doScreenOn
End Sub

Private Sub do_otchet_do()

        Call find_vid_arhh
        Call do_otchet
        
        Set wbOt = Nothing
        Call erase_arr_arh_this
        Erase c
End Sub

Private Sub do_otchet()
        On Error Resume Next

        Call arr_arh_for_otchet
        If iCol = 0 Then Exit Sub
        
        Call parse_arh
        
        Call resize_arh
        
        Call set_otchet
        
End Sub


Private Sub resize_arh()
        On Error Resume Next

        Application.Workbooks.Add (1)
        Set wbOt = ActiveWorkbook
        
        With wbOt.ActiveSheet
            .Cells(2, 2).Resize(UBound(c), iCmOt) = c
        End With
        
        Call format_otchet
        
End Sub


Private Sub parse_arh()
        If iVid = "pr" Then iCmOt = 10: Call parse_arh_pr
        If iVid = "ot" Then iCmOt = 15: Call parse_arh_ot
End Sub

Private Sub parse_arh_pr()
        On Error Resume Next

        ReDim c(LBound(nm) To UBound(nm), 1 To iCmOt)
        
        j = 1
        
        For i = LBound(nm) To UBound(nm): Waite.Label2.Caption = nm(i, 1): DoEvents
            
            If nm(i, 1) <> "" Then
            
                    c(j, 1) = nom(i, 1)
                    c(j, 2) = dt(i, 1)
                    c(j, 3) = nm(i, 1)
                    c(j, 4) = cod(i, 1)
                    c(j, 5) = col(i, 1)
                    c(j, 6) = cnZ(i, 1)
                    c(j, 7) = sm(i, 1)
                    c(j, 8) = mj(i, 1)
                    c(j, 9) = zkz(i, 1)
                    c(j, 10) = doc(i, 1)
                    
                    j = j + 1
                
            End If
                
        Next i

End Sub

Private Sub parse_arh_ot()
        On Error Resume Next

        ReDim c(LBound(nm) To UBound(nm), 1 To iCmOt)
        
        j = 1
        
        For i = LBound(nm) To UBound(nm): Waite.Label2.Caption = nm(i, 1): DoEvents
            
            If nm(i, 1) <> "" Then
            
                c(j, 1) = nom(i, 1)
                c(j, 2) = dt(i, 1)
                c(j, 3) = nm(i, 1)
                c(j, 4) = cod(i, 1)
                c(j, 5) = col(i, 1)
                c(j, 6) = cnR(i, 1)
                c(j, 7) = sm(i, 1)
                c(j, 8) = cnZ(i, 1)
                
                сумма_закуп = col(i, 1) * cnZ(i, 1)
                c(j, 9) = сумма_закуп
                
c(j, 10) = sm(i, 1) - сумма_закуп
                c(j, 11) = mj(i, 1)
                c(j, 12) = zkz(i, 1)
                
                c(j, 13) = sk(i, 1)
                c(j, 14) = opl(i, 1)
                c(j, 15) = skid(i, 1)

                j = j + 1
                
            End If
                
        Next i

End Sub


