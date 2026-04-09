Attribute VB_Name = "архив_сохранить"
Public shNmArh As String

Public Sub save_nk()

        Call proverka_arh
        If iCol >= 30 Then Exit Sub
        
        Call sv_nk_to_arh
End Sub

Private Sub proverka_arh()
        If Trim$(CStr(shNmArh)) = "" Then Call find_vid_arhh
        Call arr_arh_proverka
End Sub



Private Sub sv_nk_to_arh()
        If Trim$(CStr(shNmArh)) = "" Then Call find_vid_arhh
        Call copy_to_arh
End Sub

Public Sub find_vid_arhh()
        Select Case iVid
            Case "pr": shNmArh = "arh_prr"
            Case "ot": shNmArh = "arh_zkk"
            Case "vz": shNmArh = "arh_vzz"
            Case Else: shNmArh = ""
        End Select
End Sub




Private Sub copy_to_arh()

        Call copy_to_arh_nk
        
        Call copy_to_arh_dann

End Sub

Private Sub copy_to_arh_nk()
        On Error Resume Next

        With ThisWorkbook.Sheets(shNmArh)
            n7 = .Cells(Rows.Count, arhNm).End(xlUp).Row + 2: If n7 < 3 Then n7 = 3
            .Cells(n7, arhNN).Resize(UBound(nm), 1) = nn
            .Cells(n7, arhNm).Resize(UBound(nm), 1) = nm
            
            n2 = .Cells(Rows.Count, arhNm).End(xlUp).Row
            Range(.Cells(n7, arhCod), .Cells(n2, arhCod)).NumberFormat = "@"
            .Cells(n7, arhCod).Resize(UBound(nm), 1) = cod
            
            .Cells(n7, arhEd).Resize(UBound(nm), 1) = ed
            .Cells(n7, arhCol).Resize(UBound(nm), 1) = col
            .Cells(n7, arhCnR).Resize(UBound(nm), 1) = cnR
            .Cells(n7, arhCnZ).Resize(UBound(nm), 1) = cnZ
            .Cells(n7, arhSm).Resize(UBound(nm), 1) = sm
            .Cells(n7, arhSk).Resize(UBound(nm), 1) = sk
        End With

End Sub

Private Sub copy_to_arh_dann()
        On Error Resume Next

        With ThisWorkbook.Sheets(shNmArh)
            .Cells(n7, 1) = marker
            .Cells(n7, arhSmA) = summ
            .Cells(n7, arhComm) = sComm
        End With

        With ThisWorkbook.Sheets(shNmArh)
            n2 = .Cells(Rows.Count, arhNm).End(xlUp).Row
            Range(.Cells(n7, arhNom), .Cells(n2, arhNom)) = nomer
            Range(.Cells(n7, arhDt), .Cells(n2, arhDt)) = VBA.CDate(sDt)
            Range(.Cells(n7, arhZkz), .Cells(n2, arhZkz)) = sZkz
            Range(.Cells(n7, arhMj), .Cells(n2, arhMj)) = sMj
        End With
        
        If iVid = "ot" Then
            With ThisWorkbook.Sheets(shNmArh)
                Range(.Cells(n7, arhTlf), .Cells(n2, arhTlf)).NumberFormat = "@"
                Range(.Cells(n7, arhDt2), .Cells(n2, arhDt2)) = VBA.CDate(sDt2)
                Range(.Cells(n7, arhAdr), .Cells(n2, arhAdr)) = sAdr
                Range(.Cells(n7, arhTlf), .Cells(n2, arhTlf)) = sTlf
Range(.Cells(n7, arhOpl), .Cells(n2, arhOpl)) = iOpl
Range(.Cells(n7, arhSkid), .Cells(n2, arhSkid)) = iSkid
            End With
        End If

        If iVid = "pr" Then
            With ThisWorkbook.Sheets(shNmArh)
                Range(.Cells(n7, arhDoc), .Cells(n2, arhDoc)) = sOsn
            End With
        End If

        If iVid = "vz" Then
            With ThisWorkbook.Sheets(shNmArh)
Range(.Cells(n7, avzNk), .Cells(n2, avzNk)) = sDoc
Range(.Cells(n7, avzMk), .Cells(n2, avzMk)) = iMk
                
                 iSm = Application.Sum(.Range(.Cells(n7, arhSm), .Cells(n2, arhSm)))
                .Cells(n7, arhSmA) = iSm
            End With
        End If

        Call format_arh

End Sub





Private Sub format_arh()
        On Error Resume Next

        With ThisWorkbook.Sheets(shNmArh)
           
            Range(.Cells(n7, arhDt), .Cells(n2, arhDt2)).NumberFormat = "dd.mm.yyyy"
            
            
            Range(.Cells(n7, arhEd), .Cells(n2, arhSmA)).HorizontalAlignment = xlCenter
            Range(.Cells(n7, arhNN), .Cells(n2, arhNN)).HorizontalAlignment = xlCenter
            Range(.Cells(n7, arhNom), .Cells(n2, arhDt2)).HorizontalAlignment = xlCenter
            
            With Range(.Cells(n7, 1), .Cells(n2, 44))
                .Font.Name = "Times New Roman"
                .Font.Size = 10
            End With
            
            With Range(.Cells(n7, arhComm), .Cells(n2, arhComm))
                .WrapText = False
                .Font.Size = 8
                 If n7 <> n2 Then .Merge
                .IndentLevel = 1
                .VerticalAlignment = xlTop
            End With
            
                        
        End With

End Sub


