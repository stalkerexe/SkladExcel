Attribute VB_Name = "ф_данные"
Option Explicit

Public Sub dann_zv()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            nomer = .Range("d2")
            sZkz = .Cells(rwZv_zkz, 4).Value
            sAdr = .Cells(rwZv_adr, 4).Value
            sTlf = .Cells(rwZv_tlf, 4).Value
            sMj = .Cells(rwZv_mj, 4).Value
iOpl = .Cells(rwZv_mj, zvSm).Value
iSkid = .Cells(rwZv_mj, zvOst).Value
            sDt = VBA.CDate(.Cells(rwZv_dt, 4).Value)
            sDt2 = VBA.CDate(.Cells(rwZv_dt2, 4).Value)
            sComm = .Cells(1, zvComm)
            summ = .Cells(rwzvSm, zvSm).Value
        End With
End Sub

Public Sub dann_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Приход")
            nomer = .Range("d2")
            sZkz = .Cells(rwPr_zkz, 4).Value
            sMj = .Cells(rwPr_mj, 4).Value
            sDt = VBA.CDate(.Cells(rwPr_dt, 4).Value)
            
            summ = .Cells(rwzvSm, prSm).Value
            sComm = .Cells(1, prComm)
            
            sDoc = Cells(1, prDoc).Value
            sDocN = "'" & Cells(1, prDocN).Value
            sDocDt = Cells(1, prDocDt).Value
            
sOsn = .Cells(rwPr_doc, 4).Value
        End With
End Sub



Public Sub dann_zk_rs()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_расход")
            nomer = .Cells(iRow, zkNom)
            sZkz = .Cells(iRow, zkZkz)
            sTlf = .Cells(iRow, zkTlf)
            sAdr = .Cells(iRow, zkAdr)
            sMj = .Cells(iRow, zkMj)
iOpl = .Cells(iRow, zkOpl)
iSkid = .Cells(iRow, zkSkid).Value
            summ = .Cells(iRow, zkSm)
            sComm = .Cells(iRow + 1, zkComm).Value
            sDt = VBA.CDate(.Cells(iRow, zkDt1))
            sDt2 = VBA.CDate(.Cells(iRow, zkDt2))
        End With
End Sub

Public Sub dann_zk_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_приход")
            nomer = .Cells(iRow, pzkNom)
            sZkz = .Cells(iRow, pzkPsv)
            sMj = .Cells(iRow, pzkMj)
            sDt = VBA.CDate(.Cells(iRow, pzkDt))
            summ = .Cells(iRow, pzkSm)
            
            sComm = .Cells(iRow + 1, pzkComm).Value
            
            sDoc = .Cells(iRow, pzkDoc).Value
            sDocN = .Cells(iRow, pzkDocN).Value
            sDocDt = .Cells(iRow, pzkDocDt).Value
            sOsn = .Cells(iRow + 1, pzkOsn).Value
        End With
End Sub




Public Sub dann_arh_rs()
    On Error Resume Next
    With ThisWorkbook.Sheets(shNmArh)
        nomer = .Cells(iRow, arhNom)
        sZkz = .Cells(iRow, arhZkz)
        sTlf = .Cells(iRow, arhTlf)
        sAdr = .Cells(iRow, arhAdr)
        sMj = .Cells(iRow, arhMj)
        summ = .Cells(iRow, arhSmA)
        sDt = VBA.CDate(.Cells(iRow, arhDt))
        sDt2 = VBA.CDate(.Cells(iRow, arhDt2))
    End With
End Sub

Public Sub dann_arh_pr()
    On Error Resume Next
    With ThisWorkbook.Sheets(shNmArh)
        nomer = .Cells(iRow, arhNom)
        sZkz = .Cells(iRow, arhZkz)
        sOsn = .Cells(iRow, arhDoc)
        sComm = .Cells(iRow, arhComm)
        sMj = .Cells(iRow, arhMj)
        summ = .Cells(iRow, arhSmA)
        sDt = VBA.CDate(.Cells(iRow, arhDt))
    End With
End Sub

Public Sub dann_arh_vz()
    On Error Resume Next
    With ThisWorkbook.Sheets(shNmArh)
        nomer = .Cells(iRow, arhNom)
        sOsn = .Cells(iRow, avzNk)
        summ = .Cells(iRow, arhSmA)
        sDt = VBA.CDate(.Cells(iRow, arhDt))
        sZkz = .Cells(iRow, arhZkz)
        sMj = .Cells(iRow, arhMj)
    End With
End Sub

