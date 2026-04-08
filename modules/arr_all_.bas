Attribute VB_Name = "arr_all_"
Option Explicit

Public Sub arr_zv()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 2
            nn = Range(.Cells(rwZv, zvNN), .Cells(r7, zvNN)).Value
            nm = Range(.Cells(rwZv, zvNm), .Cells(r7, zvNm)).Value
            cod = Range(.Cells(rwZv, zvCod), .Cells(r7, zvCod)).Value
            ed = Range(.Cells(rwZv, zvEd), .Cells(r7, zvEd)).Value
            cnR = Range(.Cells(rwZv, zvCnR), .Cells(r7, zvCnR)).Value
            cnZ = Range(.Cells(rwZv, zvCnZ), .Cells(r7, zvCnZ)).Value
cn = Range(.Cells(rwZv, zvCn), .Cells(r7, zvCn)).Value
            col = Range(.Cells(rwZv, zvCol), .Cells(r7, zvCol)).Value
            sm = Range(.Cells(rwZv, zvSm), .Cells(r7, zvSm)).Value
            ost = Range(.Cells(rwZv, zvOst), .Cells(r7, zvOst)).Value
            sk = Range(.Cells(rwZv, zvSk), .Cells(r7, zvSk)).Value
            id = Range(.Cells(rwZv, 1), .Cells(r7, 1)).Value
            iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r7, zvNm)), "<>")
        End With
End Sub

Public Sub erase_arr_zv()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed: Erase cnR: Erase cnZ: Erase col: Erase sm: Erase ost: Erase sk
End Sub


Public Sub arr_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Приход")
            r7 = .Cells(Rows.Count, prNm).End(xlUp).Row + 2
            nn = Range(.Cells(rwZv, prNN), .Cells(r7, prNN)).Value
            nm = Range(.Cells(rwZv, prNm), .Cells(r7, prNm)).Value
            ed = Range(.Cells(rwZv, prEd), .Cells(r7, prEd)).Value
            cod = Range(.Cells(rwZv, prCod), .Cells(r7, prCod)).Value
            cnR = Range(.Cells(rwZv, prCnR), .Cells(r7, prCnR)).Value
            cnZ = Range(.Cells(rwZv, prCnZ), .Cells(r7, prCnZ)).Value
            col = Range(.Cells(rwZv, prCol), .Cells(r7, prCol)).Value
            sm = Range(.Cells(rwZv, prSm), .Cells(r7, prSm)).Value
            sk = Range(.Cells(rwZv, prSk), .Cells(r7, prSk)).Value
            gr = Range(.Cells(rwZv, prGr), .Cells(r7, prGr)).Value
            id = Range(.Cells(rwZv, 1), .Cells(r7, 1)).Value
            iCol = Application.CountIf(Range(.Cells(rwZv, prNm), .Cells(r7, prNm)), "<>")
        End With
End Sub

Public Sub erase_arr_pr()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed: Erase cnR: Erase cnZ: Erase col: Erase sm: Erase sk
End Sub

Public Sub erase_arr_sk()
        On Error Resume Next
        Erase gr_sk: Erase nm_sk: Erase ed_sk: Erase cod_sk: Erase cnR_sk: Erase cnZ_sk: Erase col_sk: Erase cr_sk: Erase br_sk: Erase comm_sk
End Sub




Public Sub arr_zk_this()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_расход")
            nn = Range(.Cells(row1, zkNN), .Cells(row2, zkNN)).Value
            nm = Range(.Cells(row1, zkNm), .Cells(row2, zkNm)).Value
            cod = Range(.Cells(row1, zkCod), .Cells(row2, zkCod)).Value
            ed = Range(.Cells(row1, zkEd), .Cells(row2, zkEd)).Value
            cnR = Range(.Cells(row1, zkCnR), .Cells(row2, zkCnR)).Value
            cnZ = Range(.Cells(row1, zkCnZ), .Cells(row2, zkCnZ)).Value
cn = Range(.Cells(row1, zkCn), .Cells(row2, zkCn)).Value
            col = Range(.Cells(row1, zkCol), .Cells(row2, zkCol)).Value
            sm = Range(.Cells(row1, zkSm), .Cells(row2, zkSm)).Value
            ost = Range(.Cells(row1, zkOst), .Cells(row2, zkOst)).Value
            sk = Range(.Cells(row1, zkSk), .Cells(row2, zkSk)).Value
            id = Range(.Cells(row1, zkID), .Cells(row2, zkID)).Value
            iCol = Application.CountIf(Range(.Cells(row1, zkNm), .Cells(row2, zkNm)), "<>")
        End With
End Sub

Public Sub arr_zk_this_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_приход")
            nn = Range(.Cells(row1, pzkNN), .Cells(row2, pzkNN)).Value
            nm = Range(.Cells(row1, pzkNm), .Cells(row2, pzkNm)).Value
            cod = Range(.Cells(row1, pzkCod), .Cells(row2, pzkCod)).Value
            ed = Range(.Cells(row1, pzkEd), .Cells(row2, pzkEd)).Value
            cnR = Range(.Cells(row1, pzkCnR), .Cells(row2, pzkCnR)).Value
            cnZ = Range(.Cells(row1, pzkCnZ), .Cells(row2, pzkCnZ)).Value
            col = Range(.Cells(row1, pzkCol), .Cells(row2, pzkCol)).Value
            sm = Range(.Cells(row1, pzkSm), .Cells(row2, pzkSm)).Value
            ost = Range(.Cells(row1, pzkOst), .Cells(row2, pzkOst)).Value
            sk = Range(.Cells(row1, pzkSk), .Cells(row2, pzkSk)).Value
            gr = Range(.Cells(row1, pzkGr), .Cells(row2, pzkGr)).Value
            id = Range(.Cells(row1, pzkID), .Cells(row2, pzkID)).Value
            iCol = Application.CountIf(Range(.Cells(row1, pzkNm), .Cells(row2, pzkNm)), "<>")
        End With
End Sub

Public Sub erase_arr_zk_this()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed: Erase cnR: Erase cnZ: Erase col: Erase sm: Erase ost: Erase sk
End Sub



Public Sub arr_arh_rs()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            nn = Range(.Cells(row1, arhNN), .Cells(row2, arhNN)).Value
            nm = Range(.Cells(row1, arhNm), .Cells(row2, arhNm)).Value
            cod = Range(.Cells(row1, arhCod), .Cells(row2, arhCod)).Value
            ed = Range(.Cells(row1, arhEd), .Cells(row2, arhEd)).Value
            cnR = Range(.Cells(row1, arhCnR), .Cells(row2, arhCnR)).Value
            cnZ = Range(.Cells(row1, arhCnZ), .Cells(row2, arhCnZ)).Value
            col = Range(.Cells(row1, arhCol), .Cells(row2, arhCol)).Value
            sm = Range(.Cells(row1, arhSm), .Cells(row2, arhSm)).Value
            sk = Range(.Cells(row1, arhSk), .Cells(row2, arhSk)).Value
            iCol = Application.CountIf(Range(.Cells(row1, arhNm), .Cells(row2, arhNm)), "<>")
        End With
End Sub

Public Sub arr_arh_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            nn = Range(.Cells(row1, arhNN), .Cells(row2, arhNN)).Value
            nm = Range(.Cells(row1, arhNm), .Cells(row2, arhNm)).Value
            cod = Range(.Cells(row1, arhCod), .Cells(row2, arhCod)).Value
            ed = Range(.Cells(row1, arhEd), .Cells(row2, arhEd)).Value
            cnR = Range(.Cells(row1, arhCnR), .Cells(row2, arhCnR)).Value
            cnZ = Range(.Cells(row1, arhCnZ), .Cells(row2, arhCnZ)).Value
            col = Range(.Cells(row1, arhCol), .Cells(row2, arhCol)).Value
            sm = Range(.Cells(row1, arhSm), .Cells(row2, arhSm)).Value
            sk = Range(.Cells(row1, arhSk), .Cells(row2, arhSk)).Value
            iCol = Application.CountIf(Range(.Cells(row1, arhNm), .Cells(row2, arhNm)), "<>")
        End With
End Sub

Public Sub erase_arr_arh_this()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed: Erase cnR: Erase cnZ: Erase col: Erase sm: Erase ost: Erase sk
End Sub



Public Sub arr_arh_all()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r7 = .Cells(Rows.Count, arhNom).End(xlUp).Row + 2
            mk = Range(.Cells(2, 1), .Cells(r7, 1)).Value
            nom = Range(.Cells(2, arhNom), .Cells(r7, arhNom)).Value
            zkz = Range(.Cells(2, arhZkz), .Cells(r7, arhZkz)).Value
            dt = Range(.Cells(2, arhDt), .Cells(r7, arhDt)).Value
            mj = Range(.Cells(2, arhMj), .Cells(r7, arhMj)).Value
            doc = Range(.Cells(2, arhDoc), .Cells(r7, arhDoc)).Value
            iCol = Application.CountIf(Range(.Cells(2, 1), .Cells(r7, 1)), "<>")
        End With
End Sub



Public Sub arr_arh_for_otchet()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r7 = .Cells(Rows.Count, arhNm).End(xlUp).Row + 2
            
            nm = Range(.Cells(2, arhNm), .Cells(r7, arhNm)).Value
            cod = Range(.Cells(2, arhCod), .Cells(r7, arhCod)).Value
            ed = Range(.Cells(2, arhEd), .Cells(r7, arhEd)).Value
            cnR = Range(.Cells(2, arhCnR), .Cells(r7, arhCnR)).Value
            cnZ = Range(.Cells(2, arhCnZ), .Cells(r7, arhCnZ)).Value
            col = Range(.Cells(2, arhCol), .Cells(r7, arhCol)).Value
            sm = Range(.Cells(2, arhSm), .Cells(r7, arhSm)).Value
            sk = Range(.Cells(2, arhSk), .Cells(r7, arhSk)).Value
            
            mk = Range(.Cells(2, 1), .Cells(r7, 1)).Value
            nom = Range(.Cells(2, arhNom), .Cells(r7, arhNom)).Value
            zkz = Range(.Cells(2, arhZkz), .Cells(r7, arhZkz)).Value
            dt = Range(.Cells(2, arhDt), .Cells(r7, arhDt)).Value
            mj = Range(.Cells(2, arhMj), .Cells(r7, arhMj)).Value
            doc = Range(.Cells(2, arhDoc), .Cells(r7, arhDoc)).Value
            
            opl = Range(.Cells(2, arhOpl), .Cells(r7, arhOpl)).Value
            skid = Range(.Cells(2, arhSkid), .Cells(r7, arhSkid)).Value

            iCol = Application.CountIf(Range(.Cells(2, 1), .Cells(r7, 1)), "<>")
        End With
End Sub

Public Sub arr_arh_proverka()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r7 = .Cells(Rows.Count, arhNm).End(xlUp).Row + 2
            iCol = Application.CountIf(Range(.Cells(2, 1), .Cells(r7, 1)), "<>")
        End With
End Sub


