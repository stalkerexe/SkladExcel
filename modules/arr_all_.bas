Attribute VB_Name = "arr_all_"
Option Explicit

Public Sub arr_zv()
        On Error Resume Next
        With ThisWorkbook.Sheets("Расход")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 2
            nn = Range(.Cells(rwZv, zvNN), .Cells(r7, zvNN)).value
            nm = Range(.Cells(rwZv, zvNm), .Cells(r7, zvNm)).value
            cod = Range(.Cells(rwZv, zvCod), .Cells(r7, zvCod)).value
            ed = Range(.Cells(rwZv, zvEd), .Cells(r7, zvEd)).value
            cnR = Range(.Cells(rwZv, zvCnR), .Cells(r7, zvCnR)).value
            cnZ = Range(.Cells(rwZv, zvCnZ), .Cells(r7, zvCnZ)).value
            cn = Range(.Cells(rwZv, zvCn), .Cells(r7, zvCn)).value
            col = Range(.Cells(rwZv, zvCol), .Cells(r7, zvCol)).value
            sm = Range(.Cells(rwZv, zvSm), .Cells(r7, zvSm)).value
            ost = Range(.Cells(rwZv, zvOst), .Cells(r7, zvOst)).value
            sk = Range(.Cells(rwZv, zvSk), .Cells(r7, zvSk)).value
            id = Range(.Cells(rwZv, 1), .Cells(r7, 1)).value
            iCol = Application.CountIf(Range(.Cells(rwZv, zvNm), .Cells(r7, zvNm)), "<>")
        End With
End Sub

Public Sub erase_arr_zv()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed
        Erase cnR: Erase cnZ: Erase cn
        Erase col: Erase sm: Erase ost: Erase sk
End Sub


Public Sub arr_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Приход")
            r7 = .Cells(Rows.Count, prNm).End(xlUp).Row + 2
            nn = Range(.Cells(rwZv, prNN), .Cells(r7, prNN)).value
            nm = Range(.Cells(rwZv, prNm), .Cells(r7, prNm)).value
            ed = Range(.Cells(rwZv, prEd), .Cells(r7, prEd)).value
            cod = Range(.Cells(rwZv, prCod), .Cells(r7, prCod)).value
            cnR = Range(.Cells(rwZv, prCnR), .Cells(r7, prCnR)).value
            cnZ = Range(.Cells(rwZv, prCnZ), .Cells(r7, prCnZ)).value
            col = Range(.Cells(rwZv, prCol), .Cells(r7, prCol)).value
            sm = Range(.Cells(rwZv, prSm), .Cells(r7, prSm)).value
            sk = Range(.Cells(rwZv, prSk), .Cells(r7, prSk)).value
            gr = Range(.Cells(rwZv, prGr), .Cells(r7, prGr)).value
            id = Range(.Cells(rwZv, 1), .Cells(r7, 1)).value
            iCol = Application.CountIf(Range(.Cells(rwZv, prNm), .Cells(r7, prNm)), "<>")
        End With
End Sub

' ИСПРАВЛЕНИЕ #1: добавлен Erase cn — в оригинале cn не очищался,
' из-за чего при повторном вызове в cn оставались данные предыдущей накладной
' и расчёт суммы шёл по старым ценам.
Public Sub erase_arr_pr()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed
        Erase cnR: Erase cnZ: Erase cn
        Erase col: Erase sm: Erase sk
End Sub

Public Sub erase_arr_sk()
        On Error Resume Next
        Erase gr_sk: Erase nm_sk: Erase ed_sk: Erase cod_sk
        Erase cnR_sk: Erase cnZ_sk: Erase col_sk
        Erase cr_sk: Erase br_sk: Erase comm_sk
End Sub


Public Sub arr_zk_this()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_расход")
            nn = Range(.Cells(row1, zkNN), .Cells(row2, zkNN)).value
            nm = Range(.Cells(row1, zkNm), .Cells(row2, zkNm)).value
            cod = Range(.Cells(row1, zkCod), .Cells(row2, zkCod)).value
            ed = Range(.Cells(row1, zkEd), .Cells(row2, zkEd)).value
            cnR = Range(.Cells(row1, zkCnR), .Cells(row2, zkCnR)).value
            cnZ = Range(.Cells(row1, zkCnZ), .Cells(row2, zkCnZ)).value
            cn = Range(.Cells(row1, zkCn), .Cells(row2, zkCn)).value
            col = Range(.Cells(row1, zkCol), .Cells(row2, zkCol)).value
            sm = Range(.Cells(row1, zkSm), .Cells(row2, zkSm)).value
            ost = Range(.Cells(row1, zkOst), .Cells(row2, zkOst)).value
            sk = Range(.Cells(row1, zkSk), .Cells(row2, zkSk)).value
            id = Range(.Cells(row1, zkID), .Cells(row2, zkID)).value
            iCol = Application.CountIf(Range(.Cells(row1, zkNm), .Cells(row2, zkNm)), "<>")
        End With
End Sub

Public Sub arr_zk_this_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets("Отложено_приход")
            nn = Range(.Cells(row1, pzkNN), .Cells(row2, pzkNN)).value
            nm = Range(.Cells(row1, pzkNm), .Cells(row2, pzkNm)).value
            cod = Range(.Cells(row1, pzkCod), .Cells(row2, pzkCod)).value
            ed = Range(.Cells(row1, pzkEd), .Cells(row2, pzkEd)).value
            cnR = Range(.Cells(row1, pzkCnR), .Cells(row2, pzkCnR)).value
            cnZ = Range(.Cells(row1, pzkCnZ), .Cells(row2, pzkCnZ)).value
            col = Range(.Cells(row1, pzkCol), .Cells(row2, pzkCol)).value
            sm = Range(.Cells(row1, pzkSm), .Cells(row2, pzkSm)).value
            ost = Range(.Cells(row1, pzkOst), .Cells(row2, pzkOst)).value
            sk = Range(.Cells(row1, pzkSk), .Cells(row2, pzkSk)).value
            gr = Range(.Cells(row1, pzkGr), .Cells(row2, pzkGr)).value
            id = Range(.Cells(row1, pzkID), .Cells(row2, pzkID)).value
            iCol = Application.CountIf(Range(.Cells(row1, pzkNm), .Cells(row2, pzkNm)), "<>")
        End With
End Sub

Public Sub erase_arr_zk_this()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed
        Erase cnR: Erase cnZ: Erase col: Erase sm: Erase ost: Erase sk
End Sub


' ИСПРАВЛЕНИЕ #2: arr_arh_pr теперь дополнительно считывает поле doc (документ основания),
' которое ранее отсутствовало — при формировании отчёта по приходу поле "документ" было пустым.
Public Sub arr_arh_rs()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            nn = Range(.Cells(row1, arhNN), .Cells(row2, arhNN)).value
            nm = Range(.Cells(row1, arhNm), .Cells(row2, arhNm)).value
            cod = Range(.Cells(row1, arhCod), .Cells(row2, arhCod)).value
            ed = Range(.Cells(row1, arhEd), .Cells(row2, arhEd)).value
            cnR = Range(.Cells(row1, arhCnR), .Cells(row2, arhCnR)).value
            cnZ = Range(.Cells(row1, arhCnZ), .Cells(row2, arhCnZ)).value
            col = Range(.Cells(row1, arhCol), .Cells(row2, arhCol)).value
            sm = Range(.Cells(row1, arhSm), .Cells(row2, arhSm)).value
            sk = Range(.Cells(row1, arhSk), .Cells(row2, arhSk)).value
            iCol = Application.CountIf(Range(.Cells(row1, arhNm), .Cells(row2, arhNm)), "<>")
        End With
End Sub

Public Sub arr_arh_pr()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            nn = Range(.Cells(row1, arhNN), .Cells(row2, arhNN)).value
            nm = Range(.Cells(row1, arhNm), .Cells(row2, arhNm)).value
            cod = Range(.Cells(row1, arhCod), .Cells(row2, arhCod)).value
            ed = Range(.Cells(row1, arhEd), .Cells(row2, arhEd)).value
            cnR = Range(.Cells(row1, arhCnR), .Cells(row2, arhCnR)).value
            cnZ = Range(.Cells(row1, arhCnZ), .Cells(row2, arhCnZ)).value
            col = Range(.Cells(row1, arhCol), .Cells(row2, arhCol)).value
            sm = Range(.Cells(row1, arhSm), .Cells(row2, arhSm)).value
            sk = Range(.Cells(row1, arhSk), .Cells(row2, arhSk)).value
            ' ИСПРАВЛЕНИЕ #2: считываем doc для отчёта по приходу
            doc = Range(.Cells(row1, arhDoc), .Cells(row2, arhDoc)).value
            iCol = Application.CountIf(Range(.Cells(row1, arhNm), .Cells(row2, arhNm)), "<>")
        End With
End Sub

Public Sub erase_arr_arh_this()
        On Error Resume Next
        Erase nn: Erase nm: Erase cod: Erase ed
        Erase cnR: Erase cnZ: Erase col: Erase sm: Erase ost: Erase sk
End Sub


Public Sub arr_arh_all()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r7 = .Cells(Rows.Count, arhNom).End(xlUp).Row + 2
            mk = Range(.Cells(2, 1), .Cells(r7, 1)).value
            nom = Range(.Cells(2, arhNom), .Cells(r7, arhNom)).value
            zkz = Range(.Cells(2, arhZkz), .Cells(r7, arhZkz)).value
            dt = Range(.Cells(2, arhDt), .Cells(r7, arhDt)).value
            mj = Range(.Cells(2, arhMj), .Cells(r7, arhMj)).value
            doc = Range(.Cells(2, arhDoc), .Cells(r7, arhDoc)).value
            iCol = Application.CountIf(Range(.Cells(2, 1), .Cells(r7, 1)), "<>")
        End With
End Sub


Public Sub arr_arh_for_otchet()
        On Error Resume Next
        With ThisWorkbook.Sheets(shNmArh)
            r7 = .Cells(Rows.Count, arhNm).End(xlUp).Row + 2

            nm = Range(.Cells(2, arhNm), .Cells(r7, arhNm)).value
            cod = Range(.Cells(2, arhCod), .Cells(r7, arhCod)).value
            ed = Range(.Cells(2, arhEd), .Cells(r7, arhEd)).value
            cnR = Range(.Cells(2, arhCnR), .Cells(r7, arhCnR)).value
            cnZ = Range(.Cells(2, arhCnZ), .Cells(r7, arhCnZ)).value
            col = Range(.Cells(2, arhCol), .Cells(r7, arhCol)).value
            sm = Range(.Cells(2, arhSm), .Cells(r7, arhSm)).value
            sk = Range(.Cells(2, arhSk), .Cells(r7, arhSk)).value

            mk = Range(.Cells(2, 1), .Cells(r7, 1)).value
            nom = Range(.Cells(2, arhNom), .Cells(r7, arhNom)).value
            zkz = Range(.Cells(2, arhZkz), .Cells(r7, arhZkz)).value
            dt = Range(.Cells(2, arhDt), .Cells(r7, arhDt)).value
            mj = Range(.Cells(2, arhMj), .Cells(r7, arhMj)).value
            doc = Range(.Cells(2, arhDoc), .Cells(r7, arhDoc)).value

            opl = Range(.Cells(2, arhOpl), .Cells(r7, arhOpl)).value
            skid = Range(.Cells(2, arhSkid), .Cells(r7, arhSkid)).value

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

