Attribute VB_Name = "ф_____________________________"
Option Explicit


Public Function режим_редактирования_on_pr(shNm As String)
        With ThisWorkbook.Sheets(shNm)
            .Cells(9, zvOst) = "Режим_редактирования"
.Range("d1").Value = .Range("d2").Value
        End With
End Function

Public Function режим_редактирования_off_pr(shNm As String)
        With ThisWorkbook.Sheets(shNm)
            If .Cells(9, zvOst) = "Режим_редактирования" Then
                .Cells(9, zvOst) = ""
.Range("d2").Value = .Range("d1").Value
                .Range("d1") = ""
            End If
        End With
End Function


Public Sub find_path_vid()
        If iVid = "Приход" Then shNmArh = "arh_prr"
        If iVid = "Отгрузка" Then shNmArh = "arh_zkk"
        If iVid = "Возврат" Then shNmArh = "arh_vzz"
End Sub

Public Sub dann_zvk()
        With frm_ZVK
            marker = .tb_mk.Text
            nomer = .tb_nomer.Caption
            iVid = .tb_what.Text
            iGod = .tb_year.Text
            iPapka = iVid
ind = .tb_ind.Text
        End With
End Sub



Public Sub zakaz_prog()
Form_.Show
Form_.MultiPage1.Value = 1
End Sub
