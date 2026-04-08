Attribute VB_Name = "печать_azv"
Option Explicit

Private Const shRow As Integer = 13
Private Const cmEnd = 8




Public Sub prnt_zv()
        Call doScreenOff
        Call do_blank
        Call doScreenOn
End Sub

Private Sub do_blank()
        Call dann_zv
        Call arr_zv
        Call copy_to_blank_rs
        Call print_paper
        Call clearBlank(shRow)
        ThisWorkbook.Sheets(nmBlank).Visible = 2
End Sub

Public Sub copy_to_blank_rs()
        nmBlank = "prntZv"
        ThisWorkbook.Sheets(nmBlank).Visible = True
        Call clearBlank(shRow)
        Call copy_nk
        Call format_
        Call copy_dann
        Call podp_all
        Call hidden_clm_blank_rs
End Sub

Private Sub copy_nk()
        On Error Resume Next
        With ThisWorkbook.Sheets(nmBlank)
            .Cells(shRow, zvNN).Resize(UBound(nm), 1) = nn
            .Cells(shRow, zvNm).Resize(UBound(nm), 1) = nm
            .Cells(shRow, zvCod).Resize(UBound(nm), 1) = cod
            .Cells(shRow, zvEd).Resize(UBound(nm), 1) = ed
            .Cells(shRow, zvCol).Resize(UBound(nm), 1) = col
            .Cells(shRow, zvCnR).Resize(UBound(nm), 1) = cnR
            .Cells(shRow, zvSm).Resize(UBound(nm), 1) = sm
        End With
End Sub

Private Sub copy_dann()
        On Error Resume Next
        With ThisWorkbook.Sheets(nmBlank)
            .Range("c2").Value = "Расходная накладная № " & nomer & " от " & sDt
            .Cells(rwZv_zkz, 4).Value = sZkz
            .Cells(rwZv_adr, 4).Value = sAdr
            .Cells(rwZv_tlf, 4).Value = sTlf
            .Cells(rwZv_mj, 4).Value = sMj
            .Cells(rwZv_dt, 4).Value = sDt
        End With
End Sub

Private Sub format_()
        On Error Resume Next

        With ThisWorkbook.Sheets(nmBlank)

            r9 = .Cells(Rows.Count, zvNm).End(xlUp).Row

            Range(.Cells(rwZv, zvNm), .Cells(r9, zvSm)).Borders.LineStyle = True

            Range(.Cells(rwZv, zvEd), .Cells(r9, zvSm)).HorizontalAlignment = xlCenter

            Range(.Cells(rwZv, zvCnR), .Cells(r9, zvCnR)).NumberFormat = "#,##0.00"
            Range(.Cells(rwZv, zvSm), .Cells(r9, zvSm)).NumberFormat = "#,##0.00"

            With Range(.Cells(rwZv, zvNN), .Cells(r9, zvNN))
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With

            With Range(.Cells(rwZv, 2), .Cells(r9, zvSm))
                .VerticalAlignment = xlTop
                .Font.Name = "Times New Roman"
                .Font.Size = 10
            End With

            Range(.Cells(rwZv, zvCod), .Cells(r9, zvCod)).IndentLevel = 1
            Range(.Cells(rwZv, zvNm), .Cells(r9, zvNm)).IndentLevel = 1

            With Range(.Cells(rwZv, zvNm), .Cells(r9, zvNm))
                .WrapText = True
                .Rows.AutoFit
            End With

            For i = rwZv To r9
                .Cells(i, zvNm).RowHeight = .Cells(i, zvNm).RowHeight + 3
            Next

        End With

        Call remove_green
        Range("a1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1

End Sub



Private Sub podp_all()
        Call podp_sm
        Call podp_podp
End Sub

Private Sub podp_sm()
        On Error Resume Next
        
        With ThisWorkbook.Sheets(nmBlank)

            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 1

            .Cells(r7, zvSm).RowHeight = 22

            .Cells(r7, zvCnR).Value = "Итого:"

            .Cells(r7, zvSm).Value = summ
            .Cells(r7, zvSm).NumberFormat = "#,##0.00"
            
            With Range(.Cells(r7, zvNN), .Cells(r7, zvSm))
                .Font.Name = "Times New Roman"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 11
            End With
            
        End With

End Sub

Private Sub podp_podp()
        On Error Resume Next
        r7 = r7 + 2
        With ThisWorkbook.Sheets("podp")
            .Rows("9:16").Copy ThisWorkbook.Sheets(nmBlank).Rows(r7)
            Application.CutCopyMode = False
        End With
End Sub






