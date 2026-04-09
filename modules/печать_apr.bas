Attribute VB_Name = "печать_apr"
Option Explicit

Private Const shRow As Integer = 13
Private Const cmEnd = 8




Public Sub prnt_pr()
        Call doScreenOff
        Call do_blank
        Call doScreenOn
End Sub

Private Sub do_blank()
        Call dann_pr
        Call arr_pr
        Call copy_to_blank_pr
        Call print_paper
        Call clearBlank(shRow)
        ThisWorkbook.Sheets(nmBlank).Visible = 2
End Sub

Public Sub copy_to_blank_pr()
        nmBlank = "prntPr"
        ThisWorkbook.Sheets(nmBlank).Visible = True
        Call clearBlank(shRow)
        Call copy_nk
        Call format_
        Call copy_dann
        Call podp_all
        Call hidden_clm_blank_pr
End Sub

Private Sub copy_nk()
        On Error Resume Next
        With ThisWorkbook.Sheets(nmBlank)
            .Cells(shRow, prNm).Resize(UBound(nm), 1) = nm
            .Cells(shRow, prCod).Resize(UBound(nm), 1) = cod
            .Cells(shRow, prEd).Resize(UBound(nm), 1) = ed
            .Cells(shRow, prCol).Resize(UBound(nm), 1) = col
            .Cells(shRow, prCnZ).Resize(UBound(nm), 1) = cnZ
            .Cells(shRow, prSm).Resize(UBound(nm), 1) = sm
        End With
End Sub

Private Sub copy_dann()
        On Error Resume Next
        With ThisWorkbook.Sheets(nmBlank)
            .Range("c2").value = "Приходная накладная № " & nomer & " от " & sDt
            .Cells(rwPr_zkz, 4).value = sZkz
            .Cells(rwPr_mj, 4).value = sMj
            .Cells(rwPr_doc, 4).value = sOsn
            .Cells(rwPr_dt, 4).value = sDt
        End With
End Sub

Private Sub format_()
        On Error Resume Next

        With ThisWorkbook.Sheets(nmBlank)

            r9 = .Cells(Rows.Count, prNm).End(xlUp).Row
            
            Range(.Cells(rwZv, prNm), .Cells(r9, prSm)).Borders.LineStyle = True

            Range(.Cells(rwZv, prEd), .Cells(r9, prSm)).HorizontalAlignment = xlCenter

            Range(.Cells(rwZv, prCnZ), .Cells(r9, prCnZ)).NumberFormat = "#,##0.00"
            Range(.Cells(rwZv, prSm), .Cells(r9, prSm)).NumberFormat = "#,##0.00"

            With Range(.Cells(rwZv, prNN), .Cells(r9, prNN))
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With

            With Range(.Cells(rwZv, 2), .Cells(r9, prSm))
                .VerticalAlignment = xlTop
                .Font.Name = "Times New Roman"
                .Font.Size = 10
            End With

            Range(.Cells(rwZv, prCod), .Cells(r9, prCod)).IndentLevel = 1
            Range(.Cells(rwZv, prNm), .Cells(r9, prNm)).IndentLevel = 1

            With Range(.Cells(rwZv, prNm), .Cells(r9, prNm))
                .WrapText = True
                .Rows.AutoFit
            End With
            
            j = 1
            For i = rwZv To r9
                .Cells(i, prNm).RowHeight = .Cells(i, prNm).RowHeight + 3
.Cells(i, prNN) = j
                j = j + 1
            Next
            
        End With

        Call remove_green
        Range("a1").Select
        ActiveWindow.ScrollRow = 1

End Sub



Private Sub podp_all()
        Call podp_sm
        Call podp_podp
End Sub

Private Sub podp_sm()
        On Error Resume Next
        
        With ThisWorkbook.Sheets(nmBlank)

            r7 = .Cells(Rows.Count, prNm).End(xlUp).Row + 1

            .Cells(r7, prSm).RowHeight = 22

            .Cells(r7, prCnZ).value = "Итого:"

            .Cells(r7, prSm).value = summ
            .Cells(r7, prSm).NumberFormat = "#,##0.00"
            
            With Range(.Cells(r7, prNN), .Cells(r7, prSm))
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
            .Rows("3:5").Copy ThisWorkbook.Sheets(nmBlank).Rows(r7)
            Application.CutCopyMode = False
        End With
End Sub





