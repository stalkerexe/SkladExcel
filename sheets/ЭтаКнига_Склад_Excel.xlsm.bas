' Component: ЭтаКнига  [Склад Excel.xlsm]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit


Private Sub Workbook_BeforeClose(Cancel As Boolean)
On Error Resume Next
Call verekr
ThisWorkbook.Save
End Sub



Private Sub Workbook_Open()
        On Error Resume Next
        
        Call polnekr
        
        With ThisWorkbook.Sheets("Расход")
            .Cells(rwZv_dt, 4).Value = VBA.Date
            .Cells(rwZv_dt2, 4).Value = VBA.Date
        End With
        
        With ThisWorkbook.Sheets("Приход")
            .Cells(rwPr_dt, 4).Value = VBA.Date
        End With
        
        Sheets("Главная").Select
End Sub

