Attribute VB_Name = "печать_________________________"
Option Explicit

Public nmBlank As String
Public iCol_Print As Integer



Public Function clearBlank(rw As Integer)
        On Error Resume Next
        With ThisWorkbook.Sheets(nmBlank)
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
            .Range(.Cells(rw, 2), .Cells(r7 + 44, 2)).EntireRow.Delete
        End With
End Function




Public Sub hidden_clm_blank_rs()
        On Error Resume Next

        If ThisWorkbook.Sheets("setting").Range("b6").value = 1 Then
            flag_hidden = False
        Else
            flag_hidden = True
        End If
        
        ThisWorkbook.Sheets(nmBlank).Cells(2, zvCod).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b8").value = 1 Then
            flag_hidden = False
        Else
            flag_hidden = True
        End If
        
        With ThisWorkbook.Sheets(nmBlank)
            Range(.Cells(2, zvCnR), .Cells(2, zvSm)).EntireColumn.Hidden = flag_hidden
        End With


        If ThisWorkbook.Sheets("setting").Range("b40") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets(nmBlank).Cells(rwZv_adr, 2).EntireRow.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b41") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        ThisWorkbook.Sheets(nmBlank).Cells(rwZv_tlf, 2).EntireRow.Hidden = flag_hidden

End Sub


Public Sub hidden_clm_blank_pr()
        On Error Resume Next

        If ThisWorkbook.Sheets("setting").Range("b6").value = 1 Then
            flag_hidden = False
        Else
            flag_hidden = True
        End If
        
        ThisWorkbook.Sheets(nmBlank).Cells(2, prCod).EntireColumn.Hidden = flag_hidden
        
        If ThisWorkbook.Sheets("setting").Range("b8").value = 1 Then
            flag_hidden = False
        Else
            flag_hidden = True
        End If
        
        With ThisWorkbook.Sheets(nmBlank)
            .Cells(2, prCnZ).EntireColumn.Hidden = flag_hidden
            .Cells(2, prSm).EntireColumn.Hidden = flag_hidden
        End With

        With ThisWorkbook.Sheets(nmBlank)
            .Cells(2, prCnR).EntireColumn.Hidden = True
        End With

        If ThisWorkbook.Sheets("setting").Range("b35") = 1 Then
        flag_hidden = False
        Else
        flag_hidden = True
        End If
        
        ThisWorkbook.Sheets(nmBlank).Cells(rwPr_doc, 2).EntireRow.Hidden = flag_hidden
        
End Sub




Public Sub print_paper()

        cm = 3

        With ThisWorkbook.Sheets(nmBlank)
            r7 = .Cells(Rows.Count, cm).End(xlUp).Row
            With .PageSetup
                .PrintTitleRows = "$12:$12"
                .PrintArea = "b1:i" & r7
            End With
            .PrintOut Copies:=iCol_Print, ActivePrinter:=sPrinter
        End With

End Sub
