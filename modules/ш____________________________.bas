Attribute VB_Name = "ш____________________________"


Public Function sort_do(cm1 As Integer, cm2 As Integer)
        On Error Resume Next
        With ThisWorkbook.Sheets("буфер")
            r7 = .Cells(Rows.Count, cm1).End(xlUp).Row
            .Sort.SortFields.clear
            .Sort.SortFields.Add Key:=.Cells(1, cm1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With .Sort
                .SetRange Range(Cells(1, cm1), Cells(r7, cm2))
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
End Function



