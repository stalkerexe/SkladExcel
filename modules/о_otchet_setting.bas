Attribute VB_Name = "о_otchet_setting"
Option Explicit

Public Sub set_otchet()
        On Error Resume Next

        With wbOt.ActiveSheet
        
            If ThisWorkbook.Sheets("setting").Range("b6").Value = 1 Then
                flag_hidden = False
            Else
                flag_hidden = True
            End If
            .Range("e2").EntireColumn.Hidden = flag_hidden

            If ThisWorkbook.Sheets("setting").Range("b8").Value = 1 Then
                flag_hidden = False
            Else
                flag_hidden = True
            End If
            
            If iVid = "pr" Then .Range("g2:h2").EntireColumn.Hidden = flag_hidden
            If iVid = "ot" Then .Range("g2:k2").EntireColumn.Hidden = flag_hidden

        End With

End Sub




