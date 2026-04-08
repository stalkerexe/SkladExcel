VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Comm 
   Caption         =   "Примечание"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   OleObjectBlob   =   "frm_Comm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRow_mk As Long
Dim cmNm As Integer


Private Sub OK_Click()
        Call ent_comm
End Sub


Private Sub ent_comm()

        iOperation = tb_sheet.Text

        If iOperation = "pr" Then
            sheetNm = "Отложено_приход"
            cm = pzkComm
            cmNm = pzkNm
            Sheets(sheetNm).Select
            iFiles = "zkz_prihod.xlsx"
        End If
        
        If iOperation = "rs" Then
            sheetNm = "Отложено_расход"
            cm = zkComm
            cmNm = zkNm
            Sheets(sheetNm).Select
            iFiles = "zkz_rashodd.xlsx"
        End If
        
        Call dann
        Call add_comm_hear
        
        Unload Me
        
End Sub

Private Sub dann()
        On Error Resume Next
        With Me
            marker = .tb_mk.Text
            sComm = .tb_comm.Text
            iRow = .tb_row.Value
        End With
End Sub

Private Sub add_comm_hear()
        On Error Resume Next
        
        iHg = Cells(iRow, zkNm).RowHeight

        Cells(iRow, cm) = sComm
        Cells(iRow - 1, cm).Select
        
        Call find_row2
        iCol = Application.CountIf(Range(Cells(iRow, cmNm), Cells(row2, cmNm)), "<>")

        If iCol = 1 Then
            Cells(iRow, cm).RowHeight = iHg
        End If

End Sub

Private Sub find_row2()
        On Error Resume Next
        
        With ThisWorkbook.Sheets(sheetNm)
            r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
        
            row2 = 0
            For i = iRow + 1 To r7
                If .Cells(i, 1) <> "" Then
                    row2 = i - 1
                    GoTo 99
                End If
            Next
        
        End With
        
        row2 = r7
99

End Sub





Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
Me.Top = 200
If ThisWorkbook.ActiveSheet.Name = "Отложено_приход" Then
Me.Left = Cells(2, pzkComm).Left
Else
Me.Left = Cells(2, zkComm).Left
End If
End Sub

Private Sub NO_Click()
Unload Me
End Sub


Private Sub tb_comm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If KeyCode = 13 Then
            KeyCode = 0
            tb_comm.Text = tb_comm.Text & vbCrLf
        End If
End Sub
