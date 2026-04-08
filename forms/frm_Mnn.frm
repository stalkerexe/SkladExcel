VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mnn 
   ClientHeight    =   960
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   2760
   OleObjectBlob   =   "frm_Mnn.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Mnn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shName As String

Private Sub UserForm_Initialize()
        On Error Resume Next
        
        Call forma
        Call arr_this_sheets
        Call добавить_контролы_mnn
        
        For i = 2 To iCol
            Controls("nNm" & i).Top = Controls("nNm" & i - 1).Top + Controls("nNm" & i - 1).Height + 1
        Next
        
        iCol = iCol - 1
        Me.Height = hgCntrMnn * iCol + iCol * iZazor + 32
        Me.Width = widMnn + 22
        
End Sub

Private Sub forma()
        On Error Resume Next
        Me.StartUpPosition = 0
        With ThisWorkbook.ActiveSheet.Shapes("cmb_mn")
            Me.Top = .Top + .Height + 15
            Me.Left = .Left
        End With
End Sub


Private Sub arr_this_sheets()
        On Error Resume Next

        Dim sh As Worksheet
        
        iCol = 0
        For Each sh In ThisWorkbook.Sheets
            If sh.Visible = True Then
                iCol = iCol + 1
            End If
        Next
        
        ReDim nm(1 To iCol, 1 To 1)
        
        j = 1
        For Each sh In ThisWorkbook.Sheets
            If sh.Visible = True Then
                shName = sh.Name
                nm(j, 1) = shName
                j = j + 1
            End If
        Next

End Sub



Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call cmb_all
End Sub

Private Sub cmb_all()
        On Error Resume Next
        For Each ctr In frm_Mnn.Controls
            ctr.BackColor = frm_Mnn.BackColor
        Next
End Sub

