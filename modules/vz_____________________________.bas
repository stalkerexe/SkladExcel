Attribute VB_Name = "vz_____________________________"
Option Explicit

Public iColCtr As Integer

Public Const frmHg2 As Double = 376

Public Const frmWd1 As Double = 560
Public Const frmWd2 As Double = 700

Public iMk As String
Public sNN As String

Public sCol_vz As Double


Public Sub col_Controls()
        On Error Resume Next
        iColCtr = 0
        For Each iCtr In frm_ZVK.Frame_nk.Controls
            If VBA.Left(iCtr.Name, 3) = "nNm" Then
            iColCtr = iColCtr + 1
            End If
        Next
End Sub

Public Sub proverka_vz()

        Call col_Controls
        
        iCol = 0
        
        With frm_ZVK
            For i = 1 To iColCtr
            
                str_ = .Controls("nCol_vz" & i).Value
                
                If .Controls("nCol_vz" & i).Value <> "" Then
                    If .Controls("nCol_vz" & i).Value <> 0 Then
                        iCol = iCol + 1
                    End If
                End If
            Next
        End With

End Sub


Public Sub clear_color_vz()
        On Error Resume Next
        
        For Each iCtr In frm_ZVK.Frame_nk.Controls
            iCtr.BackColor = &H80000005
33
        Next
        
        For Each iCtr In frm_ZVK.Frame_nk_vz.Controls
            If iCtr.Name = "comb_sk" Then GoTo 99
            If iCtr.Name = "SpinButton" Then GoTo 99
            iCtr.BackColor = &H80000005
99
        Next
        
End Sub


