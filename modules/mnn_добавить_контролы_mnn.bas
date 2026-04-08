Attribute VB_Name = "mnn_добавить_контролы_mnn"
Option Explicit

Public Const hgCntrMnn As Double = 20
Public Const widMnn As Double = 120
Public Const iZazor As Double = 2


Public Sub добавить_контролы_mnn()
        On Error Resume Next

        iColCtr = iCol
        If iColCtr = 0 Then Exit Sub
        
        Call контролы_накладной
        Call arr_controls_mnn

End Sub

Private Sub контролы_накладной()
        On Error Resume Next
        
        iSize = 9

        ReDim arr_Ctr_zvk(1 To iColCtr)
        
        nom_Cnr = 1
        
        For i = LBound(nm) To UBound(nm)
            If nm(i, 1) <> "" Then
            
                sNm = nm(i, 1)
                Call controls_add
                
                nom_Cnr = nom_Cnr + 1
                
            End If
        Next
        
End Sub

Private Sub arr_controls_mnn()
        On Error Resume Next

        ReDim arr_Ctr_mnn(1 To nom_Cnr)

        For i = 1 To iColCtr
            Set arr_Ctr_mnn(i).nNm = frm_Mnn.Controls("nNm" & i)
        Next
End Sub

Private Sub controls_add()
        On Error Resume Next
        nmCntr = "nNm" & nom_Cnr
        txCntr = sNm
        leftCntr = 8
        gnCntr = fmTextAlignLeft
        Call add_cntrl_nk
End Sub


Private Sub add_cntrl_nk()
        On Error Resume Next
        With frm_Mnn.Controls.Add("Forms.CommandButton.1")
            .Name = nmCntr
            .Caption = sNm
            
If sNm = ActiveSheet.Name Then .Height = 0: GoTo 33
            .Height = hgCntrMnn
            
33
.Top = 0
            
            .BackColor = frm_Mnn.BackColor
            .BackStyle = 1
            .Cancel = False
            .Default = False
            .TabStop = True
            
            .Left = leftCntr
            .Width = widMnn
            .Font.Size = 9
            .Font.Name = "Times New Roman"
            
        End With
End Sub













