Attribute VB_Name = "vz_перенос_текста"
Option Explicit
Public iControl As Control
Public widControl As Double
Public iHeight As Double
Public smHeight As Double
Public iTop As Double
Public iTop0 As Double

Private Const hgCntr = 18
Dim iColZnakov As Integer
Dim iHeightMax As Double


Private Const iCm As Integer = 8




Public Sub perenos_yes()

        Call col_Controls
        
        Call perenos_text

        frm_ZVK.Repaint

End Sub

Private Sub perenos_text()
        Call perenos_MultiLine_yes
        Call perenos_Height_yes
        Call perenos_Top_yes
        Call form_height
End Sub


Private Sub perenos_Height_yes()
        On Error Resume Next
        
        smHeight = 0
        
        For i = 1 To iColCtr
        
            ReDim c(1 To iCm, 1 To 1)
            
            iTop = frm_ZVK.Frame_nk.Controls("nNm" & i).Top
            Call fimd_max_height
            
            iHeightMax = Application.Max(c)
            If iHeightMax <= hgCntr Then iHeightMax = hgCntr
            
            
            Call высота_в_ряду
            
smHeight = smHeight + iHeightMax
            
        Next
        
End Sub

Private Sub fimd_max_height()
        On Error Resume Next

        j = 1

        For Each ctr In frm_ZVK.Frame_nk.Controls
            If ctr.Top = iTop Then
                iHeight = ctr.Height
                c(j, 1) = iHeight
                j = j + 1
            End If
        Next
        
End Sub

Private Sub высота_в_ряду()
        On Error Resume Next

        For Each ctr In frm_ZVK.Frame_nk.Controls
            If ctr.Top = iTop Then
                ctr.Height = iHeightMax
            End If
        Next
        
        For Each ctr In frm_ZVK.Frame_nk_vz.Controls
            If ctr.Name = "omb_sk" Then GoTo 99
            If ctr.Name = "SpinButton" Then GoTo 99
            
            If ctr.Top = iTop Then
                ctr.Height = iHeightMax
            End If
99
        Next
        
End Sub





Private Sub perenos_Top_yes()
        On Error Resume Next
        
        For i = 2 To iColCtr
        
            iTop0 = frm_ZVK.Controls("nNn" & i).Top
            iTop = frm_ZVK.Controls("nNm" & i - 1).Top + frm_ZVK.Controls("nNm" & i - 1).Height - 1
            
Call top_controls_row
        Next
        
End Sub

Private Sub top_controls_row()
        On Error Resume Next
        
        For Each ctr In frm_ZVK.Frame_nk.Controls
            If ctr.Top = iTop0 Then
                ctr.Top = iTop
            End If
        Next

        For Each ctr In frm_ZVK.Frame_nk_vz.Controls
            If ctr.Name = "omb_sk" Then GoTo 99
            If ctr.Name = "SpinButton" Then GoTo 99

            If ctr.Top = iTop0 Then
                ctr.Top = iTop
            End If
99
        Next

End Sub
















Private Sub perenos_MultiLine_yes()
        On Error Resume Next
        For Each ctr In frm_ZVK.Frame_nk.Controls
        
            If VBA.Left(ctr.Name, 3) = "nNm" Then

                widControl = ctr.Width
        
                ctr.MultiLine = True
                ctr.AutoSize = True
                ctr.AutoSize = False
                ctr.Width = widControl
            
            End If
        
        Next
End Sub








Private Sub form_height()
        On Error Resume Next

        With frm_ZVK

            .Frame_nk.Height = smHeight + iColCtr
            
            If .Frame_nk.Height > .Frame_nk_all.Height Then
                .ScrollBar1.Width = 12
            Else
                .ScrollBar1.Width = 0
            End If
            
            .ScrollBar1.Min = 0
            .ScrollBar1.Max = Val(.Frame_nk.Height - .Frame_nk_all.Height)

            .Frame_nk_vz.Height = .Frame_nk.Height

        End With

End Sub















Public Sub perenos_no()

        Call col_Controls
        
        Call perenos_text_no

        frm_ZVK.Repaint

End Sub

Private Sub perenos_text_no()
        Call perenos_MultiLine_no
        Call Height_controls_row_no
        Call perenos_Top_no
        Call form_height
End Sub

Private Sub perenos_MultiLine_no()
        On Error Resume Next
        For Each ctr In frm_ZVK.Frame_nk.Controls
            ctr.MultiLine = False
            ctr.AutoSize = False
        Next
End Sub

Private Sub perenos_Height_no()
        Call Height_controls_row_no
        smHeight = iColCtr * (hgCntr - 1)
End Sub

Private Sub Height_controls_row_no()
        On Error Resume Next

        For Each ctr In frm_ZVK.Frame_nk.Controls
            ctr.Height = hgCntr
        Next

        For Each ctr In frm_ZVK.Frame_nk_vz.Controls
            If ctr.Name = "omb_sk" Then GoTo 99
            If ctr.Name = "SpinButton" Then GoTo 99
            ctr.Height = hgCntr
99
        Next

End Sub



Private Sub perenos_Top_no()
        On Error Resume Next
        
        For i = 2 To iColCtr
        
            iTop0 = frm_ZVK.Controls("nNn" & i).Top
            iTop = frm_ZVK.Controls("nNn" & i - 1).Top + frm_ZVK.Controls("nNn" & i - 1).Height - 1
            frm_ZVK.Controls("nNn" & i).Top = iTop
            
Call top_controls_row_no
        Next
        
End Sub

Private Sub top_controls_row_no()
        On Error Resume Next

        For Each ctr In frm_ZVK.Frame_nk.Controls
            If ctr.Top = iTop0 Then
                ctr.Top = iTop
            End If
        Next

        For Each ctr In frm_ZVK.Frame_nk_vz.Controls
            If ctr.Name = "omb_sk" Then GoTo 99
            If ctr.Name = "SpinButton" Then GoTo 99

            If ctr.Top = iTop0 Then
                ctr.Top = iTop
            End If
99
        Next

End Sub






