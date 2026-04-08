VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Set 
   Caption         =   "Настройки программы"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   OleObjectBlob   =   "frm_Set.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBox_Cn_Click()
        If CheckBox_Cn.Value = False Then
            CheckBox_skid.Value = False: ThisWorkbook.Sheets("setting").Range("b42").Value = 0
            CheckBox_opl.Value = False:  ThisWorkbook.Sheets("setting").Range("b43").Value = 0
            Frame_opl.Visible = False
        Else
            Frame_opl.Visible = True
        End If
End Sub



Private Sub OK_nk_pr_Click()
        With ThisWorkbook.Sheets("setting")
            
            If CheckBox_doc.Value = True Then
                flag_hidden = False
                .Range("b35").Value = 1
            Else
                flag_hidden = True
                .Range("b35").Value = 0
            End If
        End With
        Call clmns_hidden
End Sub



Private Sub OK_nk_rs_Click()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("setting")
        
            If CheckBox_adr.Value = True Then
                flag_hidden = False
                .Range("b40").Value = 1
            Else
                flag_hidden = True
                .Range("b40").Value = 0
            End If
            
            If CheckBox_tlf.Value = True Then
                flag_hidden = False
                .Range("b41").Value = 1
            Else
                flag_hidden = True
                .Range("b41").Value = 0
            End If
            
            
        End With
        
        Call clmns_hidden
End Sub


Private Sub OK2_Click()
On Error Resume Next
With ThisWorkbook.Sheets("setting")
If CheckBox_Cn.Value = True Then
flag_hidden = False
.Range("b8").Value = 1
Else
flag_hidden = True
.Range("b8").Value = 0
End If
If CheckBox_Cod.Value = True Then
flag_hidden = False
.Range("b6").Value = 1
Else
flag_hidden = True
.Range("b6").Value = 0
End If
If CheckBox_Br.Value = True Then
flag_hidden = False
.Range("b9").Value = 1
Else
flag_hidden = True
.Range("b9").Value = 0
End If
If CheckBox_Cr.Value = True Then
flag_hidden = False
.Range("b11").Value = 1
Else
flag_hidden = True
.Range("b11").Value = 0
End If
If CheckBox_0.Value = True Then
flag_hidden = False
.Range("p4").Value = 1
Else
flag_hidden = True
.Range("p4").Value = 0
End If
End With

Call hidden_opl_skid

Call clmns_hidden

End Sub

Private Sub hidden_opl_skid()

        With ThisWorkbook.Sheets("setting")
        
            If CheckBox_opl.Value = True Then
                flag_hidden = False
                .Range("b42").Value = 1
            Else
                flag_hidden = True
                .Range("b42").Value = 0
            End If
            
            If CheckBox_skid.Value = True Then
                flag_hidden = False
                .Range("b43").Value = 1
            Else
                flag_hidden = True
                .Range("b43").Value = 0
            End If
                
        End With
End Sub

Private Sub OK_sk_Click()
On Error Resume Next
With ThisWorkbook.Sheets("setting")
If CheckBox_Cn.Value = True Then
flag_hidden = False
.Range("b8").Value = 1
Else
flag_hidden = True
.Range("b8").Value = 0
End If
If CheckBox_Cod.Value = True Then
flag_hidden = False
.Range("b6").Value = 1
Else
flag_hidden = True
.Range("b6").Value = 0
End If
If CheckBox_Br.Value = True Then
flag_hidden = False
.Range("b9").Value = 1
Else
flag_hidden = True
.Range("b9").Value = 0
End If
If CheckBox_Cr.Value = True Then
flag_hidden = False
.Range("b11").Value = 1
Else
flag_hidden = True
.Range("b11").Value = 0
End If
If CheckBox_0.Value = True Then
flag_hidden = False
.Range("p4").Value = 1
Else
flag_hidden = True
.Range("p4").Value = 0
End If

End With
Call clmns_hidden
End Sub




Private Sub CheckBox1_Change()
With Sheets("setting")
If CheckBox1.Value = True Then
.Range("i4") = 1
Else
.Range("i4") = 0
End If
End With
End Sub



Private Sub UserForm_Initialize()
        On Error Resume Next
        With ThisWorkbook.Sheets("setting")
            If .Range("b6") = 1 Then CheckBox_Cod.Value = True
            If .Range("b8") = 1 Then CheckBox_Cn.Value = True
            If .Range("b9") = 1 Then CheckBox_Br.Value = True
            If .Range("b11") = 1 Then CheckBox_Cr.Value = True
            If .Range("p4") = 1 Then CheckBox_0.Value = True
            If .Range("b35") = 1 Then CheckBox_doc.Value = True
            If .Range("b36") = 1 Then CheckBox_print_ot.Value = True
            If .Range("b37") = 1 Then CheckBox_print_pr.Value = True
            If .Range("b40") = 1 Then CheckBox_adr.Value = True
            If .Range("b41") = 1 Then CheckBox_tlf.Value = True
            If .Range("b42") = 1 Then CheckBox_opl.Value = True
            If .Range("b43") = 1 Then CheckBox_skid.Value = True
            If .Range("i4") = 1 Then CheckBox1.Value = True
        End With
End Sub

