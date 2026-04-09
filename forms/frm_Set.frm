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
        If CheckBox_Cn.value = False Then
            CheckBox_skid.value = False: ThisWorkbook.Sheets("setting").Range("b42").value = 0
            CheckBox_opl.value = False:  ThisWorkbook.Sheets("setting").Range("b43").value = 0
            Frame_opl.Visible = False
        Else
            Frame_opl.Visible = True
        End If
End Sub



Private Sub OK_nk_pr_Click()
        With ThisWorkbook.Sheets("setting")
            
            If CheckBox_doc.value = True Then
                flag_hidden = False
                .Range("b35").value = 1
            Else
                flag_hidden = True
                .Range("b35").value = 0
            End If
        End With
        Call clmns_hidden
End Sub



Private Sub OK_nk_rs_Click()
        On Error Resume Next
        
        With ThisWorkbook.Sheets("setting")
        
            If CheckBox_adr.value = True Then
                flag_hidden = False
                .Range("b40").value = 1
            Else
                flag_hidden = True
                .Range("b40").value = 0
            End If
            
            If CheckBox_tlf.value = True Then
                flag_hidden = False
                .Range("b41").value = 1
            Else
                flag_hidden = True
                .Range("b41").value = 0
            End If
            
            
        End With
        
        Call clmns_hidden
End Sub


Private Sub OK2_Click()
On Error Resume Next
With ThisWorkbook.Sheets("setting")
If CheckBox_Cn.value = True Then
flag_hidden = False
.Range("b8").value = 1
Else
flag_hidden = True
.Range("b8").value = 0
End If
If CheckBox_Cod.value = True Then
flag_hidden = False
.Range("b6").value = 1
Else
flag_hidden = True
.Range("b6").value = 0
End If
If CheckBox_Br.value = True Then
flag_hidden = False
.Range("b9").value = 1
Else
flag_hidden = True
.Range("b9").value = 0
End If
If CheckBox_Cr.value = True Then
flag_hidden = False
.Range("b11").value = 1
Else
flag_hidden = True
.Range("b11").value = 0
End If
If CheckBox_0.value = True Then
flag_hidden = False
.Range("p4").value = 1
Else
flag_hidden = True
.Range("p4").value = 0
End If
End With

Call hidden_opl_skid

Call clmns_hidden

End Sub

Private Sub hidden_opl_skid()

        With ThisWorkbook.Sheets("setting")
        
            If CheckBox_opl.value = True Then
                flag_hidden = False
                .Range("b42").value = 1
            Else
                flag_hidden = True
                .Range("b42").value = 0
            End If
            
            If CheckBox_skid.value = True Then
                flag_hidden = False
                .Range("b43").value = 1
            Else
                flag_hidden = True
                .Range("b43").value = 0
            End If
                
        End With
End Sub

Private Sub OK_sk_Click()
On Error Resume Next
With ThisWorkbook.Sheets("setting")
If CheckBox_Cn.value = True Then
flag_hidden = False
.Range("b8").value = 1
Else
flag_hidden = True
.Range("b8").value = 0
End If
If CheckBox_Cod.value = True Then
flag_hidden = False
.Range("b6").value = 1
Else
flag_hidden = True
.Range("b6").value = 0
End If
If CheckBox_Br.value = True Then
flag_hidden = False
.Range("b9").value = 1
Else
flag_hidden = True
.Range("b9").value = 0
End If
If CheckBox_Cr.value = True Then
flag_hidden = False
.Range("b11").value = 1
Else
flag_hidden = True
.Range("b11").value = 0
End If
If CheckBox_0.value = True Then
flag_hidden = False
.Range("p4").value = 1
Else
flag_hidden = True
.Range("p4").value = 0
End If

End With
Call clmns_hidden
End Sub




Private Sub CheckBox1_Change()
With Sheets("setting")
If CheckBox1.value = True Then
.Range("i4") = 1
Else
.Range("i4") = 0
End If
End With
End Sub



Private Sub UserForm_Initialize()
        On Error Resume Next
        With ThisWorkbook.Sheets("setting")
            If .Range("b6") = 1 Then CheckBox_Cod.value = True
            If .Range("b8") = 1 Then CheckBox_Cn.value = True
            If .Range("b9") = 1 Then CheckBox_Br.value = True
            If .Range("b11") = 1 Then CheckBox_Cr.value = True
            If .Range("p4") = 1 Then CheckBox_0.value = True
            If .Range("b35") = 1 Then CheckBox_doc.value = True
            If .Range("b36") = 1 Then CheckBox_print_ot.value = True
            If .Range("b37") = 1 Then CheckBox_print_pr.value = True
            If .Range("b40") = 1 Then CheckBox_adr.value = True
            If .Range("b41") = 1 Then CheckBox_tlf.value = True
            If .Range("b42") = 1 Then CheckBox_opl.value = True
            If .Range("b43") = 1 Then CheckBox_skid.value = True
            If .Range("i4") = 1 Then CheckBox1.value = True
        End With
End Sub

