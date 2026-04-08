Attribute VB_Name = "архив__________________________"
Option Explicit

Public widCod As Byte
Public widCn As Byte
Public iVid As String




Public Sub hidden_clm_zvk()
        On Error Resume Next
        
        Call clm_setting
        
        With frm_ZVK
        
            If widCod = 0 Then
                .lb_cod.Visible = False
                .lb_nm.Width = .lb_nm.Width + .lb_cod.Width + 2
            End If
            
            If widCn = 0 Then
                .lb_cn.Visible = False
                .lb_sm.Visible = False
                .tb_sm.Visible = False
                .lb_summ.Visible = False
            End If
            
        End With
        
End Sub

Private Sub clm_setting()
        widCod = 1: widCn = 1
        If ThisWorkbook.Sheets("setting").Range("b6").Value = 0 Then widCod = 0
        If ThisWorkbook.Sheets("setting").Range("b8").Value = 0 Then widCn = 0
End Sub
