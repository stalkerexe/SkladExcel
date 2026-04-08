VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_mnn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents nNm As MSForms.CommandButton
Attribute nNm.VB_VarHelpID = -1


Private Sub nNm_Click()
        On Error Resume Next
        sNm = nNm.Caption
        Sheets(sNm).Select
        Unload frm_Mnn
End Sub



Private Sub nNm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        On Error Resume Next
        Call cmb_all
        nNm.BackColor = &H80000005
End Sub

Private Sub cmb_all()
        On Error Resume Next
        For Each ctr In frm_Mnn.Controls
            ctr.BackColor = frm_Mnn.BackColor
        Next
End Sub


