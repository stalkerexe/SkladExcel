VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DOP_spr 
   Caption         =   "Справочники"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   OleObjectBlob   =   "DOP_spr.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DOP_spr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const n = 6

Private Sub cm_1_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_supplier
End Sub

Private Sub cm_2_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_counterparty
End Sub

Private Sub cm_3_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_nomenclature
End Sub

Private Sub cm_4_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_units
End Sub

Private Sub cm_5_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_doc_types
End Sub

Private Sub cm_6_Click()
        On Error Resume Next
        Unload Me: DoEvents
        open_dict_warehouse
End Sub


Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
With ActiveSheet.Shapes("cmbt_2")
Me.Top = .Top
Me.Left = .Left + .Width
End With
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Me.Controls("cm_" & i).ForeColor = RGB(255, 255, 255)
Next
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
End Sub
Private Sub cm_1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_1.BackColor = RGB(128, 128, 128)
End Sub
Private Sub cm_2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_2.BackColor = RGB(128, 128, 128)
End Sub
Private Sub cm_3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_3.BackColor = RGB(128, 128, 128)
End Sub
Private Sub cm_4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_4.BackColor = RGB(128, 128, 128)
End Sub
Private Sub cm_5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_5.BackColor = RGB(128, 128, 128)
End Sub
Private Sub cm_6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
For i = 1 To n
Me.Controls("cm_" & i).BackColor = RGB(58, 110, 165)
Next
cm_6.BackColor = RGB(128, 128, 128)
End Sub


