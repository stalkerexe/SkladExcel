VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DOP_sv 
   Caption         =   "Отложенные накладные"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3675
   OleObjectBlob   =   "DOP_sv.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DOP_sv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const n = 2
Private Sub cm_1_Click()
Unload Me
DoEvents
Sheets("Отложено_расход").Select
End Sub
Private Sub cm_2_Click()
Unload Me
DoEvents
Sheets("Отложено_приход").Select
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
With ActiveSheet.Shapes("cmbt_7")
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


