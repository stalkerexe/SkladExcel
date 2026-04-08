VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DOP_ot 
   Caption         =   "Отчеты"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   OleObjectBlob   =   "DOP_ot.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DOP_ot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const n = 2
        
Private Sub cm_1_Click()
Unload Me
DoEvents

iVid = "pr"
shNmArh = "arh_prr"

Call arr_arh_proverka
If iCol = 0 Then
    MsgBox "Пока нет накладных для просмотра отчета." _
    & Chr(10) & "Сначала оприходуйте или отгрузите накладные", 64, "Отчет"
    Exit Sub
End If


frm_Ot_msg.Show
End Sub

Private Sub cm_2_Click()
Unload Me
DoEvents
iVid = "ot"
shNmArh = "arh_zkk"

Call arr_arh_proverka
If iCol = 0 Then
    MsgBox "Пока нет накладных для просмотра отчета." _
    & Chr(10) & "Сначала оприходуйте или отгрузите накладные", 64, "Отчет"
    Exit Sub
End If

frm_Ot_msg.Show
End Sub



Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
If ThisWorkbook.ActiveSheet.Name = "Главная" Then
With ActiveSheet.Shapes("cmbt_4")
Me.Top = .Top
Me.Left = .Left + .Width
End With
Else
With ActiveSheet.Shapes("cmb_mn")
Me.Top = .Top + .Height + 15
Me.Left = .Left
End With
End If
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
